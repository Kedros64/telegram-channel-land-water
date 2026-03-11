#!/usr/bin/env python3
"""
regen.py — Повторная генерация поста из кэша статей без парсинга источников.

Используется для быстрой настройки промпта и визуала:
  1. Запустите monitor.py хотя бы раз — создаётся posts/articles_cache.json
  2. Правьте DEEPSEEK_POST_PROMPT или другие настройки в monitor.py
  3. Запускайте regen.py — мгновенная регенерация без 3–5 минут парсинга

Использование:
    python scripts/regen.py                      # DeepSeek выберет статью, покажет пост
    python scripts/regen.py --list               # показать список статей в кэше
    python scripts/regen.py --index 3            # взять статью с индексом 3
    python scripts/regen.py --index 1,4,7        # несколько статей через запятую
    python scripts/regen.py --num 3              # сгенерировать 3 поста (DeepSeek выбирает)
    python scripts/regen.py --send               # сгенерировать и отправить в Telegram
    python scripts/regen.py --index 2 --send     # конкретная статья → в канал
    python scripts/regen.py --send --dry-run     # показать без реальной отправки

Переменные окружения:
    DEEPSEEK_API_KEY  — ключ DeepSeek (обязателен)
    OPENAI_API_KEY    — ключ OpenAI Images (необязателен, для картинок)
    BOT_TOKEN         — токен Telegram-бота (нужен для --send)
    CHANNEL_ID        — ID канала (нужен для --send)
"""

import argparse
import json
import logging
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

# Добавляем каталог scripts в sys.path для импорта соседних модулей
sys.path.insert(0, str(Path(__file__).parent))

from monitor import (
    POSTS_DIR,
    ARTICLES_CACHE_FILE,
    MOSCOW_OFFSET,
    DEEPSEEK_API_KEY,
    OPENAI_API_KEY,
    filter_relevant_with_deepseek,
    generate_post,
    extract_image_prompt,
    generate_image,
    build_output,
)
from poster import (
    BOT_TOKEN,
    CHANNEL_ID,
    extract_posts,
    send_photo,
    send_message,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# Пауза между постами при отправке через regen (короткая — для ручного теста)
REGEN_PAUSE_BETWEEN_POSTS = 5


def load_cache() -> list[dict]:
    """Загружает кэш статей из posts/articles_cache.json."""
    if not ARTICLES_CACHE_FILE.exists():
        log.error(
            "Кэш статей не найден: %s\n"
            "Сначала запустите:  python scripts/monitor.py",
            ARTICLES_CACHE_FILE,
        )
        sys.exit(1)

    try:
        data = json.loads(ARTICLES_CACHE_FILE.read_text(encoding="utf-8"))
        log.info("Загружено %d статей из кэша (%s)", len(data), ARTICLES_CACHE_FILE.name)
        return data
    except (json.JSONDecodeError, OSError) as exc:
        log.error("Не удалось прочитать кэш: %s", exc)
        sys.exit(1)


def print_articles(articles: list[dict]) -> None:
    """Выводит пронумерованный список статей из кэша."""
    print(f"\nСтатей в кэше: {len(articles)}\n")
    for i, a in enumerate(articles):
        source = a.get("source_name", "?")
        title = (a.get("title") or "")[:80]
        print(f"  [{i:2d}] {source} — {title}")
    print()


def send_posts_to_telegram(md_file: Path, dry_run: bool) -> None:
    """Читает готовый regen-файл и отправляет посты в Telegram."""
    content = md_file.read_text(encoding="utf-8")
    posts = extract_posts(content)

    if not posts:
        log.warning("Постов для отправки не найдено в %s", md_file.name)
        return

    log.info(
        "Отправляю %d пост(а) в Telegram%s…",
        len(posts),
        " (DRY RUN)" if dry_run else "",
    )

    for i, (post_text, image_filename) in enumerate(posts):
        if i > 0 and not dry_run:
            log.info("Пауза %d сек…", REGEN_PAUSE_BETWEEN_POSTS)
            time.sleep(REGEN_PAUSE_BETWEEN_POSTS)

        log.info("Отправляю пост %d/%d…", i + 1, len(posts))

        # Фото — отдельным сообщением (без caption), текст — следом
        if image_filename:
            image_path = POSTS_DIR / image_filename
            if image_path.exists():
                send_photo(image_path, dry_run=dry_run)
            else:
                log.warning("Изображение не найдено: %s — отправляю без картинки", image_filename)

        success = send_message(post_text, dry_run=dry_run)

        if not success:
            log.error("Не удалось отправить пост %d", i + 1)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Повторная генерация постов из кэша без парсинга источников",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Примеры:\n"
            "  python scripts/regen.py --list\n"
            "  python scripts/regen.py --index 2 --send\n"
            "  python scripts/regen.py --num 3\n"
        ),
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="Показать список статей в кэше и выйти",
    )
    parser.add_argument(
        "--num",
        type=int,
        default=1,
        metavar="N",
        help="Количество постов (по умолчанию: 1, DeepSeek выбирает)",
    )
    parser.add_argument(
        "--index",
        type=str,
        default=None,
        metavar="N[,M,...]",
        help="Индексы статей из --list (через запятую, 0-based). Если не задан — DeepSeek выберет сам",
    )
    parser.add_argument(
        "--send",
        action="store_true",
        help="Отправить сгенерированный пост в Telegram-канал",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Показать пост без реальной отправки (работает с --send)",
    )
    args = parser.parse_args()

    log.info("=" * 60)
    log.info("regen — повторная генерация постов из кэша")
    log.info("=" * 60)

    if not DEEPSEEK_API_KEY:
        log.error("DEEPSEEK_API_KEY не установлен")
        sys.exit(1)

    articles = load_cache()

    # --list: показать кэш и выйти
    if args.list:
        print_articles(articles)
        return

    # Выбираем статьи: по индексам или через DeepSeek-фильтр
    if args.index is not None:
        try:
            indices = [int(x.strip()) for x in args.index.split(",")]
            selected = [articles[i] for i in indices if 0 <= i < len(articles)]
        except (ValueError, IndexError) as exc:
            log.error("Неверные индексы: %s", exc)
            sys.exit(1)

        if not selected:
            log.error("Ни одна статья не выбрана (индексы: %s, всего в кэше: %d)", args.index, len(articles))
            sys.exit(1)

        log.info("Выбрано %d статей по индексам: %s", len(selected), args.index)
        relevant = selected[: args.num]
    else:
        # DeepSeek сам выберет наиболее релевантные
        relevant = filter_relevant_with_deepseek(articles)
        if not relevant:
            log.warning("DeepSeek не нашёл релевантных статей в кэше")
            print_articles(articles)
            log.info("Используйте --index N чтобы выбрать статью вручную (см. список выше)")
            sys.exit(0)
        relevant = relevant[: args.num]

    # Генерируем посты
    now_msk = datetime.now(timezone.utc) + MOSCOW_OFFSET
    articles_with_posts: list[tuple] = []

    for article in relevant:
        try:
            post_text = generate_post(article)

            image_path = None
            if OPENAI_API_KEY:
                img_prompt = extract_image_prompt(post_text)
                if img_prompt:
                    image_path = generate_image(img_prompt, now_msk, len(articles_with_posts) + 1)

            articles_with_posts.append((article, post_text, image_path))

            if len(articles_with_posts) < len(relevant):
                time.sleep(2)
        except Exception as exc:
            log.error("Ошибка генерации поста для «%s»: %s", (article.get("title") or "")[:60], exc)

    if not articles_with_posts:
        log.error("Ни одного поста не сгенерировано")
        sys.exit(1)

    # Сохраняем в regen_YYYY-MM-DD_HH-MM.md
    # Файл с префиксом regen_ игнорируется poster.py при авто-постинге
    timestamp = now_msk.strftime("%Y-%m-%d_%H-%M")
    out_file = POSTS_DIR / f"regen_{timestamp}.md"
    content = build_output(articles_with_posts, now_msk, total_sources=len(articles))
    out_file.write_text(content, encoding="utf-8")
    log.info("Результат сохранён: %s", out_file)

    # Превью поста в консоль
    posts = extract_posts(content)
    for i, (post_text, img_name) in enumerate(posts):
        print(f"\n{'─' * 60}")
        print(f"  Пост {i + 1}/{len(posts)}" + (f"  [🖼 {img_name}]" if img_name else ""))
        print("─" * 60)
        print(post_text)

    print(f"\n{'─' * 60}")
    print(f"Файл: {out_file}")
    if not args.send:
        print("Для отправки в канал: python scripts/regen.py --send")
        if args.index is None:
            print("Для ручного выбора статьи: python scripts/regen.py --list, затем --index N")
    print("─" * 60)

    # Отправляем в Telegram если --send
    if args.send:
        if not BOT_TOKEN or not CHANNEL_ID:
            log.error("BOT_TOKEN и CHANNEL_ID обязательны для --send. Добавьте в .env")
            sys.exit(1)
        send_posts_to_telegram(out_file, dry_run=args.dry_run)

    log.info("=" * 60)
    log.info("Готово. Постов сгенерировано: %d", len(articles_with_posts))
    log.info("=" * 60)


if __name__ == "__main__":
    main()
