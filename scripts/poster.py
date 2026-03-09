#!/usr/bin/env python3
"""
poster.py — Публикация постов из posts/*.md в Telegram-канал.

Запуск:
    python scripts/poster.py

Переменные окружения:
    BOT_TOKEN  — токен Telegram-бота от @BotFather
    CHANNEL_ID — ID или @username канала (например: @my_channel или -1001234567890)
"""

import argparse
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Константы
# ---------------------------------------------------------------------------

POSTS_DIR = Path(__file__).parent.parent / "posts"
POSTED_LOG = POSTS_DIR / "posted.log"
ERRORS_LOG = POSTS_DIR / "errors.log"

BOT_TOKEN = os.getenv("BOT_TOKEN")
CHANNEL_ID = os.getenv("CHANNEL_ID")

TELEGRAM_API_BASE = "https://api.telegram.org/bot{token}"
PAUSE_BETWEEN_POSTS = 30    # секунды между постами из одного файла
TELEGRAM_MAX_LENGTH = 4096  # максимум символов в одном сообщении Telegram
TELEGRAM_CAPTION_MAX = 1024  # максимум символов в подписи к фото

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Управление логами опубликованных / ошибок
# ---------------------------------------------------------------------------


def load_posted() -> set[str]:
    """Возвращает множество уже опубликованных имён файлов."""
    if not POSTED_LOG.exists():
        return set()
    lines = POSTED_LOG.read_text(encoding="utf-8").splitlines()
    return {line.strip() for line in lines if line.strip()}


def mark_posted(original_filename: str) -> None:
    """Добавляет файл в posted.log."""
    with open(POSTED_LOG, "a", encoding="utf-8") as f:
        f.write(original_filename + "\n")


def log_error(filename: str, error: str) -> None:
    """Записывает ошибку в errors.log."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(ERRORS_LOG, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {filename}: {error}\n")


# ---------------------------------------------------------------------------
# Извлечение постов из Markdown-файла
# ---------------------------------------------------------------------------


def extract_posts(content: str) -> list[tuple[str, str | None]]:
    """
    Извлекает все блоки «ГОТОВЫЙ ПОСТ:» из markdown-файла.
    Возвращает список кортежей (текст_поста, имя_файла_изображения_или_None).
    """
    # Разбиваем на блоки по разделителю "## Пост N"
    # Каждый блок содержит метаданные поста + текст + опциональное **ИЗОБРАЖЕНИЕ:**
    post_pattern = re.compile(
        r"\*\*ГОТОВЫЙ ПОСТ:\*\*\s*\n(.*?)(?=\n\*\*ВАРИАНТЫ CTA|\n\*\*ВИЗУАЛ:|\n\*\*CTA|\n---|\n## Пост\s|\Z)",
        re.DOTALL,
    )
    image_pattern = re.compile(r"\*\*ИЗОБРАЖЕНИЕ:\*\*\s*(\S+\.png)")

    # Разбиваем на секции "## Пост N" чтобы матчить изображение к нужному посту
    sections = re.split(r"\n---\n\n## Пост \d+\n", content)

    results = []
    for section in sections:
        post_match = post_pattern.search(section)
        if not post_match:
            continue
        post_text = post_match.group(1).strip()
        if not post_text:
            continue
        img_match = image_pattern.search(section)
        image_filename = img_match.group(1).strip() if img_match else None
        results.append((post_text, image_filename))

    return results


# ---------------------------------------------------------------------------
# Конвертация Markdown → HTML для Telegram
# ---------------------------------------------------------------------------


def markdown_to_html(text: str) -> str:
    """
    Конвертирует базовую Markdown-разметку в HTML, поддерживаемый Telegram.
    Telegram поддерживает: <b>, <i>, <u>, <s>, <code>, <pre>, <a href="">.

    Обрабатывается слева направо, поэтому порядок замен важен.
    """
    # Экранируем HTML-символы, которые НЕ являются частью наших тегов
    # Сначала работаем с текстом как есть, затем применяем замены

    # [текст](url) → <a href="url">текст</a>
    text = re.sub(r"\[([^\[\]]+?)\]\((https?://[^\)]+?)\)", r'<a href="\2">\1</a>', text)

    # **текст** → <b>текст</b>
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text, flags=re.DOTALL)

    # *текст* → <i>текст</i> (только одиночные звёздочки)
    text = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"<i>\1</i>", text)

    # `текст` → <code>текст</code>
    text = re.sub(r"`([^`]+?)`", r"<code>\1</code>", text)

    return text


def sanitize_for_telegram(text: str) -> str:
    """
    Очищает текст перед отправкой: обрезает до лимита Telegram,
    удаляет некорректные символы.
    """
    # Убираем нулевые байты и управляющие символы (кроме \n и \t)
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)

    if len(text) > TELEGRAM_MAX_LENGTH:
        log.warning(
            "Текст поста (%d символов) превышает лимит Telegram (%d), обрезаю",
            len(text),
            TELEGRAM_MAX_LENGTH,
        )
        text = text[: TELEGRAM_MAX_LENGTH - 3] + "..."

    return text


# ---------------------------------------------------------------------------
# Отправка в Telegram
# ---------------------------------------------------------------------------


def send_message(text: str, dry_run: bool = False) -> bool:
    """
    Отправляет сообщение в Telegram-канал через Bot API.
    Сначала пробует с HTML-разметкой; при ошибке разбора — без разметки.
    В режиме dry_run показывает, что будет отправлено, без реальной отправки.
    Возвращает True при успехе (или при dry_run).
    """
    api_url = f"{TELEGRAM_API_BASE.format(token=BOT_TOKEN)}/sendMessage"
    html_text = markdown_to_html(sanitize_for_telegram(text))

    payload = {
        "chat_id": CHANNEL_ID,
        "text": html_text,
        "parse_mode": "HTML",
        "disable_web_page_preview": False,
    }

    if dry_run:
        log.info("=== DRY RUN — пост НЕ отправлен ===")
        log.info("chat_id: %s", CHANNEL_ID)
        log.info("parse_mode: HTML")
        log.info("disable_web_page_preview: False")
        log.info("Длина текста: %d символов", len(html_text))
        log.info("--- Текст поста (HTML) ---")
        print(html_text)
        log.info("--- Конец поста ---")
        return True

    try:
        resp = requests.post(api_url, json=payload, timeout=30)
        data = resp.json()

        if data.get("ok"):
            msg_id = data.get("result", {}).get("message_id", "?")
            log.info("✓ Пост отправлен (message_id=%s)", msg_id)
            return True

        error_desc = data.get("description", "Неизвестная ошибка")
        log.error("Telegram API вернул ошибку: %s", error_desc)

        # Если ошибка в HTML-разметке — повторяем без parse_mode
        if "can't parse" in error_desc.lower() or "bad request" in error_desc.lower():
            log.warning("Повторная попытка без HTML-разметки...")
            payload["text"] = sanitize_for_telegram(text)
            payload.pop("parse_mode", None)

            resp2 = requests.post(api_url, json=payload, timeout=30)
            data2 = resp2.json()

            if data2.get("ok"):
                log.info("✓ Пост отправлен без разметки")
                return True

            log.error(
                "Повторная попытка не удалась: %s",
                data2.get("description", "Неизвестная ошибка"),
            )

        return False

    except requests.exceptions.Timeout:
        log.error("Таймаут при отправке в Telegram (30 сек)")
        return False
    except requests.exceptions.ConnectionError as exc:
        log.error("Ошибка соединения с Telegram API: %s", exc)
        return False
    except requests.exceptions.RequestException as exc:
        log.error("HTTP-ошибка при отправке в Telegram: %s", exc)
        return False


def send_photo(image_path: Path, dry_run: bool = False) -> bool:
    """
    Отправляет фото без подписи в Telegram-канал через sendPhoto.
    Текст поста отправляется отдельным сообщением после фото — это
    обходит лимит caption (1024 символа) и позволяет публиковать
    полный текст любой длины.
    Возвращает False при ошибке — тогда публикуется только текст.
    """
    api_url = f"{TELEGRAM_API_BASE.format(token=BOT_TOKEN)}/sendPhoto"

    if dry_run:
        log.info("=== DRY RUN — фото НЕ отправлено: %s ===", image_path.name)
        return True

    try:
        with open(image_path, "rb") as f:
            resp = requests.post(
                api_url,
                data={"chat_id": CHANNEL_ID},
                files={"photo": (image_path.name, f, "image/png")},
                timeout=60,
            )
        data = resp.json()

        if data.get("ok"):
            msg_id = data.get("result", {}).get("message_id", "?")
            log.info("✓ Фото отправлено (message_id=%s)", msg_id)
            return True

        log.warning(
            "Telegram sendPhoto вернул ошибку: %s — публикую только текст",
            data.get("description", "?"),
        )
        return False

    except Exception as exc:
        log.warning("Ошибка при отправке фото (%s): %s — публикую только текст", image_path.name, exc)
        return False


# ---------------------------------------------------------------------------
# Обработка файлов
# ---------------------------------------------------------------------------


def process_file(post_file: Path, dry_run: bool = False) -> bool:
    """
    Обрабатывает один файл: извлекает посты и отправляет в Telegram.
    В режиме dry_run показывает посты без реальной отправки.
    Возвращает True, если все посты успешно отправлены (или файл пуст).
    """
    log.info("Обрабатываю файл: %s", post_file.name)

    content = post_file.read_text(encoding="utf-8")

    # Файл без постов (записано «нет материалов»)
    if (
        "релевантных материалов не найдено" in content
        and "ГОТОВЫЙ ПОСТ" not in content
    ):
        log.info(
            "Файл %s не содержит постов (нет материалов) — помечаю как обработанный",
            post_file.name,
        )
        return True

    posts = extract_posts(content)

    if not posts:
        log.warning(
            "В файле %s не найден блок «ГОТОВЫЙ ПОСТ:» — пропускаю",
            post_file.name,
        )
        return True  # Считаем обработанным, чтобы не зациклиться

    log.info("Найдено постов для отправки: %d", len(posts))
    all_sent = True

    for i, (post_text, image_filename) in enumerate(posts):
        if i > 0 and not dry_run:
            log.info("Пауза %d сек перед следующим постом...", PAUSE_BETWEEN_POSTS)
            time.sleep(PAUSE_BETWEEN_POSTS)

        log.info("Отправляю пост %d/%d...", i + 1, len(posts))

        # Отправляем фото отдельным сообщением (без caption) — обходим лимит 1024 символа
        if image_filename:
            image_path = POSTS_DIR / image_filename
            if image_path.exists():
                send_photo(image_path, dry_run=dry_run)
            else:
                log.warning("Файл изображения не найден: %s — отправляю без картинки", image_filename)

        # Текст поста — всегда отдельным сообщением после фото
        success = send_message(post_text, dry_run=dry_run)

        if not success:
            all_sent = False
            error_msg = f"Не удалось отправить пост {i + 1}/{len(posts)}"
            log.error("%s из файла %s", error_msg, post_file.name)
            log_error(post_file.name, error_msg)

    return all_sent


# ---------------------------------------------------------------------------
# Точка входа
# ---------------------------------------------------------------------------


def main() -> None:
    parser = argparse.ArgumentParser(description="Публикация постов в Telegram-канал")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Показать посты без реальной отправки в Telegram",
    )
    parser.add_argument(
        "--post",
        action="store_true",
        help="Реальная отправка постов в Telegram (поведение по умолчанию)",
    )
    args = parser.parse_args()

    dry_run = args.dry_run

    log.info("=" * 60)
    log.info("Запуск постера%s", " (DRY RUN)" if dry_run else "")
    log.info("=" * 60)

    # Проверяем обязательные переменные окружения (в dry-run BOT_TOKEN/CHANNEL_ID
    # не обязательны для вывода, но нужны для формирования payload)
    if not BOT_TOKEN and not dry_run:
        log.error(
            "BOT_TOKEN не установлен. "
            "Добавьте его в .env или GitHub Secrets."
        )
        sys.exit(1)

    if not CHANNEL_ID and not dry_run:
        log.error(
            "CHANNEL_ID не установлен. "
            "Добавьте его в .env или GitHub Secrets."
        )
        sys.exit(1)

    POSTS_DIR.mkdir(exist_ok=True)
    posted = load_posted()

    # Собираем необработанные .md файлы (без суффикса _posted и не в posted.log)
    candidates = sorted(
        p
        for p in POSTS_DIR.glob("*.md")
        if not p.stem.endswith("_posted")
        and not p.stem.startswith("regen_")   # regen.py файлы не публикуем автоматически
        and p.name not in posted
        and p.name not in {"posted.log", "errors.log"}
    )

    if not candidates:
        log.info("Новых файлов для публикации не найдено")
        return

    log.info("Файлов для публикации: %d", len(candidates))

    for post_file in candidates:
        try:
            success = process_file(post_file, dry_run=dry_run)

            if success and not dry_run:
                # Переименовываем: 2026-03-01_09.md → 2026-03-01_09_posted.md
                new_name = post_file.stem + "_posted.md"
                new_path = post_file.parent / new_name
                post_file.rename(new_path)
                mark_posted(post_file.name)
                log.info(
                    "✓ Файл помечен как опубликованный: %s → %s",
                    post_file.name,
                    new_name,
                )
            elif success and dry_run:
                log.info(
                    "DRY RUN: файл %s НЕ переименован и НЕ помечен",
                    post_file.name,
                )
            else:
                log.error(
                    "✗ Файл %s опубликован не полностью — "
                    "повторная попытка при следующем запуске",
                    post_file.name,
                )

        except Exception as exc:
            log.error(
                "Неожиданная ошибка при обработке %s: %s",
                post_file.name,
                exc,
            )
            log_error(post_file.name, str(exc))

    log.info("=" * 60)
    log.info("Постер завершил работу%s", " (DRY RUN)" if dry_run else "")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
