#!/usr/bin/env python3
"""
poster.py — Публикация постов из posts/*.md в Telegram-канал.

Запуск:
    python scripts/poster.py

Переменные окружения:
    BOT_TOKEN  — токен Telegram-бота от @BotFather
    CHANNEL_ID — ID или @username канала (например: @my_channel или -1001234567890)
"""

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
PAUSE_BETWEEN_POSTS = 30   # секунды между постами из одного файла
TELEGRAM_MAX_LENGTH = 4096  # максимум символов в одном сообщении Telegram

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


def extract_posts(content: str) -> list[str]:
    """
    Извлекает все блоки «ГОТОВЫЙ ПОСТ:» из markdown-файла.
    Возвращает список текстов постов (без заголовка блока).
    """
    # Ищем всё между **ГОТОВЫЙ ПОСТ:** и следующим разделителем
    pattern = re.compile(
        r"\*\*ГОТОВЫЙ ПОСТ:\*\*\s*\n(.*?)(?=\n\*\*ВАРИАНТЫ CTA|\n---|\n## Пост\s|\Z)",
        re.DOTALL,
    )
    posts = [m.group(1).strip() for m in pattern.finditer(content)]
    return [p for p in posts if p]


# ---------------------------------------------------------------------------
# Конвертация Markdown → HTML для Telegram
# ---------------------------------------------------------------------------


def markdown_to_html(text: str) -> str:
    """
    Конвертирует базовую Markdown-разметку в HTML, поддерживаемый Telegram.
    Telegram поддерживает: <b>, <i>, <u>, <s>, <code>, <pre>, <a href="">.

    Порядок: сначала извлекаем Markdown-конструкции, потом экранируем
    обычный текст, чтобы &, <, > не сломали HTML-парсер Telegram.
    """
    # Шаг 1: Вытаскиваем Markdown-конструкции в плейсхолдеры,
    # чтобы их содержимое не затронуло дальнейшее экранирование.

    links: list[tuple[str, str]] = []   # (display_text, url)
    bolds: list[str] = []
    italics: list[str] = []
    codes: list[str] = []

    def save_link(m: re.Match) -> str:
        links.append((m.group(1), m.group(2)))
        return f"\x00LINK{len(links) - 1}\x00"

    def save_bold(m: re.Match) -> str:
        bolds.append(m.group(1))
        return f"\x00BOLD{len(bolds) - 1}\x00"

    def save_italic(m: re.Match) -> str:
        italics.append(m.group(1))
        return f"\x00ITAL{len(italics) - 1}\x00"

    def save_code(m: re.Match) -> str:
        codes.append(m.group(1))
        return f"\x00CODE{len(codes) - 1}\x00"

    text = re.sub(r"\[([^\[\]]+?)\]\((https?://[^\)]+?)\)", save_link, text)
    text = re.sub(r"\*\*(.+?)\*\*", save_bold, text, flags=re.DOTALL)
    text = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", save_italic, text)
    text = re.sub(r"`([^`]+?)`", save_code, text)

    # Шаг 2: Экранируем HTML-спецсимволы в оставшемся тексте
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    # Шаг 3: Восстанавливаем Markdown-конструкции как HTML-теги
    for i, (display, url) in enumerate(links):
        display = display.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        text = text.replace(f"\x00LINK{i}\x00", f'<a href="{url}">{display}</a>')
    for i, content in enumerate(bolds):
        content = content.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        text = text.replace(f"\x00BOLD{i}\x00", f"<b>{content}</b>")
    for i, content in enumerate(italics):
        content = content.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        text = text.replace(f"\x00ITAL{i}\x00", f"<i>{content}</i>")
    for i, content in enumerate(codes):
        content = content.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        text = text.replace(f"\x00CODE{i}\x00", f"<code>{content}</code>")

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


def send_message(text: str) -> bool:
    """
    Отправляет сообщение в Telegram-канал через Bot API.
    Сначала пробует с HTML-разметкой; при ошибке разбора — без разметки.
    При ошибке 429 (flood control) — до 3 повторных попыток с задержкой.
    Возвращает True при успехе.
    """
    api_url = f"{TELEGRAM_API_BASE.format(token=BOT_TOKEN)}/sendMessage"
    html_text = markdown_to_html(sanitize_for_telegram(text))

    payload = {
        "chat_id": CHANNEL_ID,
        "text": html_text,
        "parse_mode": "HTML",
        "disable_web_page_preview": False,
    }

    retry_delays = [10, 30, 60]  # секунды между попытками при 429

    try:
        for attempt, delay in enumerate(retry_delays + [None], start=1):
            resp = requests.post(api_url, json=payload, timeout=30)
            data = resp.json()

            # Flood control — ждём и повторяем
            if resp.status_code == 429:
                retry_after = data.get("parameters", {}).get("retry_after", delay)
                if delay is not None:
                    log.warning(
                        "Telegram 429 (попытка %d/%d), жду %d сек…",
                        attempt, len(retry_delays) + 1, retry_after,
                    )
                    time.sleep(retry_after)
                    continue
                # Все попытки исчерпаны
                log.error("Telegram API: исчерпаны все попытки после 429")
                return False

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


# ---------------------------------------------------------------------------
# Обработка файлов
# ---------------------------------------------------------------------------


def process_file(post_file: Path) -> bool:
    """
    Обрабатывает один файл: извлекает посты и отправляет в Telegram.
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

    for i, post_text in enumerate(posts):
        if i > 0:
            log.info("Пауза %d сек перед следующим постом...", PAUSE_BETWEEN_POSTS)
            time.sleep(PAUSE_BETWEEN_POSTS)

        log.info("Отправляю пост %d/%d...", i + 1, len(posts))
        success = send_message(post_text)

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
    log.info("=" * 60)
    log.info("Запуск постера")
    log.info("=" * 60)

    # Проверяем обязательные переменные окружения
    if not BOT_TOKEN:
        log.error(
            "BOT_TOKEN не установлен. "
            "Добавьте его в .env или GitHub Secrets."
        )
        sys.exit(1)

    if not CHANNEL_ID:
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
        and p.name not in posted
        and p.name not in {"posted.log", "errors.log"}
    )

    if not candidates:
        log.info("Новых файлов для публикации не найдено")
        return

    log.info("Файлов для публикации: %d", len(candidates))

    for file_idx, post_file in enumerate(candidates):
        if file_idx > 0:
            log.info("Пауза %d сек перед следующим файлом...", PAUSE_BETWEEN_POSTS)
            time.sleep(PAUSE_BETWEEN_POSTS)

        try:
            success = process_file(post_file)

            if success:
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
    log.info("Постер завершил работу")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
