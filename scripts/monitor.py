#!/usr/bin/env python3
"""
monitor.py — Мониторинг источников и генерация постов для Telegram-канала
о земельном и водном праве России.

Запуск:
    python scripts/monitor.py

Переменные окружения:
    GEMINI_API_KEY — ключ API Google Gemini
"""

import logging
import os
import sys
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import urljoin

import google.generativeai as genai
import openpyxl
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Константы и настройки
# ---------------------------------------------------------------------------

MOSCOW_OFFSET = timedelta(hours=3)
REPO_ROOT = Path(__file__).parent.parent
POSTS_DIR = REPO_ROOT / "posts"
SOURCES_FILE = REPO_ROOT / "sources.xlsx"
LAST_CHECK_FILE = POSTS_DIR / "last_check.txt"

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = "gemini-2.0-flash"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Ключевые слова для фильтрации релевантности
# ---------------------------------------------------------------------------

INCLUDE_KEYWORDS = [
    "земельный", "земельного", "земельной", "земельных",
    "водный", "водного", "водной", "водных",
    "водопользование", "водопользования",
    "росводресурсы", "росводресурс",
    "росреестр", "росреестра",
    "гидротехнич",
    "пруд", "пруда", "пруду", "прудов",
    "водоём", "водоема", "водоему", "водоёма",
    "аренда земли", "аренды земли",
    "земельный участок", "земельного участка", "земельные участки",
    "лесфонд", "лесной фонд",
    "кадастр", "кадастровый", "кадастровой",
    "самовольное занятие", "самовольный захват",
    "предписание",
    "судебная практика", "судебное решение",
    "нарушение водного", "нарушение земельного",
    "межевание", "межевой план",
    "сервитут",
    "водоохранная зона",
    "береговая полоса",
    "гидротехническое сооружение", "гтс",
    "декларация безопасности",
    "росприроднадзор",
    "роснедра", "роснедр",
    "минприроды",
    "земельный кодекс",
    "водный кодекс",
    "лесной кодекс",
    "право собственности на землю",
    "изъятие земель",
    "перевод земель",
]

# Жёсткие стоп-слова: нерелевантно ТОЛЬКО если НЕТ ни одного включающего слова
HARD_EXCLUDE_KEYWORDS = [
    "убийство", "убит", "теракт", "взрыв", "наркотики", "наркотик",
]

# ---------------------------------------------------------------------------
# Промпт для Gemini API
# ---------------------------------------------------------------------------

GEMINI_POST_PROMPT = """\
Ты — редактор Telegram-канала о земельном и водном праве России.
Канал читают предприниматели, фермеры, арендаторы, владельцы участков и водоёмов.
Автор канала — консультант по оформлению водопользования, ГТС, прудов, земельных участков и лесфонда.

Вот исходный материал:
ЗАГОЛОВОК: {title}
ТЕКСТ/АНОНС: {content}
ИСТОЧНИК: {source_name}
ССЫЛКА: {url}

Сформируй готовый пост для Telegram строго по структуре:

**ГОТОВЫЙ ПОСТ:**
[Пост 800–1500 символов с пробелами. Структура:
1. Цепляющая первая фраза (интрига, лёгкий юмор или провокация)
2. Суть материала простым языком — что произошло и что это значит для читателя на практике
3. Если упомянут нормативный акт — обязательно назови его
4. Если материал про конкретный регион — упомяни
5. Вывод с лёгкой иронией над бюрократией или ситуацией (не над людьми и организациями)
6. Один CTA — выбери наиболее подходящий по тону из трёх ниже

Тон: живой, немного хулиганский, без канцелярита, без хамства. Никаких юридических гарантий — только "как правило", "по практике", "есть риск что".]

**ВАРИАНТЫ CTA:**
Нейтральный: [вариант]
С юмором: [вариант]
Прямой: [вариант]

**ВИЗУАЛ:**
Prompt (EN): [промпт для Midjourney/DALL-E, flat design или editorial cartoon, 16:9, no text in image, отражает суть метафорично, элементы российской действительности]
Описание (RU): [2–3 слова]
Запасной вариант: [более простой промпт]
"""

# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------


def moscow_now() -> datetime:
    """Текущее время по Москве (UTC+3)."""
    return datetime.now(timezone.utc) + MOSCOW_OFFSET


def get_last_check() -> datetime:
    """
    Возвращает время последней проверки (UTC).
    Если файла нет — 8 часов назад от текущего момента.
    """
    if LAST_CHECK_FILE.exists():
        try:
            ts_str = LAST_CHECK_FILE.read_text(encoding="utf-8").strip()
            return datetime.fromisoformat(ts_str)
        except (ValueError, OSError) as exc:
            log.warning("Не удалось прочитать last_check.txt: %s", exc)
    return datetime.now(timezone.utc) - timedelta(hours=8)


def save_last_check(dt: datetime) -> None:
    """Сохраняет время текущей проверки в UTC ISO-формате."""
    POSTS_DIR.mkdir(exist_ok=True)
    LAST_CHECK_FILE.write_text(dt.isoformat(), encoding="utf-8")


def make_session() -> requests.Session:
    """Создаёт HTTP-сессию с корректными заголовками."""
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "Accept-Encoding": "gzip, deflate, br",
        }
    )
    return session


# ---------------------------------------------------------------------------
# Чтение источников
# ---------------------------------------------------------------------------


def read_sources() -> list[dict]:
    """
    Читает sources.xlsx и возвращает список активных источников.
    Завершает программу если файл не найден.
    """
    if not SOURCES_FILE.exists():
        log.error("Файл sources.xlsx не найден: %s", SOURCES_FILE)
        sys.exit(1)

    wb = openpyxl.load_workbook(SOURCES_FILE)

    if "Sources" not in wb.sheetnames:
        log.error("В sources.xlsx нет листа 'Sources'")
        sys.exit(1)

    ws = wb["Sources"]
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

    sources = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        record = dict(zip(headers, row))
        active_val = str(record.get("active", "")).strip().lower()
        if active_val in ("да", "yes", "1", "true"):
            sources.append(record)

    log.info("Загружено активных источников: %d", len(sources))
    return sources


# ---------------------------------------------------------------------------
# Парсинг источников
# ---------------------------------------------------------------------------


def fetch_site_articles(source: dict, session: requests.Session) -> list[dict]:
    """
    Получает список статей/новостей с сайта через HTTP GET + BeautifulSoup.
    Возвращает список словарей: {title, content, url, source_name}.
    Если источник недоступен — логирует предупреждение и возвращает [].
    """
    url = str(source.get("url", "")).strip()
    name = str(source.get("name", url))
    articles = []

    try:
        resp = session.get(url, timeout=20)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or "utf-8"
        soup = BeautifulSoup(resp.text, "lxml")

        # Удаляем шум: навигацию, подвал, скрипты
        for tag in soup(["script", "style", "nav", "footer", "header", "aside", "noscript"]):
            tag.decompose()

        # Пробуем разные CSS-селекторы для поиска новостных элементов
        # (разные сайты используют разную разметку)
        news_blocks = []
        selectors_to_try = [
            "article",
            ".news-item",
            ".article-item",
            ".post-item",
            ".material-item",
            ".b-news-item",
            ".news__item",
            ".entry",
            ".list-item",
            "li.item",
        ]
        for selector in selectors_to_try:
            found = soup.select(selector)
            if found:
                news_blocks = found[:20]
                log.debug("Источник '%s': использован селектор '%s'", name, selector)
                break

        # Fallback: ищем все заголовки (h2/h3/h4) со ссылками
        if not news_blocks:
            for heading in soup.find_all(["h2", "h3", "h4"]):
                if heading.find("a"):
                    news_blocks.append(heading)
            news_blocks = news_blocks[:20]

        seen_titles: set[str] = set()
        for block in news_blocks:
            # Извлекаем заголовок и ссылку
            link_tag = block.find("a") if block.name != "a" else block
            if not link_tag:
                continue

            title = link_tag.get_text(strip=True)
            if not title or len(title) < 10:
                title = block.get_text(" ", strip=True)[:200]

            if not title or title in seen_titles:
                continue
            seen_titles.add(title)

            href = str(link_tag.get("href", "")).strip()
            if href and not href.startswith("http"):
                href = urljoin(url, href)

            content = block.get_text(separator=" ", strip=True)[:600]

            articles.append(
                {
                    "title": title,
                    "content": content,
                    "url": href or url,
                    "source_name": name,
                }
            )

        log.info("Источник '%s': найдено элементов — %d", name, len(articles))

    except requests.exceptions.Timeout:
        log.warning("Источник '%s': таймаут запроса (20 сек)", name)
    except requests.exceptions.ConnectionError as exc:
        log.warning("Источник '%s': ошибка соединения — %s", name, exc)
    except requests.exceptions.HTTPError as exc:
        log.warning("Источник '%s': HTTP-ошибка — %s", name, exc)
    except Exception as exc:
        log.warning("Источник '%s': непредвиденная ошибка — %s", name, exc)

    return articles


def fetch_telegram_channel(source: dict, _session: requests.Session) -> list[dict]:
    """
    ЗАГЛУШКА для мониторинга Telegram-каналов.

    Прямой парсинг Telegram-каналов требует авторизации через MTProto API.
    Для полноценной реализации используйте Telethon или Pyrogram:

        1. Зарегистрируйтесь на https://my.telegram.org
        2. Создайте приложение → получите API_ID и API_HASH
        3. Установите: pip install telethon>=1.36.0
        4. Добавьте переменные в .env:
               TELEGRAM_API_ID=12345678
               TELEGRAM_API_HASH=abc123def456...
               TELEGRAM_SESSION=base64_encoded_session_string
        5. Замените эту функцию на реализацию через TelegramClient
        6. Активируйте источник в sources.xlsx (active = да)

    Подробнее — см. раздел "Мониторинг Telegram-каналов" в CLAUDE.md.
    """
    name = str(source.get("name", source.get("url", "Unknown")))
    log.info(
        "Источник '%s' (telegram) — пропущен. "
        "Мониторинг Telegram-каналов требует отдельной настройки MTProto. "
        "Инструкция: см. CLAUDE.md → раздел 'Мониторинг Telegram-каналов'.",
        name,
    )
    return []


# ---------------------------------------------------------------------------
# Оценка релевантности
# ---------------------------------------------------------------------------


def score_relevance(article: dict) -> int:
    """
    Оценивает релевантность статьи по ключевым словам.
    Возвращает счёт >= 1 если релевантно, 0 — если нет.
    """
    text = (
        (article.get("title") or "") + " " + (article.get("content") or "")
    ).lower()

    # Проверяем жёсткие стоп-слова
    for stop_kw in HARD_EXCLUDE_KEYWORDS:
        if stop_kw in text:
            # Допускаем только если есть хотя бы одно включающее слово
            has_include = any(kw in text for kw in INCLUDE_KEYWORDS[:6])
            if not has_include:
                return 0

    # Считаем совпадения с включающими ключевыми словами
    score = sum(1 for kw in INCLUDE_KEYWORDS if kw in text)
    return score


def collect_relevant_articles(sources: list[dict]) -> list[dict]:
    """
    Обходит все активные источники, собирает статьи,
    фильтрует по релевантности, возвращает топ-3.
    """
    session = make_session()
    all_articles: list[dict] = []

    for source in sources:
        src_type = str(source.get("type", "site")).strip().lower()

        if src_type == "telegram":
            articles = fetch_telegram_channel(source, session)
        else:
            articles = fetch_site_articles(source, session)
            time.sleep(1)  # Вежливая пауза между запросами к сайтам

        for article in articles:
            score = score_relevance(article)
            if score > 0:
                article["score"] = score
                all_articles.append(article)

    # Сортируем по релевантности (больше совпадений = выше)
    all_articles.sort(key=lambda x: x.get("score", 0), reverse=True)
    top_articles = all_articles[:3]

    log.info(
        "Всего релевантных материалов: %d | Выбрано для публикации: %d",
        len(all_articles),
        len(top_articles),
    )
    return top_articles


# ---------------------------------------------------------------------------
# Генерация постов через Gemini API
# ---------------------------------------------------------------------------


def generate_post(article: dict, model: genai.GenerativeModel) -> str:
    """Генерирует готовый пост через Google Gemini API."""
    prompt = GEMINI_POST_PROMPT.format(
        title=article.get("title", "Без заголовка"),
        content=(article.get("content") or "")[:2000],
        source_name=article.get("source_name", ""),
        url=article.get("url", ""),
    )

    log.info("Генерирую пост: «%s»", (article.get("title") or "")[:70])

    response = model.generate_content(prompt)
    return response.text


# ---------------------------------------------------------------------------
# Формирование и сохранение результата
# ---------------------------------------------------------------------------


def build_output(
    articles_with_posts: list[tuple],
    now_msk: datetime,
    total_sources: int,
) -> str:
    """Формирует итоговый Markdown-файл с постами."""
    date_str = now_msk.strftime("%Y-%m-%d")
    time_str = now_msk.strftime("%H:%M")

    lines = [
        f"# Мониторинг {date_str} {time_str}",
        f"Проверено источников: {total_sources}",
        f"Найдено релевантных материалов: {len(articles_with_posts)}",
        "",
    ]

    if not articles_with_posts:
        lines.append(
            f"[{date_str}][{time_str}] — релевантных материалов не найдено. "
            f"Проверено источников: {total_sources}"
        )
        return "\n".join(lines)

    for i, (article, post_text) in enumerate(articles_with_posts, start=1):
        lines += [
            "---",
            "",
            f"## Пост {i}",
            "",
            f"**Источник:** {article.get('source_name', '')} — {article.get('url', '')}",
            f"**Суть:** {article.get('title', '')}",
            "",
            post_text.strip(),
            "",
        ]

    return "\n".join(lines)


def save_output(content: str, now_msk: datetime) -> Path:
    """Сохраняет результат в posts/ГГГГ-ММ-ДД_ЧЧ.md."""
    POSTS_DIR.mkdir(exist_ok=True)
    filename = now_msk.strftime("%Y-%m-%d_%H") + ".md"
    out_file = POSTS_DIR / filename
    out_file.write_text(content, encoding="utf-8")
    log.info("Результат сохранён: %s", out_file)
    return out_file


# ---------------------------------------------------------------------------
# Точка входа
# ---------------------------------------------------------------------------


def main() -> None:
    log.info("=" * 60)
    log.info("Запуск мониторинга источников")
    log.info("=" * 60)

    now_utc = datetime.now(timezone.utc)
    now_msk = now_utc + MOSCOW_OFFSET
    last_check = get_last_check()

    log.info("Текущее время (МСК): %s", now_msk.strftime("%Y-%m-%d %H:%M"))
    log.info("Последняя проверка:  %s", last_check.isoformat())

    # 1. Читаем источники
    sources = read_sources()
    if not sources:
        log.warning("Нет активных источников в sources.xlsx — выходим")
        save_last_check(now_utc)
        return

    # 2. Собираем релевантные материалы
    relevant = collect_relevant_articles(sources)

    if not relevant:
        log.info("Релевантных материалов не найдено")
        content = build_output([], now_msk, len(sources))
        save_output(content, now_msk)
        save_last_check(now_utc)
        return

    # 3. Генерируем посты через Gemini API
    if not GEMINI_API_KEY:
        log.error(
            "GEMINI_API_KEY не установлен. "
            "Генерация постов невозможна. Добавьте ключ в .env или GitHub Secrets."
        )
        save_last_check(now_utc)
        sys.exit(1)

    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel(GEMINI_MODEL)
    log.info("Используется модель: %s", GEMINI_MODEL)

    articles_with_posts: list[tuple] = []

    for article in relevant:
        try:
            post_text = generate_post(article, model)
            articles_with_posts.append((article, post_text))
            time.sleep(2)  # Пауза между вызовами Gemini API
        except Exception as exc:
            log.error(
                "Ошибка генерации поста для «%s»: %s",
                (article.get("title") or "")[:60],
                exc,
            )

    # 4. Сохраняем результат
    content = build_output(articles_with_posts, now_msk, len(sources))
    save_output(content, now_msk)
    save_last_check(now_utc)

    log.info("=" * 60)
    log.info("Мониторинг завершён. Сформировано постов: %d", len(articles_with_posts))
    log.info("=" * 60)


if __name__ == "__main__":
    main()
