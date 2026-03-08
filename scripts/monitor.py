#!/usr/bin/env python3
"""
monitor.py — Мониторинг источников и генерация постов для Telegram-канала
о земельном и водном праве России.

Запуск:
    python scripts/monitor.py

Переменные окружения:
    GEMINI_API_KEY — ключ API Google Gemini
"""

import json
import logging
import os
import re
import sys
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import urljoin, urlparse

import urllib3
import xml.etree.ElementTree as ET

import openpyxl
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# Некоторые госсайты используют самоподписанные или устаревшие сертификаты
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

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
GEMINI_API_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "{model}:generateContent?key={key}"
)

MAX_ARTICLES_PER_SOURCE = 10  # Сколько заголовков брать с каждого сайта
MAX_POSTS_TO_GENERATE = 3     # Сколько постов генерировать за один запуск

# Пути RSS-фидов для перебора (добавляются к базовому URL)
RSS_CANDIDATE_PATHS = [
    "/rss",
    "/rss.xml",
    "/feed",
    "/feed.xml",
    "/atom.xml",
    "/news/rss",
    "/news/feed",
    "/export/rss",
    "/lenta/rss",
]

# Расширенный набор CSS-селекторов для HTML-парсинга новостных сайтов
HTML_NEWS_SELECTORS = [
    "article",
    ".news-item",
    ".article-item",
    ".news__item",
    ".b-news-item",
    ".post-item",
    ".material-item",
    ".entry",
    ".list-item",
    ".item",
    ".card",
    ".news-card",
    ".publication",
    ".doc-item",
    "li.item",
    "li.news",
    "li.article",
    ".pressrelease",
    ".press-release",
    ".news-list__item",
    ".articles-list__item",
    # Характерные для российских государственных и правовых сайтов
    ".news-feed__item",
    ".page-news__item",
    ".list-news-item",
    ".document-item",
    ".event-card",
    ".news-block__item",
    ".press-item",
    ".content-item",
    ".col-news",
    ".feed-item",
    ".law-item",
    "div.row-item",
    "tr.news-row",
    "td.news-title",
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Промпты для Gemini API
# ---------------------------------------------------------------------------

GEMINI_FILTER_PROMPT = """\
Ты — редактор Telegram-канала о земельном и водном праве России.
Вот пронумерованный список заголовков новостей с разных сайтов:

{headlines_list}

Выбери только те новости, которые касаются:
- земельного или водного законодательства РФ
- оформления земли, воды, ГТС, прудов, кадастра
- судебной практики по земельным и водным темам
- Росреестра, Росводресурсов, Рослесхоза, Роснедр, Минприроды, Росприроднадзора
- экологических проверок, штрафов, предписаний по земле и воде
- аренды земли, межевания, сервитутов, водоохранных зон

Ответь ТОЛЬКО валидным JSON без пояснений и без markdown-обёртки:
{{"relevant_ids": [1, 3, 5]}}

Если подходящих новостей нет — верни: {{"relevant_ids": []}}
"""

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


def gemini_generate(prompt: str) -> str:
    """
    Вызывает Gemini REST API напрямую через requests.
    При ошибке 429 делает до 3 повторных попыток с задержкой 5/10/20 сек.
    """
    url = GEMINI_API_URL.format(model=GEMINI_MODEL, key=GEMINI_API_KEY)
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.7, "maxOutputTokens": 2048},
    }
    delays = [5, 10, 20]
    for attempt in range(1, len(delays) + 2):  # попытки 1..4
        resp = requests.post(url, json=payload, timeout=60)
        if resp.status_code == 429 and attempt <= len(delays):
            wait = delays[attempt - 1]
            log.warning(
                "Gemini API: 429 Too Many Requests (попытка %d/4), жду %d сек…",
                attempt, wait,
            )
            time.sleep(wait)
            continue
        resp.raise_for_status()
        data = resp.json()
        return data["candidates"][0]["content"]["parts"][0]["text"]
    # Последняя попытка уже выкинула исключение через raise_for_status
    raise RuntimeError("Gemini API: исчерпаны все попытки")


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


def _extract_channel_name(url: str) -> str:
    """Из t.me/channelname или @channelname извлекает имя канала."""
    url = url.strip().rstrip("/")
    # Убираем @, если задан как @username
    if url.startswith("@"):
        return url.lstrip("@")
    # Из полного URL берём последний сегмент пути
    path = urlparse(url).path
    return path.strip("/").split("/")[0].lstrip("@")


def fetch_telegram_channel(source: dict, session: requests.Session) -> list[dict]:
    """
    Парсит публичный веб-вид Telegram-канала через https://t.me/s/{channel}.
    Работает для всех публичных каналов без API-ключей и MTProto.
    Если канал приватный или недоступен — логирует предупреждение и возвращает [].
    """
    url = str(source.get("url", "")).strip()
    name = str(source.get("name", url))
    channel = _extract_channel_name(url)

    if not channel:
        log.warning("Источник '%s': не удалось извлечь имя канала из URL '%s'", name, url)
        return []

    web_url = f"https://t.me/s/{channel}"
    articles = []

    try:
        resp = session.get(web_url, timeout=30, verify=False)
        if resp.status_code == 404:
            log.warning("Источник '%s': канал @%s не найден или приватный", name, channel)
            return []
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")

        # Telegram web view: каждое сообщение — .tgme_widget_message_wrap
        messages = soup.select(".tgme_widget_message_wrap")
        # Сообщения идут от старых к новым — берём последние (самые свежие)
        messages = messages[-MAX_ARTICLES_PER_SOURCE:]

        seen: set[str] = set()
        for msg in reversed(messages):  # сначала самые новые
            text_el = msg.select_one(".tgme_widget_message_text")
            if not text_el:
                # Пост без текста (только фото/видео) — пропускаем
                continue

            text = text_el.get_text(separator=" ", strip=True)
            if len(text) < 30 or text in seen:
                continue
            seen.add(text)

            # Ссылка на конкретное сообщение
            date_el = msg.select_one("a.tgme_widget_message_date")
            msg_url = (
                date_el["href"]
                if date_el and date_el.get("href")
                else f"https://t.me/{channel}"
            )

            # Первый абзац/строка — как заголовок
            first_line = text.split("\n")[0].strip()
            title = first_line if len(first_line) >= 30 else text[:200]

            articles.append(
                {
                    "title": title[:200],
                    "content": text[:600],
                    "url": msg_url,
                    "source_name": name,
                }
            )

        log.info("Источник '%s' (@%s): найдено сообщений — %d", name, channel, len(articles))

    except requests.exceptions.Timeout:
        log.warning("Источник '%s' (@%s): таймаут запроса", name, channel)
    except requests.exceptions.ConnectionError as exc:
        log.warning("Источник '%s' (@%s): ошибка соединения — %s", name, channel, exc)
    except requests.exceptions.HTTPError as exc:
        log.warning("Источник '%s' (@%s): HTTP-ошибка — %s", name, channel, exc)
    except Exception as exc:
        log.warning("Источник '%s' (@%s): непредвиденная ошибка — %s", name, channel, exc)

    return articles


def _parse_feed_entries(text: str) -> list[dict]:
    """
    Простой парсер RSS 2.0 и Atom через stdlib xml.
    Возвращает список {'title', 'link', 'summary'}.
    """
    try:
        root = ET.fromstring(text)
    except ET.ParseError:
        return []

    tag = root.tag
    entries = []

    # RSS 2.0: <rss><channel><item>...
    items = root.findall(".//item")
    if items:
        for item in items:
            title = (item.findtext("title") or "").strip()
            link = (item.findtext("link") or "").strip()
            summary = (item.findtext("description") or "").strip()
            entries.append({"title": title, "link": link, "summary": summary})
        return entries

    # Atom: <feed xmlns="http://www.w3.org/2005/Atom">
    atom_ns = "http://www.w3.org/2005/Atom"
    ns_tag = f"{{{atom_ns}}}"
    atom_entries = root.findall(f"{ns_tag}entry")
    if not atom_entries and "feed" in tag.lower():
        atom_entries = root.findall("entry")
        ns_tag = ""

    for entry in atom_entries:
        title_el = entry.find(f"{ns_tag}title")
        title = (title_el.text or "").strip() if title_el is not None else ""

        link_el = entry.find(f"{ns_tag}link")
        if link_el is not None:
            link = link_el.get("href") or link_el.text or ""
        else:
            link = ""

        summary_el = entry.find(f"{ns_tag}summary") or entry.find(f"{ns_tag}content")
        summary = (summary_el.text or "").strip() if summary_el is not None else ""

        entries.append({"title": title, "link": link, "summary": summary})

    return entries


def _try_rss(base_url: str, name: str, session: requests.Session) -> list[dict]:
    """
    Пробует найти и распарсить RSS/Atom-фид сайта.
    Возвращает список статей или [] если RSS не найден/недоступен.
    """
    parsed = urlparse(base_url)
    origin = f"{parsed.scheme}://{parsed.netloc}"

    # Проверяем: сам URL + стандартные RSS-пути
    candidates = [base_url] + [origin + path for path in RSS_CANDIDATE_PATHS]

    for rss_url in candidates:
        try:
            resp = session.get(rss_url, timeout=15, verify=False)
            if not resp.ok:
                continue

            ct = resp.headers.get("content-type", "")
            text = resp.text.strip()
            is_feed = (
                "xml" in ct
                or "rss" in ct
                or "atom" in ct
                or text.startswith("<?xml")
                or "<rss" in text[:500]
                or "<feed" in text[:500]
            )
            if not is_feed:
                continue

            feed_entries = _parse_feed_entries(text)
            if not feed_entries:
                continue

            articles = []
            for entry in feed_entries[:MAX_ARTICLES_PER_SOURCE]:
                title = (entry.get("title") or "").strip()
                if not title or len(title) < 10:
                    continue

                # summary может содержать HTML — чистим
                raw_summary = entry.get("summary") or entry.get("description") or ""
                summary_text = BeautifulSoup(raw_summary, "lxml").get_text(
                    separator=" ", strip=True
                )[:600]

                articles.append(
                    {
                        "title": title,
                        "content": summary_text,
                        "url": entry.get("link") or base_url,
                        "source_name": name,
                    }
                )

            if articles:
                log.info(
                    "Источник '%s': RSS найден (%s), статей: %d",
                    name, rss_url, len(articles),
                )
                return articles

        except Exception:
            continue  # Этот кандидат не подошёл — пробуем следующий

    return []


def _try_html(url: str, name: str, session: requests.Session) -> list[dict]:
    """
    Парсит HTML-страницу сайта: пробует CSS-селекторы, затем заголовки h2/h3/h4.
    """
    articles = []

    try:
        resp = session.get(url, timeout=30, verify=False)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or "utf-8"
        soup = BeautifulSoup(resp.text, "lxml")

        for tag in soup(["script", "style", "nav", "footer", "header", "aside", "noscript"]):
            tag.decompose()

        # Перебираем CSS-селекторы
        news_blocks: list = []
        for selector in HTML_NEWS_SELECTORS:
            found = soup.select(selector)
            if len(found) >= 2:  # Хотя бы 2 блока — похоже на список новостей
                news_blocks = found[:MAX_ARTICLES_PER_SOURCE]
                log.debug("Источник '%s': HTML-селектор '%s' (%d блоков)", name, selector, len(found))
                break

        # Fallback: любые заголовки со ссылкой
        if not news_blocks:
            for heading in soup.find_all(["h2", "h3", "h4"]):
                if heading.find("a"):
                    news_blocks.append(heading)
            news_blocks = news_blocks[:MAX_ARTICLES_PER_SOURCE]

        seen_titles: set[str] = set()
        for block in news_blocks:
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

    except requests.exceptions.Timeout:
        log.warning("Источник '%s': таймаут запроса HTML (30 сек)", name)
    except requests.exceptions.ConnectionError as exc:
        log.warning("Источник '%s': ошибка соединения HTML — %s", name, exc)
    except requests.exceptions.HTTPError as exc:
        log.warning("Источник '%s': HTTP-ошибка HTML — %s", name, exc)
    except Exception as exc:
        log.warning("Источник '%s': непредвиденная ошибка HTML — %s", name, exc)

    return articles


def fetch_site_articles(source: dict, session: requests.Session) -> list[dict]:
    """
    Получает статьи с сайта: сначала пробует RSS, потом HTML-парсинг.
    Возвращает до MAX_ARTICLES_PER_SOURCE статей.
    """
    url = str(source.get("url", "")).strip()
    name = str(source.get("name", url))

    # Шаг 1: RSS (надёжнее, структурированные данные)
    articles = _try_rss(url, name, session)
    if articles:
        return articles

    # Шаг 2: HTML-парсинг как fallback
    log.debug("Источник '%s': RSS не найден, пробую HTML-парсинг…", name)
    articles = _try_html(url, name, session)
    log.info("Источник '%s': найдено элементов (HTML) — %d", name, len(articles))
    return articles


# ---------------------------------------------------------------------------
# Сбор статей со всех источников (без фильтрации)
# ---------------------------------------------------------------------------


def collect_all_articles(sources: list[dict]) -> list[dict]:
    """
    Обходит все активные источники, собирает до MAX_ARTICLES_PER_SOURCE
    статей с каждого сайта. Никакой тематической фильтрации — это делает Gemini.
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

        all_articles.extend(articles)

    log.info("Всего собрано статей со всех источников: %d", len(all_articles))
    return all_articles


# ---------------------------------------------------------------------------
# Фильтрация релевантных статей через Gemini API
# ---------------------------------------------------------------------------


def _parse_relevant_ids(response_text: str, max_id: int) -> list[int]:
    """
    Извлекает список relevant_ids из ответа Gemini.
    Обрабатывает варианты: чистый JSON, JSON в ```-блоке, частично сломанный ответ.
    """
    # Убираем markdown-обёртку если есть
    clean = re.sub(r"```(?:json)?\s*|\s*```", "", response_text).strip()

    try:
        data = json.loads(clean)
        ids = data.get("relevant_ids", [])
        # Оставляем только корректные числовые ID в допустимом диапазоне
        return [int(i) for i in ids if isinstance(i, (int, float)) and 1 <= int(i) <= max_id]
    except (json.JSONDecodeError, ValueError, TypeError) as exc:
        log.warning("Не удалось разобрать JSON от Gemini: %s | Ответ: %s", exc, clean[:200])
        return []


FILTER_BATCH_SIZE = 80  # Максимум заголовков в одном запросе к Gemini

# Ключевые слова для пре-фильтрации (до Gemini) и fallback (если Gemini недоступен)
RELEVANCE_KEYWORDS = [
    "земельн",
    "водн",
    "гтс",
    "пруд",
    "кадастр",
    "росреестр",
    "росводресурс",
    "рослесхоз",
    "роснедр",
    "минприрод",
    "росприроднадзор",
    "аренда земли",
    "межеван",
    "сервитут",
    "водоохранн",
    "экологическ",
    "лесфонд",
    "недр",
    "водопользован",
    "водоём",
    "водоем",
    "береговая полоса",
    "гидротехнич",
]
STOP_WORDS = ["убийство", "теракт", "наркотики"]


def pre_filter_by_keywords(articles: list[dict]) -> list[dict]:
    """
    Пре-фильтрация по ключевым словам.
    Используется до отправки в Gemini (сокращает батч) и как fallback когда Gemini недоступен.
    Статья проходит если содержит хотя бы одно ключевое слово в title+content (без учёта регистра)
    и не содержит жёстких стоп-слов.
    """
    result = []
    for article in articles:
        text = (
            (article.get("title") or "") + " " + (article.get("content") or "")
        ).lower()
        has_stop = any(sw in text for sw in STOP_WORDS)
        has_keyword = any(kw in text for kw in RELEVANCE_KEYWORDS)
        if has_keyword and not has_stop:
            result.append(article)
    return result


def filter_relevant_with_gemini(articles: list[dict]) -> list[dict]:
    """
    Отправляет заголовки статей в Gemini для фильтрации.
    Если заголовков > FILTER_BATCH_SIZE — разбивает на батчи по 80 штук,
    обрабатывает последовательно с паузой 2 сек между батчами.
    При недоступности Gemini API — fallback на keyword-фильтрацию.
    Возвращает отфильтрованный список статей (максимум MAX_POSTS_TO_GENERATE).
    """
    if not articles:
        return []

    # Разбиваем на батчи
    batches = [
        articles[i: i + FILTER_BATCH_SIZE]
        for i in range(0, len(articles), FILTER_BATCH_SIZE)
    ]
    total_batches = len(batches)
    log.info(
        "Отправляю %d заголовков в Gemini для фильтрации (батчей: %d)…",
        len(articles), total_batches,
    )

    all_relevant: list[dict] = []
    gemini_failed_batches = 0

    for batch_num, batch in enumerate(batches, start=1):
        # Нумерация внутри батча начинается с 1 — IDs локальные
        headlines_lines = []
        for idx, article in enumerate(batch, start=1):
            source = article.get("source_name", "")
            title = article.get("title", "").strip()
            headlines_lines.append(f"{idx}. [{source}] {title}")

        headlines_list = "\n".join(headlines_lines)
        prompt = GEMINI_FILTER_PROMPT.format(headlines_list=headlines_list)

        log.info("Батч %d/%d: %d заголовков", batch_num, total_batches, len(batch))

        try:
            response_text = gemini_generate(prompt)
            log.debug("Ответ Gemini (батч %d): %s", batch_num, response_text[:300])
        except Exception as exc:
            log.error("Ошибка Gemini API при фильтрации батча %d: %s", batch_num, exc)
            gemini_failed_batches += 1
            if batch_num < total_batches:
                time.sleep(2)
            continue

        local_ids = _parse_relevant_ids(response_text, max_id=len(batch))
        batch_relevant = [batch[i - 1] for i in local_ids]
        all_relevant.extend(batch_relevant)

        log.info(
            "Батч %d/%d: выбрано %d релевантных | IDs: %s",
            batch_num, total_batches, len(batch_relevant), local_ids,
        )

        if batch_num < total_batches:
            time.sleep(2)

    # Если Gemini не ответил ни на один батч — используем keyword-fallback
    if gemini_failed_batches == total_batches:
        log.warning(
            "Gemini API недоступен (все %d батчей упали) — "
            "переключаюсь на keyword-фильтрацию",
            total_batches,
        )
        fallback = pre_filter_by_keywords(articles)
        log.info("Keyword-fallback: %d релевантных из %d", len(fallback), len(articles))
        return fallback[:MAX_POSTS_TO_GENERATE]

    if not all_relevant:
        log.info("Gemini не нашёл релевантных новостей")
        return []

    log.info("Итого релевантных: %d из %d", len(all_relevant), len(articles))
    return all_relevant[:MAX_POSTS_TO_GENERATE]


# ---------------------------------------------------------------------------
# Генерация постов через Gemini API
# ---------------------------------------------------------------------------


def generate_post(article: dict) -> str:
    """Генерирует готовый пост через Google Gemini REST API."""
    prompt = GEMINI_POST_PROMPT.format(
        title=article.get("title", "Без заголовка"),
        content=(article.get("content") or "")[:2000],
        source_name=article.get("source_name", ""),
        url=article.get("url", ""),
    )

    log.info("Генерирую пост: «%s»", (article.get("title") or "")[:70])
    return gemini_generate(prompt)


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

    # 1. Проверяем наличие Gemini API ключа сразу — он нужен для обоих шагов
    if not GEMINI_API_KEY:
        log.error(
            "GEMINI_API_KEY не установлен. "
            "Добавьте ключ в .env или GitHub Secrets."
        )
        sys.exit(1)

    log.info("Используется модель: %s", GEMINI_MODEL)

    # 2. Читаем источники
    sources = read_sources()
    if not sources:
        log.warning("Нет активных источников в sources.xlsx — выходим")
        save_last_check(now_utc)
        return

    # 3. Собираем все статьи без фильтрации по ключевым словам
    all_articles = collect_all_articles(sources)

    if not all_articles:
        log.info("Ни одной статьи не собрано ни с одного источника")
        content = build_output([], now_msk, len(sources))
        save_output(content, now_msk)
        save_last_check(now_utc)
        return

    # 4. Пре-фильтрация по ключевым словам — отсекаем заведомо нерелевантное до Gemini
    prefiltered = pre_filter_by_keywords(all_articles)
    log.info(
        "Пре-фильтрация по ключевым словам: %d -> %d статей",
        len(all_articles), len(prefiltered),
    )

    if not prefiltered:
        log.info("После keyword-фильтрации статей не осталось — нет релевантных материалов")
        content = build_output([], now_msk, len(sources))
        save_output(content, now_msk)
        save_last_check(now_utc)
        return

    # 5. Финальная фильтрация через Gemini (работает только с уже отфильтрованным пулом)
    relevant = filter_relevant_with_gemini(prefiltered)

    if not relevant:
        log.info("Релевантных материалов не найдено по оценке Gemini")
        content = build_output([], now_msk, len(sources))
        save_output(content, now_msk)
        save_last_check(now_utc)
        return

    # 6. Генерируем посты для релевантных статей
    articles_with_posts: list[tuple] = []

    for article in relevant:
        try:
            post_text = generate_post(article)
            articles_with_posts.append((article, post_text))
            time.sleep(2)  # Пауза между вызовами Gemini API
        except Exception as exc:
            log.error(
                "Ошибка генерации поста для «%s»: %s",
                (article.get("title") or "")[:60],
                exc,
            )

    # 6. Сохраняем результат
    content = build_output(articles_with_posts, now_msk, len(sources))
    save_output(content, now_msk)
    save_last_check(now_utc)

    log.info("=" * 60)
    log.info("Мониторинг завершён. Сформировано постов: %d", len(articles_with_posts))
    log.info("=" * 60)


if __name__ == "__main__":
    main()
