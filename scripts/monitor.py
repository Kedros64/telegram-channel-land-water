#!/usr/bin/env python3
"""
monitor.py — Мониторинг источников и генерация постов для Telegram-канала
о земельном и водном праве России.

Запуск:
    python scripts/monitor.py

Переменные окружения:
    DEEPSEEK_API_KEY — ключ API DeepSeek
"""

import base64
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
from openai import OpenAI

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
ARTICLES_CACHE_FILE = POSTS_DIR / "articles_cache.json"
ARTICLES_POOL_FILE  = POSTS_DIR / "articles_pool.json"   # Накопительный пул статей
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_MODEL = "deepseek-chat"
DEEPSEEK_API_BASE = "https://api.deepseek.com"

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_IMAGE_MODEL = "dall-e-3"

MAX_ARTICLES_PER_SOURCE = 10  # Сколько заголовков брать с каждого сайта

# Максимум обычных постов в день (6 регулярных + 1 саммари = 7 запусков)
MAX_REGULAR_POSTS_PER_DAY = 6
SUMMARY_POST_ARTICLES = 5  # Сколько новостей включать в вечерний саммари

# Статьи старше этого числа дней удаляются из пула (если уже использованы)
POOL_MAX_AGE_DAYS = 14

# Слова-маркеры регионального характера новости (для приоритизации федеральных)
REGIONAL_MARKERS = [
    "область", "край", "республика", "округ", "район",
    "город ", "села ", "посёлок", "поселок", "деревня",
    "губернатор", "мэр ", "глава района", "глава города",
    "областн", "краев", "муниципальн", "региональн",
]

# Ключевые слова для классификации статей по тематике
WATER_CATEGORY_KEYWORDS = [
    "водн", "гтс", "пруд", "водоём", "водоем",
    "береговая полоса", "водоохранн", "водопользован",
    "гидротехнич", "росводресурс",
]
LAND_CATEGORY_KEYWORDS = [
    "земельн", "кадастр", "росреестр", "аренда земли",
    "межеван", "сервитут", "лесфонд", "рослесхоз",
    "роснедр", "недр", "росприроднадзор", "минприрод", "экологическ",
]

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
# Промпты для DeepSeek API
# ---------------------------------------------------------------------------

DEEPSEEK_FILTER_PROMPT = """\
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

DEEPSEEK_POST_PROMPT = """\
Ты — редактор Telegram-канала о земельном и водном праве России.
Канал читают предприниматели, фермеры, арендаторы, владельцы участков и водоёмов.
Автор канала — консультант по оформлению водопользования, ГТС, прудов, земельных участков, участков в лесфонде, оформлению водозабора и водосброса.

ВАЖНО ПРО ТОН: Мы на стороне читателя. Мы присоединяемся к его эмоциям и переживаниям. Читатель должен чувствовать: «эти ребята понимают мою боль, они со мной заодно и готовы помочь». Не поучаем, не свысока — а как коллега, который тоже через это прошёл.

Вот исходный материал:
ЗАГОЛОВОК: {title}
ТЕКСТ/АНОНС: {content}
ИСТОЧНИК: {source_name}
ССЫЛКА: {url}

Сформируй готовый пост для Telegram строго по структуре ниже.

**ГОТОВЫЙ ПОСТ:**
[Пост 800–1500 символов с пробелами.

СТРУКТУРА:
1. ЗАГОЛОВОК-ХУК — первая строка поста, выделена жирным (**жирный**). Одна короткая фраза или предложение, которое цепляет внимание и отражает суть. Никаких вводных слов перед ним.
2. Суть материала простым языком — что произошло и что это значит для читателя на практике. Разбивай на абзацы. Если уместно — используй маркеры или нумерованный список.
3. Если упомянут нормативный акт — обязательно назови его.
4. Если материал про конкретный регион — упомяни, но подчеркни как это касается всех.
5. Присоединение к эмоциям читателя — покажи что ты понимаешь его чувства (раздражение, тревогу, недоумение). Дай понять что мы на одной стороне.
6. Вывод с лёгкой иронией над ситуацией или бюрократией — дружелюбной, «подмигивающей», не злой и не токсичной.
7. Один CTA в конце — ненавязчивое упоминание что автор канала помогает с оформлением водопользования, ГТС, земельных участков, лесфонда, водозабора и водосброса.

ОФОРМЛЕНИЕ:
- Добавляй уместные смайлики по тексту для усиления эмоций и смысловых акцентов.
- Важные мысли и ключевые фразы выделяй жирным (**жирный**). Жирного не больше 20–25% текста.
- Ирония и самоирония приветствуются — лёгкие, неожиданные формулировки, без грубости.
- Тон: живой, немного хулиганский, без канцелярита, без хамства. Мы — свои.
- Никаких юридических гарантий — только «как правило», «по практике», «есть риск что».]

**ВАРИАНТЫ CTA:**
Нейтральный: [вариант]
С юмором: [вариант]
Прямой: [вариант]

**ВИЗУАЛ:**
Prompt (EN): [Photorealistic image, natural lighting, high detail, realistic colors, NO TEXT, NO LETTERING, NO CAPTIONS, NO WORDS anywhere in the image. Slightly humorous scene related to the topic — gentle irony or unexpected situation. Описывай атмосферу и контекст поста. Если в кадре люди — только нейтральные собирательные образы, без реальных известных личностей.]
Описание (RU): [2–3 слова]
Запасной вариант: [более простой промпт в том же фотореалистичном стиле, NO TEXT, NO LETTERING]
"""

DEEPSEEK_HUMOROUS_POST_PROMPT = """\
Ты — редактор Telegram-канала о земельном и водном праве России.
Канал читают предприниматели, фермеры, арендаторы, владельцы участков и водоёмов.
Автор канала — консультант по оформлению водопользования, ГТС, прудов, земельных участков, участков в лесфонде, оформлению водозабора и водосброса.

ВАЖНО: Это ЮМОРИСТИЧЕСКИЙ пост, который выходит в обеденное время или в конце рабочего дня. Задача — разрядить людей, вызвать улыбку. Больше сарказма, иронии и остроумия чем обычно. Но при этом мы — на стороне читателя, мы понимаем его боль и подшучиваем над бюрократией, а не над людьми.

Вот исходный материал:
ЗАГОЛОВОК: {title}
ТЕКСТ/АНОНС: {content}
ИСТОЧНИК: {source_name}
ССЫЛКА: {url}

Сформируй готовый пост для Telegram строго по структуре ниже.

**ГОТОВЫЙ ПОСТ:**
[Пост 800–1500 символов с пробелами.

СТРУКТУРА:
1. ЗАГОЛОВОК-ХУК — первая строка, выделена жирным (**жирный**). Ироничная, цепляющая фраза с сарказмом. Может быть неожиданным наблюдением или абсурдным выводом из ситуации.
2. Суть материала простым языком — но подано с юмором и сарказмом. Больше иронических отступлений и наблюдений. Допускаются гиперболы, риторические вопросы, неожиданные сравнения.
3. Если упомянут нормативный акт — обязательно назови его (можно с ироничным комментарием).
4. Если материал про конкретный регион — упомяни, можно пошутить про «один регион, зато какой».
5. Саркастический вывод — остроумное резюме ситуации. Мы посмеиваемся вместе с читателем.
6. Один CTA в конце — можно с юмором, но с упоминанием что автор помогает с оформлением.

ОФОРМЛЕНИЕ:
- Смайлики — щедрее чем обычно, для усиления комичности.
- Важные фразы жирным (**жирный**). Жирного не больше 25%.
- Максимум сарказма и иронии, но без злости, хамства и оскорблений.
- Подшучиваем над СИТУАЦИЯМИ и бюрократией, НЕ над людьми и организациями напрямую.
- Тон: как друг рассказывает анекдот из своей практики — с хохотком и сочувствием.
- Никаких юридических гарантий.]

**ВАРИАНТЫ CTA:**
Нейтральный: [вариант]
С юмором: [вариант]
Прямой: [вариант]

**ВИЗУАЛ:**
Prompt (EN): [Photorealistic image, natural lighting, high detail, realistic colors, NO TEXT, NO LETTERING, NO CAPTIONS, NO WORDS. HUMOROUS scene — exaggerated or absurd situation related to the topic. Visual irony, unexpected juxtaposition. Gentle comedy, not offensive. Если в кадре люди — нейтральные собирательные образы.]
Описание (RU): [2–3 слова]
Запасной вариант: [более простой юмористический промпт, NO TEXT, NO LETTERING]
"""

DEEPSEEK_SUMMARY_POST_PROMPT = """\
Ты — редактор Telegram-канала о земельном и водном праве России.
Канал читают предприниматели, фермеры, арендаторы, владельцы участков и водоёмов.
Автор канала — консультант по оформлению водопользования, ГТС, прудов, земельных участков и лесфонда.

Это ВЕЧЕРНИЙ САММАРИ — дайджест за день. Задача — дать читателю быстрый обзор главных новостей дня, чтобы он мог за 1 минуту понять что произошло важного.

Вот {count} новостей для саммари:
{articles_block}

Сформируй саммари-пост для Telegram строго по структуре ниже.

**ГОТОВЫЙ ПОСТ:**
[Пост 1200–2000 символов с пробелами.

СТРУКТУРА:
1. ЗАГОЛОВОК: **📋 Дайджест дня: земля и вода** (или похожий, с эмодзи)
2. Краткое вступление (1 предложение) — «Что произошло сегодня в мире земли и воды»
3. Для каждой из {count} новостей — пронумерованный блок:
   - Жирный подзаголовок (**заголовок**)
   - 2–3 предложения: суть и что это значит на практике
   - Если есть нормативный акт — упомяни
4. Краткое заключение (1–2 предложения) с лёгкой иронией
5. CTA — ненавязчивое упоминание услуг автора

ОФОРМЛЕНИЕ:
- Каждую новость отделяй пустой строкой
- Смайлики к каждой новости для быстрого сканирования
- Жирный для подзаголовков
- Присоединение к эмоциям читателя — «мы с вами в одной лодке»
- Тон: живой, дружелюбный, без канцелярита]

**ВАРИАНТЫ CTA:**
Нейтральный: [вариант]
С юмором: [вариант]
Прямой: [вариант]

**ВИЗУАЛ:**
Prompt (EN): [Photorealistic collage-style image divided into {count} panels, each panel showing a different scene related to one of the topics. Natural lighting, realistic colors, NO TEXT, NO LETTERING, NO CAPTIONS, NO WORDS anywhere. Slightly humorous tone. Panels separated by thin lines.]
Описание (RU): [2–3 слова]
Запасной вариант: [single photorealistic scene combining elements of land, water, and bureaucracy themes, NO TEXT, NO LETTERING]
"""

# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------


def deepseek_generate(prompt: str) -> str:
    """
    Вызывает DeepSeek API через openai-совместимый клиент.
    При ошибке 429 делает до 3 повторных попыток с задержкой 5/10/20 сек.
    """
    client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_API_BASE)
    delays = [5, 10, 20]
    last_exc = None
    for attempt in range(1, len(delays) + 2):  # попытки 1..4
        try:
            response = client.chat.completions.create(
                model=DEEPSEEK_MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.7,
                max_tokens=2048,
            )
            return response.choices[0].message.content
        except Exception as exc:
            err_str = str(exc)
            if "429" in err_str and attempt <= len(delays):
                wait = delays[attempt - 1]
                log.warning(
                    "DeepSeek API: 429 Too Many Requests (попытка %d/4), жду %d сек…",
                    attempt, wait,
                )
                time.sleep(wait)
                last_exc = exc
                continue
            raise
    raise RuntimeError(f"DeepSeek API: исчерпаны все попытки. Последняя ошибка: {last_exc}")


def extract_image_prompt(post_text: str) -> str:
    """
    Извлекает английский промпт из секции **ВИЗУАЛ:** сгенерированного поста.
    Возвращает строку или '' если секция отсутствует.
    """
    match = re.search(r"\*{0,2}Prompt \(EN\):\*{0,2}\s*(.+?)(?:\n|$)", post_text)
    if match:
        return match.group(1).strip().strip("[]")
    return ""


def generate_image(prompt: str, now_msk: datetime, post_num: int) -> Path | None:
    """
    Генерирует изображение через OpenAI Images API (gpt-image-1-mini).
    Сохраняет PNG в posts/YYYY-MM-DD_HH_post_N.png.
    При любой ошибке логирует предупреждение и возвращает None —
    пост в любом случае будет опубликован, просто без картинки.
    """
    if not OPENAI_API_KEY:
        log.debug("OPENAI_API_KEY не задан — генерация изображений пропущена")
        return None

    log.info("Генерирую изображение для поста %d…", post_num)
    try:
        resp = requests.post(
            "https://api.openai.com/v1/images/generations",
            headers={
                "Authorization": f"Bearer {OPENAI_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": OPENAI_IMAGE_MODEL,
                "prompt": prompt[:1000],  # API limit
                "size": "1024x1024",
                "quality": "standard",
                "response_format": "b64_json",
                "n": 1,
            },
            timeout=90,
        )
        resp.raise_for_status()
        data = resp.json()

        b64_data = data["data"][0].get("b64_json")
        if not b64_data:
            log.warning("OpenAI Images API не вернул b64_json для поста %d", post_num)
            return None

        img_bytes = base64.b64decode(b64_data)
        filename = now_msk.strftime("%Y-%m-%d_%H") + f"_post_{post_num}.png"
        img_path = POSTS_DIR / filename
        img_path.write_bytes(img_bytes)
        log.info("Изображение сохранено: %s (%d KB)", filename, len(img_bytes) // 1024)
        return img_path

    except requests.exceptions.HTTPError as exc:
        log.warning("OpenAI Images API HTTP-ошибка (пост %d): %s", post_num, exc)
    except requests.exceptions.Timeout:
        log.warning("OpenAI Images API: таймаут (пост %d)", post_num)
    except Exception as exc:
        log.warning("Не удалось сгенерировать изображение (пост %d): %s", post_num, exc)
    return None


def moscow_now() -> datetime:
    """Текущее время по Москве (UTC+3)."""
    return datetime.now(timezone.utc) + MOSCOW_OFFSET


def detect_post_type(now_msk: datetime) -> str:
    """
    Определяет тип поста по текущему времени МСК.
    Возвращает: "regular", "humorous" или "summary".
    Если время не совпадает ни с одним слотом (ручной запуск) — "regular".
    """
    hour, minute = now_msk.hour, now_msk.minute
    # 13:10 и 19:30 — юмористические
    if (hour == 13 and 0 <= minute <= 20) or (hour == 19 and 20 <= minute <= 40):
        return "humorous"
    # 20:10 — вечерний саммари
    if hour == 20 and 0 <= minute <= 20:
        return "summary"
    return "regular"


def is_regional(article: dict) -> bool:
    """
    Проверяет, является ли статья узко-региональной.
    Региональные новости получают меньший приоритет при выборе.
    """
    text = (
        (article.get("title") or "") + " " + (article.get("content") or "")
    ).lower()
    regional_count = sum(1 for marker in REGIONAL_MARKERS if marker in text)
    return regional_count >= 2


def validate_url(url: str, session: requests.Session | None = None) -> bool:
    """
    Проверяет работоспособность URL (HEAD-запрос с fallback на GET).
    Возвращает True если URL отвечает 2xx/3xx.
    """
    if not url or not url.startswith("http"):
        return False
    try:
        s = session or requests.Session()
        resp = s.head(url, timeout=5, verify=False, allow_redirects=True)
        if resp.ok:
            return True
        # Некоторые серверы не поддерживают HEAD — пробуем GET
        resp = s.get(url, timeout=5, verify=False, stream=True)
        return resp.ok
    except Exception:
        return False


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
# Фильтрация релевантных статей через DeepSeek API
# ---------------------------------------------------------------------------


def _parse_relevant_ids(response_text: str, max_id: int) -> list[int]:
    """
    Извлекает список relevant_ids из ответа DeepSeek.
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
        log.warning("Не удалось разобрать JSON от DeepSeek: %s | Ответ: %s", exc, clean[:200])
        return []


FILTER_BATCH_SIZE = 80  # Максимум заголовков в одном запросе к DeepSeek

# Ключевые слова для пре-фильтрации (до DeepSeek) и fallback (если DeepSeek недоступен)
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
    Используется до отправки в DeepSeek (сокращает батч) и как fallback когда DeepSeek недоступен.
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


# ---------------------------------------------------------------------------
# Пул статей: классификация, накопление, выбор для генерации
# ---------------------------------------------------------------------------


def classify_article(article: dict) -> str:
    """
    Классифицирует статью как 'land', 'water' или 'both'.
    Возвращает 'both' если содержит ключевые слова обеих тематик.
    Fallback на 'land' если ни один водный ключ не найден.
    """
    text = (
        (article.get("title") or "") + " " + (article.get("content") or "")
    ).lower()
    has_water = any(kw in text for kw in WATER_CATEGORY_KEYWORDS)
    has_land  = any(kw in text for kw in LAND_CATEGORY_KEYWORDS)
    if has_water and has_land:
        return "both"
    if has_water:
        return "water"
    return "land"  # по умолчанию — земельная тематика


def load_articles_pool() -> list[dict]:
    """Загружает накопительный пул статей из JSON-файла."""
    if not ARTICLES_POOL_FILE.exists():
        return []
    try:
        return json.loads(ARTICLES_POOL_FILE.read_text(encoding="utf-8"))
    except Exception as exc:
        log.warning("Не удалось загрузить пул статей: %s", exc)
        return []


def save_articles_pool(pool: list[dict]) -> None:
    """Сохраняет пул статей в JSON-файл."""
    POSTS_DIR.mkdir(exist_ok=True)
    ARTICLES_POOL_FILE.write_text(
        json.dumps(pool, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def merge_into_pool(
    pool: list[dict], new_articles: list[dict], now_msk: datetime
) -> list[dict]:
    """
    Добавляет новые статьи в пул (дедупликация по URL).
    Удаляет статьи, которые: уже использованы И старше POOL_MAX_AGE_DAYS дней.
    Неиспользованные статьи хранятся пока не будут опубликованы.
    """
    existing_urls = {a.get("url", "") for a in pool}
    added = 0
    for article in new_articles:
        url = article.get("url", "")
        if not url or url in existing_urls:
            continue
        pool.append({
            **article,
            "collected_at": now_msk.isoformat(),
            "category": classify_article(article),
            "used": False,
        })
        existing_urls.add(url)
        added += 1

    # Очистка: удаляем использованные статьи старше POOL_MAX_AGE_DAYS
    cutoff_str = (now_msk - timedelta(days=POOL_MAX_AGE_DAYS)).isoformat()
    before = len(pool)
    pool = [
        a for a in pool
        if not a.get("used") or a.get("collected_at", "") > cutoff_str
    ]
    purged = before - len(pool)

    log.info(
        "Пул: добавлено %d новых, удалено %d устаревших, итого %d статей (%d неиспользованных)",
        added, purged, len(pool),
        sum(1 for a in pool if not a.get("used")),
    )
    return pool


def get_today_used_count(pool: list[dict], today_str: str) -> int:
    """Считает сколько постов уже выбрано/опубликовано сегодня."""
    return sum(
        1 for a in pool
        if a.get("used_at", "").startswith(today_str)
    )


def _sort_by_priority(articles: list[dict]) -> list[dict]:
    """
    Сортирует статьи по приоритету:
    1. Федеральные (не региональные) — выше
    2. Свежие — выше (collected_at desc)
    Двухпроходная стабильная сортировка: сначала по дате desc, затем по региональности asc.
    """
    # Проход 1: свежие первыми
    result = sorted(articles, key=lambda a: a.get("collected_at", ""), reverse=True)
    # Проход 2: федеральные (regional=False) первыми, региональные — в конце
    result = sorted(result, key=lambda a: is_regional(a))
    return result


def select_posts_from_pool(
    pool: list[dict], today_str: str, count: int = 1
) -> list[dict]:
    """
    Выбирает N неиспользованных статей из пула.
    Приоритет: федеральные новости > региональные, свежие > старые.
    Возвращает [] если пул пуст или дневной лимит исчерпан.
    """
    used_today = get_today_used_count(pool, today_str)
    remaining = MAX_REGULAR_POSTS_PER_DAY - used_today
    log.info("Сегодня использовано: %d/%d постов", used_today, MAX_REGULAR_POSTS_PER_DAY)

    if remaining <= 0 and count == 1:
        # Лимит применяется только к обычным постам (count=1).
        # Вечерний саммари (count=SUMMARY_POST_ARTICLES) выходит независимо от лимита.
        log.info("Суточный лимит обычных постов исчерпан")
        return []

    # Неиспользованные, отсортированные по приоритету
    unused = [a for a in pool if not a.get("used")]
    if not unused:
        log.warning("Пул статей пуст — нет материалов для поста")
        return []

    # Сортируем: федеральные первыми, потом по свежести
    prioritized = _sort_by_priority(unused)

    selected = prioritized[:count]
    log.info(
        "Выбрано %d статей из пула (запрошено %d): %s",
        len(selected), count,
        [a.get("title", "")[:50] for a in selected],
    )
    return selected


def filter_relevant_with_deepseek(articles: list[dict]) -> list[dict]:
    """
    Отправляет заголовки статей в DeepSeek для фильтрации.
    Если заголовков > FILTER_BATCH_SIZE — разбивает на батчи по 80 штук,
    обрабатывает последовательно с паузой 2 сек между батчами.
    При недоступности DeepSeek API — fallback на keyword-фильтрацию.
    Возвращает все найденные релевантные статьи (без ограничения количества).
    Ограничение по количеству постов накладывается позже при выборе из пула.
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
        "Отправляю %d заголовков в DeepSeek для фильтрации (батчей: %d)…",
        len(articles), total_batches,
    )

    all_relevant: list[dict] = []
    deepseek_failed_batches = 0

    for batch_num, batch in enumerate(batches, start=1):
        # Нумерация внутри батча начинается с 1 — IDs локальные
        headlines_lines = []
        for idx, article in enumerate(batch, start=1):
            source = article.get("source_name", "")
            title = article.get("title", "").strip()
            headlines_lines.append(f"{idx}. [{source}] {title}")

        headlines_list = "\n".join(headlines_lines)
        prompt = DEEPSEEK_FILTER_PROMPT.format(headlines_list=headlines_list)

        log.info("Батч %d/%d: %d заголовков", batch_num, total_batches, len(batch))

        try:
            response_text = deepseek_generate(prompt)
            log.debug("Ответ DeepSeek (батч %d): %s", batch_num, response_text[:300])
        except Exception as exc:
            log.error("Ошибка DeepSeek API при фильтрации батча %d: %s", batch_num, exc)
            deepseek_failed_batches += 1
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

    # Если DeepSeek не ответил ни на один батч — используем keyword-fallback
    if deepseek_failed_batches == total_batches:
        log.warning(
            "DeepSeek API недоступен (все %d батчей упали) — "
            "переключаюсь на keyword-фильтрацию",
            total_batches,
        )
        fallback = pre_filter_by_keywords(articles)
        log.info("Keyword-fallback: %d релевантных из %d", len(fallback), len(articles))
        return fallback

    if not all_relevant:
        log.info("DeepSeek не нашёл релевантных новостей")
        return []

    log.info("Итого релевантных: %d из %d", len(all_relevant), len(articles))
    return all_relevant


# ---------------------------------------------------------------------------
# Генерация постов через DeepSeek API
# ---------------------------------------------------------------------------


def generate_post(article: dict, post_type: str = "regular") -> str:
    """
    Генерирует готовый пост через DeepSeek API.
    post_type: "regular" или "humorous" — определяет промпт.
    """
    template = DEEPSEEK_HUMOROUS_POST_PROMPT if post_type == "humorous" else DEEPSEEK_POST_PROMPT
    prompt = template.format(
        title=article.get("title", "Без заголовка"),
        content=(article.get("content") or "")[:2000],
        source_name=article.get("source_name", ""),
        url=article.get("url", ""),
    )

    log.info("Генерирую %s пост: «%s»", post_type, (article.get("title") or "")[:70])
    return deepseek_generate(prompt)


def generate_summary_post(articles: list[dict]) -> str:
    """Генерирует вечерний саммари-пост по нескольким статьям."""
    articles_block = ""
    for i, a in enumerate(articles, start=1):
        articles_block += (
            f"\n{i}. ЗАГОЛОВОК: {a.get('title', 'Без заголовка')}\n"
            f"   ТЕКСТ/АНОНС: {(a.get('content') or '')[:500]}\n"
            f"   ИСТОЧНИК: {a.get('source_name', '')}\n"
            f"   ССЫЛКА: {a.get('url', '')}\n"
        )

    prompt = DEEPSEEK_SUMMARY_POST_PROMPT.format(
        count=len(articles),
        articles_block=articles_block,
    )

    log.info("Генерирую саммари-пост по %d новостям", len(articles))
    return deepseek_generate(prompt)


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

    for i, (article, post_text, image_path) in enumerate(articles_with_posts, start=1):
        lines += [
            "---",
            "",
            f"## Пост {i}",
            "",
            f"**Источник:** {article.get('source_name', '')} — {article.get('url', '')}",
            f"**Суть:** {article.get('title', '')}",
        ]
        if image_path:
            lines.append(f"**ИЗОБРАЖЕНИЕ:** {image_path.name}")
        lines += [
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
    post_type = detect_post_type(now_msk)

    log.info("Текущее время (МСК): %s", now_msk.strftime("%Y-%m-%d %H:%M"))
    log.info("Последняя проверка:  %s", last_check.isoformat())
    log.info("Тип поста: %s", post_type)

    # 1. Проверяем наличие DeepSeek API ключа
    if not DEEPSEEK_API_KEY:
        log.error(
            "DEEPSEEK_API_KEY не установлен. "
            "Добавьте ключ в .env или GitHub Secrets."
        )
        sys.exit(1)

    log.info("Используется модель: %s", DEEPSEEK_MODEL)

    # 2. Читаем источники
    sources = read_sources()
    if not sources:
        log.warning("Нет активных источников в sources.xlsx — выходим")
        save_last_check(now_utc)
        return

    # 3. Собираем все статьи
    all_articles = collect_all_articles(sources)

    if not all_articles:
        log.info("Ни одной статьи не собрано ни с одного источника")
        # Всё равно проверяем пул — в нём могут быть статьи от предыдущих запусков
        pool = load_articles_pool()
    else:
        # 4. Пре-фильтрация по ключевым словам
        prefiltered = pre_filter_by_keywords(all_articles)
        log.info(
            "Пре-фильтрация по ключевым словам: %d -> %d статей",
            len(all_articles), len(prefiltered),
        )

        # Сохраняем кэш для regen.py
        try:
            POSTS_DIR.mkdir(exist_ok=True)
            ARTICLES_CACHE_FILE.write_text(
                json.dumps(prefiltered, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            log.info("Кэш статей сохранён: %d записей → %s", len(prefiltered), ARTICLES_CACHE_FILE.name)
        except Exception as exc:
            log.warning("Не удалось сохранить кэш статей: %s", exc)

        # 5. Финальная фильтрация через DeepSeek
        if prefiltered:
            relevant = filter_relevant_with_deepseek(prefiltered)
        else:
            relevant = []

        # 6. Обновляем накопительный пул (дедупликация по URL)
        pool = load_articles_pool()
        pool = merge_into_pool(pool, relevant, now_msk)

    today_str = now_msk.strftime("%Y-%m-%d")

    # 7. Выбираем статьи из пула в зависимости от типа поста
    if post_type == "summary":
        # Вечерний саммари: берём 5 неиспользованных статей
        selected_list = select_posts_from_pool(pool, today_str, count=SUMMARY_POST_ARTICLES)
    else:
        # Обычный или юмористический: берём 1 статью
        selected_list = select_posts_from_pool(pool, today_str, count=1)

    if not selected_list:
        save_articles_pool(pool)
        content = build_output([], now_msk, len(sources))
        save_output(content, now_msk)
        save_last_check(now_utc)
        return

    # Отмечаем выбранные статьи как использованные
    selected_urls = {a["url"] for a in selected_list}
    for a in pool:
        if a.get("url") in selected_urls:
            a["used"] = True
            a["used_at"] = now_msk.isoformat()
    save_articles_pool(pool)

    # 8. Валидация ссылок — убираем нерабочие URL из постов
    session = make_session()
    for article in selected_list:
        url = article.get("url", "")
        if url and not validate_url(url, session):
            log.warning("Ссылка не работает, убираю из поста: %s", url)
            article["url"] = ""  # Пустая ссылка — не будет включена в пост

    # 9. Генерация постов
    articles_with_posts: list[tuple] = []

    if post_type == "summary":
        # Саммари-пост: один пост с несколькими новостями
        try:
            post_text = generate_summary_post(selected_list)

            image_path = None
            if OPENAI_API_KEY:
                img_prompt = extract_image_prompt(post_text)
                if img_prompt:
                    image_path = generate_image(img_prompt, now_msk, 1)

            # Для саммари используем первую статью как «представителя» в метаданных
            meta_article = {
                "title": f"Дайджест дня: {len(selected_list)} новостей",
                "source_name": "Саммари",
                "url": "",
            }
            articles_with_posts.append((meta_article, post_text, image_path))
        except Exception as exc:
            log.error("Ошибка генерации саммари-поста: %s", exc)
    else:
        # Обычный или юмористический — по одному посту
        for article in selected_list:
            try:
                post_text = generate_post(article, post_type=post_type)

                image_path = None
                if OPENAI_API_KEY:
                    img_prompt = extract_image_prompt(post_text)
                    if img_prompt:
                        image_path = generate_image(
                            img_prompt, now_msk, len(articles_with_posts) + 1
                        )

                articles_with_posts.append((article, post_text, image_path))
                time.sleep(2)
            except Exception as exc:
                log.error(
                    "Ошибка генерации поста для «%s»: %s",
                    (article.get("title") or "")[:60],
                    exc,
                )

    # 10. Сохраняем результат
    content = build_output(articles_with_posts, now_msk, len(sources))
    save_output(content, now_msk)
    save_last_check(now_utc)

    log.info("=" * 60)
    log.info(
        "Мониторинг завершён. Тип: %s. Сформировано постов: %d",
        post_type, len(articles_with_posts),
    )
    log.info("=" * 60)


if __name__ == "__main__":
    main()
