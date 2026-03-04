# CLAUDE.md — Руководство для AI-ассистента

## Описание проекта

**Репозиторий:** `Kedros64/telegram-channel-land-water`

Этот репозиторий обеспечивает автоматизацию **Telegram-канала о земельном и водном праве России**.

### Тематика канала

Канал публикует:
- Изменения в земельном и водном законодательстве
- Оформление водопользования, гидротехнических сооружений (ГТС), прудов, водоёмов
- Аренду и оформление земельных участков
- Судебную практику по земельным и водным спорам
- Абсурдные и показательные бюрократические кейсы из этой сферы

### Бизнес-цель

Привлечение клиентов на услуги по оформлению водопользования, гидротехнических сооружений, земельных участков и всего смежного. Автор канала — консультант в этой сфере.

### Аудитория

Предприниматели, фермеры, арендаторы, управленцы, владельцы участков и водоёмов — люди которым важно: «что это значит для меня на практике».

---

## Архитектура системы

```
sources.xlsx
     │
     ▼
scripts/monitor.py          ← читает источники (Telegram каналы + сайты),
     │                         парсит через t.me/s/{channel} или RSS/HTML,
     │  (два типа вызовов     фильтрует релевантные через Gemini (батч-запрос),
     │   Gemini API)          генерирует готовые посты (по одному запросу на статью)
     ▼
posts/ГГГГ-ММ-ДД_ЧЧ.md    ← сохраняет результат мониторинга
     │
     ▼
scripts/poster.py           ← извлекает блоки «ГОТОВЫЙ ПОСТ:»,
     │                         конвертирует Markdown → HTML,
     │                         отправляет через Telegram Bot API
     ▼
Telegram Bot API
     │
     ▼
Telegram-канал
```

Автоматизация запускается через **GitHub Actions** трижды в день:
- 09:00 МСК (06:00 UTC)
- 14:00 МСК (11:00 UTC)
- 19:00 МСК (16:00 UTC)

---

## Структура репозитория

```
telegram-channel-land-water/
├── CLAUDE.md                              # Этот файл — руководство AI-ассистента
├── sources.xlsx                           # Список источников для мониторинга
├── requirements.txt                       # Зависимости Python (5 пакетов)
├── .env.example                           # Пример переменных окружения
├── .gitignore                             # Исключения git
├── posts/                                 # Результаты мониторинга и посты
│   ├── .gitkeep                           # Чтобы папка отслеживалась git
│   ├── last_check.txt                     # Время последней проверки (UTC ISO)
│   ├── ГГГГ-ММ-ДД_ЧЧ.md                 # Результат сессии мониторинга
│   ├── ГГГГ-ММ-ДД_ЧЧ_posted.md          # Опубликованный файл (переименован)
│   ├── posted.log                         # Лог опубликованных (не в git)
│   └── errors.log                         # Лог ошибок публикации (не в git)
└── scripts/
│   ├── monitor.py                         # Мониторинг + генерация постов
│   └── poster.py                          # Публикация в Telegram
└── .github/
    └── workflows/
        └── monitor_and_post.yml           # GitHub Actions workflow
```

---

## Зависимости (requirements.txt)

Только 5 пакетов — без тяжёлых SDK:

| Пакет            | Версия  | Назначение                                      |
|------------------|---------|-------------------------------------------------|
| `openpyxl`       | ≥3.1.0  | Чтение sources.xlsx                             |
| `requests`       | ≥2.31.0 | HTTP-запросы к сайтам, Gemini API, Telegram API |
| `beautifulsoup4` | ≥4.12.0 | HTML-парсинг новостных страниц                  |
| `python-dotenv`  | ≥1.0.0  | Загрузка переменных из .env                     |
| `lxml`           | ≥5.0.0  | Быстрый HTML-парсер для BeautifulSoup           |

**Намеренно не используются:**
- `feedparser` — заменён на stdlib `xml.etree.ElementTree` для RSS/Atom парсинга
- `google-generativeai` SDK — заменён на прямые REST-вызовы через `requests`

Не добавляй эти пакеты обратно без явного обоснования.

---

## Описание скриптов

### scripts/monitor.py

**Задача:** Мониторинг источников и генерация постов через двухшаговый пайплайн Gemini.

**Алгоритм:**
1. Читает `sources.xlsx`, фильтрует строки с `active = да`
2. Для каждого источника типа `telegram` — парсит публичный веб-вид `https://t.me/s/{channel}`
3. Для каждого источника типа `site` — сначала пробует RSS (несколько стандартных путей), затем HTML-парсинг как fallback
4. Собирает все статьи со всех источников **без** предварительной фильтрации по ключевым словам
5. Отправляет все заголовки **одним** батч-запросом в Gemini → получает JSON `{"relevant_ids": [...]}`
6. Для каждой релевантной статьи (максимум `MAX_POSTS_TO_GENERATE=3`) генерирует пост через Gemini (один запрос, пауза 2 сек между вызовами)
7. Сохраняет результат в `posts/ГГГГ-ММ-ДД_ЧЧ.md`
8. Обновляет `posts/last_check.txt`

**Ключевые константы:**
```python
MAX_ARTICLES_PER_SOURCE = 10  # Сколько статей/сообщений брать с каждого источника
MAX_POSTS_TO_GENERATE = 3     # Максимум постов за один запуск
GEMINI_MODEL = "gemini-2.0-flash"
```

**Переменные окружения:**
- `GEMINI_API_KEY` — ключ Google Gemini API

**Особенности:**
- SSL-проверка отключена (`verify=False`) — госсайты часто используют самоподписанные/устаревшие сертификаты; `urllib3` warnings отключены глобально
- Пауза 1 сек между HTTP-запросами к сайтам (вежливый краулер)
- Пауза 2 сек между вызовами Gemini API при генерации постов
- User-Agent имитирует Chrome — снижает риск блокировки
- Если источник недоступен — логирует предупреждение и продолжает с остальными

**Парсинг Telegram-каналов (`fetch_telegram_channel`):**
- URL: `https://t.me/s/{channel}` — публичный веб-вид, без MTProto/API-ключей
- Имя канала извлекается из `t.me/channelname` или `@channelname`
- Селектор сообщений: `.tgme_widget_message_wrap` → `.tgme_widget_message_text`
- Берутся последние `MAX_ARTICLES_PER_SOURCE` сообщений (в reversed порядке — сначала свежие)
- Приватные/несуществующие каналы возвращают HTTP 404 → логируется предупреждение, выполнение продолжается

**Парсинг сайтов — RSS-first стратегия (`fetch_site_articles`):**
- `_try_rss()`: проверяет базовый URL + 9 стандартных путей: `/rss`, `/rss.xml`, `/feed`, `/feed.xml`, `/atom.xml`, `/news/rss`, `/news/feed`, `/export/rss`, `/lenta/rss`
- RSS-парсинг через stdlib `xml.etree.ElementTree` — поддерживает RSS 2.0 и Atom
- `_try_html()` (fallback): 20 CSS-селекторов (`article`, `.news-item`, `.card`, и др.), затем заголовки `h2/h3/h4` со ссылками
- HTML-парсинг: удаляет `script`, `style`, `nav`, `footer`, `header`, `aside`, `noscript` перед обходом

**Двухшаговый пайплайн Gemini:**
1. `filter_relevant_with_gemini()` — отправляет пронумерованные заголовки батчем, получает JSON `{"relevant_ids": [...]}`, парсит через `_parse_relevant_ids()` (устойчив к markdown-обёртке и частично сломанным ответам)
2. `generate_post()` — по одному запросу на статью, использует `GEMINI_POST_PROMPT`

**Gemini REST API (`gemini_generate`):**
- URL: `https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}`
- Параметры: `temperature=0.7`, `maxOutputTokens=2048`
- Прямой вызов через `requests.post` без SDK

---

### scripts/poster.py

**Задача:** Публикация готовых постов в Telegram-канал.

**Алгоритм:**
1. Сканирует `posts/*.md` — берёт файлы без суффикса `_posted` и не в `posted.log`
2. Извлекает блоки `**ГОТОВЫЙ ПОСТ:**` регуляркой (граница — `**ВАРИАНТЫ CTA**`, `---` или `## Пост`)
3. Конвертирует Markdown-разметку в HTML (Telegram `parse_mode=HTML`)
4. Отправляет каждый пост через `sendMessage` (Telegram Bot API)
5. При ошибке разбора HTML — повторяет без `parse_mode` (plain text)
6. Между постами из одного файла — пауза 30 секунд (`PAUSE_BETWEEN_POSTS`)
7. При успехе: переименовывает файл (`_posted.md`) и записывает в `posted.log`
8. При ошибке: записывает в `errors.log`, продолжает со следующим файлом

**Поддерживаемая Markdown → HTML конвертация (`markdown_to_html`):**
- `[текст](url)` → `<a href="url">текст</a>`
- `**текст**` → `<b>текст</b>`
- `*текст*` → `<i>текст</i>` (одиночные звёздочки)
- `` `текст` `` → `<code>текст</code>`

**Ограничения:**
- Максимум 4096 символов (лимит Telegram); длинные посты обрезаются с `...` (`sanitize_for_telegram`)
- Управляющие символы `[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]` удаляются

**Переменные окружения:**
- `BOT_TOKEN` — токен Telegram-бота
- `CHANNEL_ID` — ID или @username канала

---

## Файл sources.xlsx

Хранится в корне репозитория. Лист: **Sources**.

| Столбец  | Описание                                    |
|----------|---------------------------------------------|
| `name`   | Название источника (для логов и постов)     |
| `url`    | Ссылка: `https://site.ru/news` или `t.me/channel` |
| `type`   | `telegram` или `site`                       |
| `active` | `да` / `нет` (также: `yes`, `1`, `true`)    |

**Текущий масштаб:** ~51 активный источник (38 Telegram-каналов + 13 сайтов) — по состоянию на 2026-03-04.

Строки с `active ≠ да/yes/1/true` пропускаются без ошибок.

---

## Формат файлов posts/*.md

```markdown
# Мониторинг ГГГГ-ММ-ДД ЧЧ:ММ
Проверено источников: N
Найдено релевантных материалов: M

---

## Пост 1

**Источник:** Название источника — https://example.com/article
**Суть:** Заголовок материала

**ГОТОВЫЙ ПОСТ:**
Текст готового поста (800–1500 символов)

**ВАРИАНТЫ CTA:**
Нейтральный: [вариант]
С юмором: [вариант]
Прямой: [вариант]

**ВИЗУАЛ:**
Prompt (EN): [промпт для генерации изображения]
Описание (RU): [2–3 слова]
Запасной вариант: [упрощённый промпт]

---
## Пост 2
...
```

При отсутствии материалов:
```markdown
# Мониторинг ГГГГ-ММ-ДД ЧЧ:ММ
Проверено источников: N
Найдено релевантных материалов: 0

[ГГГГ-ММ-ДД][ЧЧ:ММ] — релевантных материалов не найдено. Проверено источников: N
```

**КРИТИЧНО:** Формат `**ГОТОВЫЙ ПОСТ:**` — жёсткий якорь для регулярки в `poster.py:extract_posts()`. Изменение формата требует синхронного обновления паттерна в этой функции.

---

## GitHub Actions Workflow (monitor_and_post.yml)

Workflow состоит из **4 последовательных шагов** в одном job:

| Шаг | Действие |
|-----|----------|
| `Checkout repository` | `actions/checkout@v4`, `fetch-depth: 0` |
| `Setup Python` | `actions/setup-python@v5`, Python 3.11, pip cache |
| `Install dependencies` | `pip install -r requirements.txt` |
| `Run monitor` | `python scripts/monitor.py` (переменная `GEMINI_API_KEY`) |
| `Commit generated posts` | `git add posts/` + коммит `Auto: posts ...` если есть изменения |
| `Post to Telegram channel` | `python scripts/poster.py` (переменные `BOT_TOKEN`, `CHANNEL_ID`) |
| `Commit post results` | `git pull --rebase --autostash` + `git add posts/` + коммит `Auto: mark posted ...` |

**Необходимые права:** `permissions: contents: write` (для push из Actions).

---

## Тон и стиль канала

- Лёгкий, немного хулиганский, с аккуратным юмором — без хамства и токсичности
- Подшучиваем над **ситуациями** и бюрократией — **не над людьми и организациями**
- Язык живой, без канцелярита; никаких юридических гарантий
- Формулировки: «как правило», «по практике», «есть риск что», «по сложившейся практике»
- Если материал о конкретном регионе — упоминаем регион
- Если упомянут нормативный акт — обязательно называем его
- В конце подходящих постов — ненавязчивое упоминание, что автор канала занимается оформлением

---

## Настройка автопостинга

### Шаг 1: Создать Telegram-бота и получить BOT_TOKEN

1. Открой [@BotFather](https://t.me/BotFather) в Telegram
2. Отправь команду `/newbot`
3. Придумай имя бота (например: `Land Water News Bot`) и username (например: `landwaterbot`)
4. BotFather пришлёт `BOT_TOKEN` в формате `1234567890:ABCdef...`

### Шаг 2: Добавить бота администратором канала

1. Открой настройки своего Telegram-канала
2. Перейди в **Администраторы → Добавить администратора**
3. Найди своего бота по username
4. Включи право **«Публикация сообщений»** (остальные можно отключить)

### Шаг 3: Узнать CHANNEL_ID

- **Публичный канал:** CHANNEL_ID = `@username_канала` (например, `@my_land_channel`)
- **Приватный канал:** перешли любое сообщение из канала боту [@userinfobot](https://t.me/userinfobot) — он покажет числовой ID в формате `-1001234567890`

### Шаг 4: Получить Google Gemini API Key

1. Перейди на [aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)
2. Нажми **Create API Key**
3. Скопируй ключ (начинается с `AIza...`)

### Шаг 5: Добавить секреты в GitHub

1. Открой репозиторий на GitHub
2. Перейди в **Settings → Secrets and variables → Actions**
3. Нажми **New repository secret** и добавь по одному:
   - `GEMINI_API_KEY` — ключ от Google Gemini
   - `BOT_TOKEN` — токен Telegram-бота
   - `CHANNEL_ID` — ID или @username канала

### Шаг 6: Протестировать

1. Перейди в **Actions → Monitor and Post**
2. Нажми **Run workflow → Run workflow**
3. Наблюдай за логами — должны пройти шаги:
   - `Run monitor` — появится файл `posts/ГГГГ-ММ-ДД_ЧЧ.md`
   - `Post to Telegram` — посты появятся в канале
   - `Commit post results` — в репозитории файл переименуется в `_posted.md`

### Шаг 7: Локальное тестирование

```bash
# Клонировать репозиторий
git clone https://github.com/Kedros64/telegram-channel-land-water.git
cd telegram-channel-land-water

# Установить зависимости
pip install -r requirements.txt

# Создать .env из примера и заполнить токены
cp .env.example .env
nano .env

# Запустить мониторинг
python scripts/monitor.py

# Опубликовать посты
python scripts/poster.py
```

---

## Расширение источников

### Добавить сайт

1. Открой `sources.xlsx` (Excel / LibreOffice Calc)
2. Добавь строку в лист **Sources**: `type = site`, `active = да`
3. В `url` — ссылка на **страницу-список** новостей, не главную. Например: `https://rosreestr.gov.ru/press/archive/`
4. Если у сайта есть RSS — можно указать прямую ссылку на RSS-фид; `monitor.py` определит его автоматически
5. `git add sources.xlsx && git commit -m "Add source: ..."`

### Добавить Telegram-канал

1. Добавь строку в `sources.xlsx`: `type = telegram`, `active = да`
2. В `url` — `https://t.me/channelname` или `@channelname`
3. Канал должен быть **публичным** — приватные возвращают 404 и пропускаются с предупреждением
4. `git add sources.xlsx && git commit -m "Add Telegram source: ..."`

---

## Мониторинг Telegram-каналов

Реализовано через `https://t.me/s/{channel}` — **без MTProto и API-ключей**, работает для всех публичных каналов.

Если потребуется мониторинг **приватных** каналов — нужен Telethon (MTProto API):
1. Зарегистрируйся на [my.telegram.org](https://my.telegram.org), получи `API_ID` и `API_HASH`
2. Добавь `telethon>=1.36.0` в `requirements.txt`
3. Добавь `TELEGRAM_API_ID`, `TELEGRAM_API_HASH`, `TELEGRAM_SESSION` в `.env` и GitHub Secrets
4. Перепиши `fetch_telegram_channel()` в `monitor.py` с использованием `TelegramClient`

---

## Разработка

### Ветки

- `feature/<описание>` — новые функции
- `fix/<описание>` — исправление ошибок
- `claude/<task-id>` — управляемые AI-ассистентом

### Коммиты

Тема — до 72 символов, императивный стиль:

```
Add Gemini-based relevance filtering
Fix RSS parsing for Atom feeds
Update sources.xlsx: add 5 new Telegram channels
Replace feedparser with stdlib xml.etree.ElementTree
```

Автоматические коммиты workflow имеют формат:
- `Auto: posts ГГГГ-ММ-ДД ЧЧ:ММ UTC`
- `Auto: mark posted ГГГГ-ММ-ДД ЧЧ:ММ UTC`

---

## Переменные окружения

| Переменная       | Где используется | Описание                                    |
|------------------|------------------|---------------------------------------------|
| `GEMINI_API_KEY` | `monitor.py`     | Ключ Google Gemini API для фильтрации и генерации постов |
| `BOT_TOKEN`      | `poster.py`      | Токен Telegram-бота от @BotFather           |
| `CHANNEL_ID`     | `poster.py`      | ID или @username Telegram-канала            |

Все переменные хранятся:
- **Локально:** в файле `.env` (в `.gitignore`, не коммитится)
- **В CI/CD:** в GitHub Secrets репозитория

---

## Ключевые архитектурные решения

| Решение | Обоснование |
|---|---|
| Gemini-based relevance filtering (батч) | Один API-вызов на все заголовки — точнее ключевых слов, не требует поддержки словарей |
| Прямые REST-вызовы к Gemini (без SDK) | Убирает зависимость `google-generativeai`; проще отлаживать и контролировать |
| stdlib XML вместо feedparser | Одна зависимость меньше; feedparser избыточен для RSS 2.0 и Atom |
| RSS-first + HTML fallback | RSS структурированнее и стабильнее; HTML — универсальный fallback |
| t.me/s/{channel} для Telegram | Без MTProto и API-ключей; работает для всех публичных каналов |
| HTML parse_mode в Telegram | Надёжнее MarkdownV2 — меньше проблем с экранированием спецсимволов |
| SSL `verify=False` + urllib3 warnings off | Госсайты часто используют устаревшие/самоподписанные сертификаты |
| Переименование в `_posted.md` | Состояние публикации видно в git-истории; не зависит от внешних файлов |
| Пауза 30 сек между постами | Предотвращает флуд-фильтр Telegram |
| `last_check.txt` в git | Состояние проверки сохраняется между запусками workflow |

---

## Что нельзя делать

- Не коммить `.env` — только `.env.example`
- Не хардкодить токены и ключи в коде
- Не пушить напрямую в `main`
- Не удалять `posts/last_check.txt` — сломает отслеживание новых публикаций
- Не менять формат `**ГОТОВЫЙ ПОСТ:**` без обновления регулярки в `poster.py:extract_posts()`
- Не коммить `sources.xlsx` с личными данными или credentials
- Не добавлять `feedparser` или `google-generativeai` обратно — они намеренно заменены stdlib и REST-вызовами

---

*Последнее обновление: 2026-03-04 — Актуализация: Gemini-фильтрация (батч), RSS-first стратегия, t.me/s парсинг Telegram, 51 источник, stdlib XML, прямые REST-вызовы к Gemini API*
