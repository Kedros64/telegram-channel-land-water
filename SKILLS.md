# SKILLS.md — плейбук проекта (переиспользуемые приёмы)

## Пути, доступы, окружение
- **Локальная рабочая копия / источник истины:** `C:\Users\17D3~1\AppData\Local\Temp\claude\landwater-bot\` (это папка сборки: `bot.py`, `scripts/`, `sources.xlsx`, `assets/`, `deploy_secrets.json`, ключи, `_*`-скретчи). ⚠️ Это временная папка сессии — durable-исходник кода лежит в репо `bot/app.tar.gz`.
- **Секреты:** локально `deploy_secrets.json`; на сервере `/opt/app/.env`.
- **SSH к серверу:** `ssh -i bot_key -o StrictHostKeyChecking=no root@$(cat _tw_ipv4.txt)` (текущий IP в `_tw_ipv4.txt`, сейчас `201.24.60.189`). Ключ в Timeweb — id `714895`.
- **Timeweb API:** токен в `tw_token.txt` (заголовок `Authorization: Bearer …`, аккаунт `hk054512`), база `https://api.timeweb.cloud/api/v1`.
- **GitHub:** repo `Kedros64/telegram-channel-land-water`, ветка `claude/claude-md-mm6x6kpv5ry5lx3t-0cgxW`; коммит файлов — Contents API (`PUT /contents/{path}` с base64 + `sha` существующего).

## Деплой обновления кода (venv/зависимости уже стоят)
```bash
cd landwater-bot
python -m py_compile bot.py scripts/*.py            # ВСЕГДА перед сборкой
tar czf app.tar.gz bot.py requirements.txt sources.xlsx \
  scripts/__init__.py scripts/monitor.py scripts/poster.py scripts/genapi.py \
  scripts/overlay.py scripts/channels.py assets/Montserrat-Black.ttf
# → PUT app.tar.gz в repo bot/app.tar.gz (Contents API, с sha)
# на сервере тянем тот же пакет и рестартим:
ssh -i bot_key root@$(cat _tw_ipv4.txt) 'cd /opt/app && set -a && . ./.env && set +a && \
  curl -fsSL -H "Authorization: Bearer $GITHUB_TOKEN" -H "Accept: application/vnd.github.raw" \
   "https://api.github.com/repos/$GITHUB_REPO/contents/bot/app.tar.gz?ref=$GITHUB_BRANCH" -o app.tar.gz && \
  tar xzf app.tar.gz -C /opt/app && systemctl reset-failed landwater-bot && systemctl restart landwater-bot'
```

## Диагностика бота
- **Поллит ли Telegram:** `curl "https://api.telegram.org/bot<token>/getWebhookInfo"` → если `pending_update_count` не 0 и растёт, а `url` пуст — процесс НЕ забирает апдейты (упал/не поллит).
- **Логи/статус на сервере:** `tail -n 80 /var/log/landwater-bot.log`; `systemctl status landwater-bot`; `systemctl show landwater-bot -p ActiveState,SubState,MainPID,NRestarts`.
- **Health:** `curl http://<ip>:3000/health`.
- **Состояние:** `/opt/app/data/state.json` (текущий pending-пост), `/opt/app/data/preferences.json` (выученные правила), `/opt/app/data/articles_pool.json` (пул новостей).
- **Проверить парсер на реальном тексте:** `venv/bin/python -c "import sys;sys.path.insert(0,'/opt/app');import scripts.monitor as m,json;print(len(m.extract_image_prompt(json.load(open('/opt/app/data/state.json',encoding='utf-8'))['raw_text'])))"`.

## gen-api.ru (картинки / речь)
- **Картинка:** `POST /api/v1/networks/nano-banana-2` `{prompt, resolution:"0.5K", aspect_ratio:"1:1", is_sync:false}` → `request_id` → поллить `GET /api/v1/request/get/{id}` (status `success`) → `result[0]` = URL. Заголовок `Authorization: Bearer <GENAPI_KEY>`.
- **Whisper (STT):** `POST /api/v1/networks/whisper {audio_url}`. Telegram отдаёт голос как `.oga` — gen-api его НЕ принимает; кладём в GitHub как `.ogg` и передаём raw-URL (`github_put_bytes`).
- Публичного списка моделей нет (каталог на сайте gen-api.ru). **Проверка ключа = реальная мелкая генерация** (эндпоинт `/api/v1/user` даёт ложный «неверный ключ» на рабочем ключе).

## Telegram
- **Публичность канала** (нужно для синхробота Дзена): `getChat?chat_id=-100…` → поле `username` (у нас `pro_zemlyu_i_vodu`).
- Картинка-в-теле: если plain-текст ≤ `TELEGRAM_CAPTION_MAX` (1024) → `sendPhoto` с caption; иначе `sendMessage` + превью по raw-URL картинки из репо.

## VK-публикация (`scripts/channels.py`)
- `wall.post`: `owner_id = -abs(group_id)`, `from_group=1`, `v=5.199`, `message` = plain (без markdown — есть `channels._plain`), `random_id` уникальный.
- Фото (3 шага): `photos.getWallUploadServer` (лучше **user-токен админа** — у community-токена бывает **error 27**) → multipart POST файла (поле `photo`) на `upload_url` → `photos.saveWallPhoto` → attachment `photo{owner}_{id}`.
- Включение канала: env `EXTRA_CHANNELS=vk`, `VK_TOKEN`, `VK_GROUP_ID`, при необходимости `VK_UPLOAD_TOKEN`.

## Timeweb (провижн нового сервера)
- Создать: `POST /servers` `{name, os_id:99 (Ubuntu 24.04), preset_id, bandwidth:100, cloud_init, ssh_keys_ids:[714895]}`. Дешёвые РФ-тарифы: `5039` (ru-2, 1CPU/2GB, 264 ₽/мес).
- **IPv4:** назначать ТОЛЬКО после статуса `on` → `POST /servers/{id}/ips {type:ipv4}` → затем **reboot** (иначе гость не поднимет сеть).
- **Чистить мусор:** `GET /api/v1/floating-ips` → `DELETE /api/v1/floating-ips/{id}` для всех с `resource_id=null` (они платные и раздувают анти-фрод-порог).
- **cloud-init/bootstrap** (исправленный): `apt-get update` → `apt-get install -y python3.12-venv python3-pip` → `python3 -m venv` → `venv/bin/python -m pip install -r requirements.txt` → systemd `enable --now`.
- Действия сервера: `POST /servers/{id}/action {"action":"reboot"}` (эндпоинт в единственном числе — `/action`).

## Ведение контекста
- После значимой задачи — обнови **PROGRESS.md** (что сделано / в процессе / решения / грабли).
- Новый рабочий приём/команду/нюанс — добавляй сюда, в **SKILLS.md**.
