#!/usr/bin/env python3
"""
trigger_bot.py — Telegram-бот для ручного запуска мониторинга.

Отправь /run боту — он немедленно запустит мониторинг через GitHub Actions.
Работает в режиме long polling; запусти один раз и оставь работать в фоне.

Запуск:
    python scripts/trigger_bot.py

Переменные окружения (добавь в .env):
    BOT_TOKEN      — токен Telegram-бота (уже настроен)
    OWNER_CHAT_ID  — числовой Telegram ID владельца (узнай через @userinfobot)
    GH_PAT         — GitHub Personal Access Token (нужны права: workflow)
    GITHUB_REPO    — репозиторий (по умолчанию: Kedros64/telegram-channel-land-water)
    GITHUB_BRANCH  — ветка для запуска (по умолчанию: master)
"""

import logging
import os
import sys
import time
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "")
OWNER_CHAT_ID = str(os.getenv("OWNER_CHAT_ID", ""))
GITHUB_TOKEN = os.getenv("GH_PAT", "")
GITHUB_REPO = os.getenv("GITHUB_REPO", "Kedros64/telegram-channel-land-water")
GITHUB_BRANCH = os.getenv("GITHUB_BRANCH", "master")
WORKFLOW_FILE = "monitor_and_post.yml"

REPO_ROOT = Path(__file__).parent.parent
LAST_CHECK_FILE = REPO_ROOT / "posts" / "last_check.txt"

TG_API = f"https://api.telegram.org/bot{BOT_TOKEN}"
GH_DISPATCH_URL = (
    f"https://api.github.com/repos/{GITHUB_REPO}"
    f"/actions/workflows/{WORKFLOW_FILE}/dispatches"
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Telegram API helpers
# ---------------------------------------------------------------------------


def send(chat_id: int | str, text: str) -> None:
    """Отправляет сообщение в Telegram с HTML-разметкой."""
    try:
        requests.post(
            f"{TG_API}/sendMessage",
            json={"chat_id": chat_id, "text": text, "parse_mode": "HTML"},
            timeout=15,
        )
    except Exception as exc:
        log.warning("send_message failed: %s", exc)


def get_updates(offset: int = 0) -> list[dict]:
    """Long polling: возвращает список новых обновлений от Telegram."""
    try:
        resp = requests.get(
            f"{TG_API}/getUpdates",
            params={"offset": offset, "timeout": 30, "allowed_updates": ["message"]},
            timeout=40,
        )
        if resp.ok:
            return resp.json().get("result", [])
    except Exception as exc:
        log.warning("getUpdates failed: %s", exc)
    return []


# ---------------------------------------------------------------------------
# GitHub Actions trigger
# ---------------------------------------------------------------------------


def trigger_workflow(chat_id: int | str) -> None:
    """Запускает GitHub Actions workflow через repository dispatch."""
    if not GITHUB_TOKEN:
        send(
            chat_id,
            "❌ <b>GH_PAT не настроен.</b>\n"
            "Добавьте Personal Access Token в .env (нужны права: workflow).",
        )
        return

    try:
        resp = requests.post(
            GH_DISPATCH_URL,
            headers={
                "Authorization": f"token {GITHUB_TOKEN}",
                "Accept": "application/vnd.github.v3+json",
            },
            json={"ref": GITHUB_BRANCH},
            timeout=15,
        )
    except Exception as exc:
        send(chat_id, f"❌ Ошибка соединения с GitHub API: {exc}")
        return

    if resp.status_code == 204:
        runs_url = f"https://github.com/{GITHUB_REPO}/actions"
        send(
            chat_id,
            "✅ <b>Мониторинг запущен!</b>\n\n"
            "Посты появятся в канале через 2–3 минуты.\n"
            f'<a href="{runs_url}">Следить за прогрессом в Actions →</a>',
        )
        log.info("Workflow dispatched успешно (repo=%s, branch=%s)", GITHUB_REPO, GITHUB_BRANCH)
    elif resp.status_code == 401:
        send(chat_id, "❌ GH_PAT недействителен или недостаточно прав (нужен scope: workflow).")
        log.error("GitHub 401: %s", resp.text[:200])
    elif resp.status_code == 404:
        send(chat_id, f"❌ Репозиторий или workflow не найдены: {GITHUB_REPO}/{WORKFLOW_FILE}")
        log.error("GitHub 404: repo=%s workflow=%s", GITHUB_REPO, WORKFLOW_FILE)
    elif resp.status_code == 422:
        send(chat_id, f"❌ Ветка не найдена: {GITHUB_BRANCH}")
        log.error("GitHub 422: branch=%s", GITHUB_BRANCH)
    else:
        send(chat_id, f"❌ GitHub API вернул {resp.status_code}:\n<code>{resp.text[:300]}</code>")
        log.error("GitHub dispatch: %s %s", resp.status_code, resp.text[:300])


# ---------------------------------------------------------------------------
# Command handlers
# ---------------------------------------------------------------------------


def handle_command(chat_id: int | str, text: str) -> None:
    # Убираем @botname из команды (например: /run@mybot → /run)
    cmd = text.split()[0].lower().split("@")[0]

    if cmd == "/run":
        send(chat_id, "⏳ Инициирую запуск мониторинга...")
        trigger_workflow(chat_id)

    elif cmd == "/status":
        if LAST_CHECK_FILE.exists():
            ts = LAST_CHECK_FILE.read_text(encoding="utf-8").strip()
            send(chat_id, f"🕐 Последняя проверка источников:\n<code>{ts} UTC</code>")
        else:
            send(chat_id, "Данных о последней проверке нет.")

    elif cmd in ("/start", "/help"):
        send(
            chat_id,
            "<b>Доступные команды:</b>\n\n"
            "/run — немедленный запуск мониторинга и публикации\n"
            "/status — время последней проверки источников\n"
            "/help — эта справка",
        )


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------


def main() -> None:
    if not BOT_TOKEN:
        log.error("BOT_TOKEN не установлен")
        sys.exit(1)

    if not OWNER_CHAT_ID:
        log.error(
            "OWNER_CHAT_ID не установлен.\n"
            "Узнайте свой Telegram ID через @userinfobot и добавьте в .env:\n"
            "  OWNER_CHAT_ID=123456789"
        )
        sys.exit(1)

    log.info("Trigger bot запущен (owner_id=%s, repo=%s/%s)", OWNER_CHAT_ID, GITHUB_REPO, GITHUB_BRANCH)
    log.info("Ожидаю команды /run и /status от владельца...")

    offset = 0
    while True:
        updates = get_updates(offset=offset)

        for update in updates:
            offset = update["update_id"] + 1

            msg = update.get("message", {})
            chat_id = msg.get("chat", {}).get("id")
            text = (msg.get("text") or "").strip()

            if not chat_id or not text.startswith("/"):
                continue

            if str(chat_id) != OWNER_CHAT_ID:
                log.debug("Сообщение от незнакомого chat_id=%s — игнорирую", chat_id)
                continue

            log.info("Команда от владельца: %r", text)
            handle_command(chat_id, text)

        if not updates:
            time.sleep(1)


if __name__ == "__main__":
    main()
