"""
telegram_sender.py
------------------
Telegram output for the report bot. Bahasa Indonesia by convention.

PR 1 surface:
  send_message(text)              — posts a plain text message
  send_document(file_path, caption)  — posts a file (xlsx) with optional caption

Both helpers respect TELEGRAM_CHAT_ID from config and never accept a
chat_id argument. Sending to other chats would defeat the security
gate enforced by the Worker.
"""

from __future__ import annotations

from pathlib import Path

import requests

from src import config


_API_BASE = f"https://api.telegram.org/bot{config.TELEGRAM_BOT_TOKEN}"


def send_message(text: str) -> bool:
    """
    Posts a plain text message to TELEGRAM_CHAT_ID.
    Returns True on HTTP 200, False otherwise. Never raises.
    """
    try:
        response = requests.post(
            f"{_API_BASE}/sendMessage",
            json={
                "chat_id": config.TELEGRAM_CHAT_ID,
                "text":    text,
                # Plain text — no parse_mode, so '*' '_' etc. render literally.
            },
            timeout=30,
        )
        return response.status_code == 200
    except Exception as exc:
        print(f"  [telegram] sendMessage failed: {exc}")
        return False


def send_document(file_path: Path | str, caption: str = "") -> bool:
    """
    Posts a file as a Telegram document to TELEGRAM_CHAT_ID.

    Used by PR 2/3 to deliver generated xlsx files to chat. The file
    appears as a downloadable attachment, NOT inline.

    Returns True on HTTP 200, False otherwise. Never raises.
    """
    path = Path(file_path)
    if not path.exists():
        print(f"  [telegram] sendDocument: file not found: {path}")
        return False

    try:
        with open(path, "rb") as f:
            files = {"document": (path.name, f)}
            data = {"chat_id": config.TELEGRAM_CHAT_ID}
            if caption:
                data["caption"] = caption
            response = requests.post(
                f"{_API_BASE}/sendDocument",
                files=files,
                data=data,
                timeout=120,  # xlsx upload can take a few seconds
            )
        return response.status_code == 200
    except Exception as exc:
        print(f"  [telegram] sendDocument failed: {exc}")
        return False
