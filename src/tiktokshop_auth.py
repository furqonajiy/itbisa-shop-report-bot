"""
tiktokshop_auth.py
------------------
TikTok Shop token lifecycle, mirrored from itbisa-shop-stock-bot
(which mirrored from itbisa-tiktokshop-order-bot).

Token file format (data/tiktokshop_tokens.json), four fields:
  {
    "access_token": "...",
    "refresh_token": "...",
    "access_token_expires_at":  "2026-05-09T00:00:00+00:00",
    "refresh_token_expires_at": "2125-01-01T00:00:00+00:00"
  }

Both expiry fields are stored as ISO. TikTok returns Unix timestamps;
we convert at save-time so the on-disk format is human-readable.

Refresh failure raises RuntimeError (not a custom class — matches the
contract of the existing TikTok order bot, where main.py catches all
exceptions and forwards to Telegram).
"""

import json
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

from src import config


def get_valid_access_token() -> str:
    """Returns a non-expired access_token, refreshing if needed."""
    tokens = _load_tokens()

    if not _is_access_token_expiring_soon(tokens):
        return tokens["access_token"]

    print("  [tiktokshop_auth] Access token near expiry, refreshing...")
    new_tokens = _refresh_access_token(tokens["refresh_token"])
    _save_tokens(new_tokens)
    return new_tokens["access_token"]


# ============================================================
# Token file I/O
# ============================================================

def _load_tokens() -> dict:
    path = Path(config.TIKTOKSHOP_TOKEN_FILE)
    if not path.exists():
        raise RuntimeError(
            f"TikTok Shop token file not found at {path}. "
            f"Run scripts/bootstrap_tiktokshop_tokens.py first."
        )
    with open(path, "r") as f:
        return json.load(f)


def _save_tokens(tokens: dict) -> None:
    path = Path(config.TIKTOKSHOP_TOKEN_FILE)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    with open(tmp, "w") as f:
        json.dump(tokens, f, indent=2)
    tmp.replace(path)


def _is_access_token_expiring_soon(tokens: dict) -> bool:
    expires_at = datetime.fromisoformat(tokens["access_token_expires_at"])
    if expires_at.tzinfo is None:
        expires_at = expires_at.replace(tzinfo=timezone.utc)
    buffer = timedelta(minutes=config.TIKTOKSHOP_TOKEN_REFRESH_BUFFER_MINUTES)
    return datetime.now(timezone.utc) + buffer >= expires_at


# ============================================================
# Refresh flow — auth host, plain GET, NOT signed
# ============================================================

def _refresh_access_token(refresh_token: str) -> dict:
    """
    GET https://auth.tiktok-shops.com/api/v2/token/refresh
        ?app_key=...&app_secret=...&refresh_token=...&grant_type=refresh_token

    Plain GET with query params. NOT signed. NOT in scope of the
    Open API signing rules.
    """
    path = "/api/v2/token/refresh"
    url = f"{config.TIKTOKSHOP_AUTH_BASE_URL}{path}"
    params = {
        "app_key":       config.TIKTOKSHOP_APP_KEY,
        "app_secret":    config.TIKTOKSHOP_APP_SECRET,
        "refresh_token": refresh_token,
        "grant_type":    "refresh_token",
    }

    response = requests.get(url, params=params, timeout=30)
    response.raise_for_status()
    data = response.json()

    if data.get("code") not in (0, "0"):
        raise RuntimeError(
            f"TikTok Shop refresh failed: code={data.get('code')} "
            f"message={data.get('message')}"
        )

    payload = data.get("data") or {}
    access_token  = payload.get("access_token")
    refresh_new   = payload.get("refresh_token")
    access_exp    = payload.get("access_token_expire_in")
    refresh_exp   = payload.get("refresh_token_expire_in")

    if not access_token or not refresh_new:
        raise RuntimeError(f"TikTok Shop refresh missing tokens: {data}")

    # TikTok returns Unix timestamps for these fields, NOT durations.
    # The order bot's bootstrap script discovered this empirically.
    return {
        "access_token":              access_token,
        "refresh_token":             refresh_new,
        "access_token_expires_at":   _ts_to_iso(int(access_exp)),
        "refresh_token_expires_at":  _ts_to_iso(int(refresh_exp)),
    }


def _ts_to_iso(unix_ts: int) -> str:
    return datetime.fromtimestamp(unix_ts, tz=timezone.utc).isoformat()
