"""
config.py
---------
All environment variables and constants in one place. Both Shopee and
TikTok Shop credentials live here because this bot talks to both
platforms (and TikTok serves both Tiktok and Tokopedia channels).

Loaded once at import time — missing required variables fail loudly
at startup rather than mid-run. Mirrors the layout of config.py in
itbisa-shop-stock-bot so a developer who knows one knows this.
"""

import os
from datetime import date, datetime
from pathlib import Path

from dotenv import load_dotenv


# Load .env from project root for local dev. In CI the values come
# from GitHub Actions secrets and the .env file is absent.
PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")


def _required(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise RuntimeError(
            f"Missing required environment variable: {name}. "
            f"Set it in .env (local) or GitHub Secrets (CI)."
        )
    return value


def _optional(name: str, default: str = "") -> str:
    return os.environ.get(name, default)


# ============================================================
# Shopee
# ============================================================

SHOPEE_PARTNER_ID  = _required("SHOPEE_PARTNER_ID")
SHOPEE_PARTNER_KEY = _required("SHOPEE_PARTNER_KEY")
SHOPEE_SHOP_ID     = _required("SHOPEE_SHOP_ID")

# Live Shopee Open API host. Same as the order bot. Switching to
# sandbox would require running bootstrap_shopee_tokens.py against
# the sandbox URL and is not supported by this repo's tooling.
SHOPEE_API_BASE_URL = "https://partner.shopeemobile.com"

SHOPEE_TOKEN_FILE = PROJECT_ROOT / "data" / "shopee_tokens.json"

# Refresh access_token this many minutes before its declared expiry.
# 10 min matches the order bot — keeps refresh timing in lockstep.
SHOPEE_TOKEN_REFRESH_BUFFER_MINUTES = 10


# ============================================================
# TikTok Shop
# ============================================================

TIKTOKSHOP_APP_KEY    = _required("TIKTOKSHOP_APP_KEY")
TIKTOKSHOP_APP_SECRET = _required("TIKTOKSHOP_APP_SECRET")
TIKTOKSHOP_SHOP_ID    = _required("TIKTOKSHOP_SHOP_ID")

# TikTok Shop uses two distinct hosts: auth and Open API.
TIKTOKSHOP_AUTH_BASE_URL     = "https://auth.tiktok-shops.com"
TIKTOKSHOP_OPEN_API_BASE_URL = "https://open-api.tiktokglobalshop.com"

TIKTOKSHOP_TOKEN_FILE = PROJECT_ROOT / "data" / "tiktokshop_tokens.json"

TIKTOKSHOP_TOKEN_REFRESH_BUFFER_MINUTES = 10


# ============================================================
# Telegram
# ============================================================

TELEGRAM_BOT_TOKEN = _required("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID   = _required("TELEGRAM_CHAT_ID")


# ============================================================
# Report bot behaviour
# ============================================================

# Output xlsx files. Lives under bot-state branch in production;
# local-only when running scripts/* directly.
OUTPUT_DIR    = PROJECT_ROOT / "output"
OUTPUT_SHOPEE_DIR     = OUTPUT_DIR / "shopee"
OUTPUT_TIKTOKSHOP_DIR = OUTPUT_DIR / "tiktokshop"
OUTPUT_TOKOPEDIA_DIR  = OUTPUT_DIR / "tokopedia"

# Period state — tracks which months have been generated. PR 2.
STATE_DIR = PROJECT_ROOT / "state"

# Politeness delay between paginated API calls (seconds). Same value
# as the order bots, never hit a rate limit in production.
DELAY_BETWEEN_CALLS_SECONDS = 0.3

# Pagination defaults.
SHOPEE_WALLET_PAGE_SIZE  = 100   # max per Shopee docs
TIKTOKSHOP_FINANCE_PAGE_SIZE = 50  # conservative; max is 100


# ============================================================
# Period helpers
# ============================================================

def resolve_period_label(today: date, target_year: int, target_month: int) -> str:
    """
    Returns the filename suffix for a generated report.

    Rule (per operator decision):
      • Run for current month → 'YYYY-MM-DD' (today's date).
      • Run for past month   → 'YYYY-MM'.

    Examples:
      today=2026-04-15, target=2026-04 → '2026-04-15'  (in-progress)
      today=2026-04-15, target=2026-03 → '2026-03'     (finalized)
      today=2026-05-01, target=2026-04 → '2026-04'     (just-closed prev month)
    """
    if today.year == target_year and today.month == target_month:
        return today.strftime("%Y-%m-%d")
    return f"{target_year:04d}-{target_month:02d}"


def parse_period_arg(period_str: str) -> tuple[int, int]:
    """
    Parses 'YYYY-MM' to (year, month). Raises ValueError on bad input.
    Used by scripts/dump_reason_codes.py and (later) by main.py.
    """
    parts = period_str.split("-")
    if len(parts) != 2:
        raise ValueError(f"Bad period format '{period_str}', expected YYYY-MM")
    year, month = int(parts[0]), int(parts[1])
    if not (2020 <= year <= 2100 and 1 <= month <= 12):
        raise ValueError(f"Bad period values year={year} month={month}")
    return year, month


def month_bounds_utc(year: int, month: int) -> tuple[datetime, datetime]:
    """
    Returns (start, end_exclusive) datetimes in UTC for the given
    calendar month.

    Both Shopee and TikTok APIs accept Unix timestamps; callers can
    int(start.timestamp()) and int(end.timestamp()) freely.

    NOTE: We use UTC bounds, not WIB, because both APIs document
    their time params as Unix timestamps (timezone-agnostic). The
    one-hour shift between UTC and WIB is irrelevant for monthly
    aggregation — a transaction at 23:30 UTC on Jan 31 is also at
    06:30 WIB on Feb 1, and the legacy generator (which used WIB
    via seller-center exports) would have put it in February too.
    For consistency with what you used to see in the legacy report,
    if you want WIB bounds instead, this is the one knob to turn.
    """
    from datetime import timezone

    start = datetime(year, month, 1, tzinfo=timezone.utc)
    if month == 12:
        end = datetime(year + 1, 1, 1, tzinfo=timezone.utc)
    else:
        end = datetime(year, month + 1, 1, tzinfo=timezone.utc)
    return start, end
