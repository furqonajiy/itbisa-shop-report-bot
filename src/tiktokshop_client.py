"""
tiktokshop_client.py
--------------------
TikTok Shop Open API integration for the report bot.

PR 1 surface:
  describe()
    Identifier for log/Telegram headers.

  iter_statements(start_ts_unix, end_ts_unix) -> Iterator[dict]
    Paginates GET /finance/202309/statements for the period. Each
    yielded dict is one settlement statement.

  iter_statement_transactions(statement_id) -> Iterator[dict]
    Paginates GET /finance/202309/statements/{id}/statement_transactions.
    Each yielded dict is one settlement-line transaction. This is the
    TikTok equivalent of Shopee's wallet_transaction rows and is the
    source of truth for BisaRemit / KERUGIAN TAMBAHAN / KEUNTUNGAN
    classification.

PR 2 will add:
  iter_orders(start_ts_unix, end_ts_unix)  — order/202309/orders/search
  get_order_settlement(order_id)           — finance/.../order_settlements

Internal:
  _call_signed(method, path, ...)
  _check_ok(response, context)
  _get_shop_cipher() — fetched once per process, cached in module-level
                       global. Bootstrap endpoint (/authorization/202309/shops)
                       is the one Open API call that does NOT require cipher.
"""

from __future__ import annotations

import hashlib
import hmac
import json as _json
import time
from typing import Iterator

import requests

from src import config, tiktokshop_auth


# ============================================================
# API versions per endpoint
# ============================================================
# We deliberately use the latest stable TikTok Open API versions
# this app's scopes have access to. If you upgrade to a newer
# version (e.g. 202407 finance), update both the path and the
# `version` query param the signing layer adds.

_FINANCE_API_VERSION = "202309"
_AUTH_API_VERSION    = "202309"

# Finance statements API tops out around 100 per page; we use the
# config default for politeness.
_STATEMENT_TXN_PAGE_SIZE = 50


# Cached shop_cipher for the duration of one process run. Fetched
# lazily on first signed call. Same pattern as the order bot and
# the stock bot.
_cached_shop_cipher: str | None = None


# ============================================================
# Public surface
# ============================================================

def describe() -> str:
    return "TikTok Shop"


def iter_statements(
    start_ts_unix: int,
    end_ts_unix: int,
) -> Iterator[dict]:
    """
    Yields every settlement statement created in [start, end_exclusive).

    Endpoint: GET /finance/{version}/statements
    Query: statement_time_ge, statement_time_lt, page_size, page_token
    """
    path = f"/finance/{_FINANCE_API_VERSION}/statements"
    page_token = ""

    while True:
        extra_query = {
            "statement_time_ge": str(int(start_ts_unix)),
            "statement_time_lt": str(int(end_ts_unix)),
            "page_size":         str(config.TIKTOKSHOP_FINANCE_PAGE_SIZE),
            # 202309 statements REQUIRES both of these (HTTP 400 / code
            # 36009004 'SortField is a required field' otherwise). Order
            # is irrelevant for our aggregation, so DESC by statement_time.
            "sort_field":        "statement_time",
            "sort_order":        "DESC",
        }
        if page_token:
            extra_query["page_token"] = page_token

        response = _call_signed("GET", path, extra_query=extra_query)
        _check_ok(response, context=f"statements page_token={page_token!r}")

        payload = response.json().get("data") or {}
        statements = payload.get("statements") or []
        for s in statements:
            yield s

        next_token = payload.get("next_page_token") or ""
        if not next_token:
            return
        page_token = next_token
        time.sleep(config.DELAY_BETWEEN_CALLS_SECONDS)


def iter_statement_transactions(statement_id: str) -> Iterator[dict]:
    """
    Yields every transaction inside a settlement statement.

    Endpoint: GET /finance/{version}/statements/{statement_id}/statement_transactions
    Query: page_size, page_token
    """
    path = (
        f"/finance/{_FINANCE_API_VERSION}/statements/"
        f"{statement_id}/statement_transactions"
    )
    page_token = ""

    while True:
        extra_query: dict[str, str] = {
            "page_size": str(_STATEMENT_TXN_PAGE_SIZE),
        }
        if page_token:
            extra_query["page_token"] = page_token

        response = _call_signed("GET", path, extra_query=extra_query)
        _check_ok(response, context=f"statement_transactions stmt={statement_id} pt={page_token!r}")

        payload = response.json().get("data") or {}
        txns = payload.get("statement_transactions") or []
        for t in txns:
            yield t

        next_token = payload.get("next_page_token") or ""
        if not next_token:
            return
        page_token = next_token
        time.sleep(config.DELAY_BETWEEN_CALLS_SECONDS)


# ============================================================
# Internal: shop cipher (fetched once per run)
# ============================================================

def _get_shop_cipher() -> str:
    """
    Returns the shop cipher, fetching it once and caching for the run.

    The cipher comes from /authorization/{version}/shops — the one
    endpoint that does NOT require shop_cipher (chicken-and-egg).
    """
    global _cached_shop_cipher
    if _cached_shop_cipher:
        return _cached_shop_cipher

    response = _call_signed(
        "GET",
        f"/authorization/{_AUTH_API_VERSION}/shops",
        include_cipher=False,
    )
    _check_ok(response, context="authorized shops")

    shops = response.json()["data"]["shops"]
    matching = next(
        (s for s in shops if str(s["id"]) == str(config.TIKTOKSHOP_SHOP_ID)),
        None,
    )
    if matching is None:
        raise RuntimeError(
            f"Shop {config.TIKTOKSHOP_SHOP_ID} not found in authorized shops. "
            f"Found: {[s['id'] for s in shops]}"
        )

    _cached_shop_cipher = matching["cipher"]
    print(f"  [tiktokshop] Fetched shop cipher for shop {config.TIKTOKSHOP_SHOP_ID}")
    return _cached_shop_cipher


# ============================================================
# Internal: signed HTTP transport
# ============================================================

def _call_signed(
    method: str,
    path: str,
    *,
    extra_query: dict[str, str] | None = None,
    body: dict | list | None = None,
    include_cipher: bool = True,
) -> requests.Response:
    """
    Signs and dispatches an Open API call.

    Signing per system prompt:
      1. Exclude 'sign' and 'access_token' from query params; drop empty values.
      2. Sort remaining params by key, concatenate as key+value (no separator).
      3. canonical = path + sorted_param_string + raw_body_string
      4. wrapped   = app_secret + canonical + app_secret
      5. sign      = HMAC-SHA256(app_secret, wrapped).hexdigest()

    Transport: x-tts-access-token header carries the access token
    (NOT in the signature input).
    """
    access_token = tiktokshop_auth.get_valid_access_token()
    timestamp = str(int(time.time()))

    query: dict[str, str] = {
        "app_key":   config.TIKTOKSHOP_APP_KEY,
        "shop_id":   str(config.TIKTOKSHOP_SHOP_ID),
        "timestamp": timestamp,
    }
    if extra_query:
        query.update(extra_query)

    if include_cipher:
        cipher = _get_shop_cipher()
        if cipher:
            query["shop_cipher"] = cipher

    # Body for signing: compact JSON if dict/list, else empty string.
    raw_body = ""
    if body is not None:
        raw_body = _json.dumps(body, separators=(",", ":"))

    sign = _build_signature(path, query, raw_body)
    query_with_sign = {**query, "sign": sign}

    url = f"{config.TIKTOKSHOP_OPEN_API_BASE_URL}{path}"
    headers = {
        "x-tts-access-token": access_token,
        "Content-Type":       "application/json",
    }

    if method.upper() == "GET":
        return requests.get(url, params=query_with_sign, headers=headers, timeout=60)
    if method.upper() == "POST":
        return requests.post(
            url,
            params=query_with_sign,
            data=raw_body if body is not None else None,
            headers=headers,
            timeout=60,
        )
    raise ValueError(f"Unsupported method: {method}")


def _build_signature(path: str, query: dict[str, str], raw_body: str) -> str:
    """Implements the 5-step signing rule documented above _call_signed."""
    # Step 1: exclude sign and access_token; drop empty values.
    filtered = {
        k: v for k, v in query.items()
        if k not in ("sign", "access_token") and v not in (None, "")
    }
    # Step 2: sort by key, concat as key+value with no separator.
    sorted_pairs = sorted(filtered.items(), key=lambda kv: kv[0])
    sorted_param_string = "".join(f"{k}{v}" for k, v in sorted_pairs)

    # Step 3-5: wrap and hash.
    canonical = f"{path}{sorted_param_string}{raw_body}"
    wrapped = f"{config.TIKTOKSHOP_APP_SECRET}{canonical}{config.TIKTOKSHOP_APP_SECRET}"
    return hmac.new(
        config.TIKTOKSHOP_APP_SECRET.encode("utf-8"),
        wrapped.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()


def _check_ok(response: requests.Response, *, context: str) -> None:
    """Asserts a successful TikTok Shop response or raises with platform detail."""
    if response.status_code != 200:
        raise RuntimeError(
            f"TikTok Shop HTTP {response.status_code} on {context}: {response.text[:500]}"
        )

    data = response.json()
    code = data.get("code")
    if code not in (0, "0"):
        raise RuntimeError(
            f"TikTok Shop API error on {context}: code={code} "
            f"message={data.get('message')}"
        )
