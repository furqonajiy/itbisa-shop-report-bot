"""
shopee_client.py
----------------
Shopee Open API integration for the report bot.

PR 1 surface:
  describe()
    Identifier for log/Telegram headers.

  get_wallet_transactions(start_ts_unix, end_ts_unix) -> Iterator[dict]
    Paginates payment.get_wallet_transaction_list across the period.
    Each yielded dict is one transaction row exactly as Shopee returns
    it. Used by scripts/dump_reason_codes.py to discover all distinct
    transaction_type / reason values for a real shop.

PR 2 will add:
  get_orders_for_period()        — order.get_order_list + get_order_detail
  get_escrow_detail(order_sn)    — payment.get_escrow_detail

Internal helpers (also reusable by PR 2):
  _call_signed(method, path, ...) -> requests.Response
  _check_ok(response, context)
"""

from __future__ import annotations

import hashlib
import hmac
import json as _json
import time
from typing import Iterator

import requests

from src import config, shopee_auth


# ============================================================
# Public surface
# ============================================================

def describe() -> str:
    return "Shopee"


def get_wallet_transactions(
    start_ts_unix: int,
    end_ts_unix: int,
) -> Iterator[dict]:
    """
    Yields every wallet transaction in [start_ts_unix, end_ts_unix).

    Shopee endpoint: POST /api/v2/payment/get_wallet_transaction_list
    Body fields used:
      create_time_from: int (inclusive)
      create_time_to:   int (exclusive)
      page_no:          int (1-indexed)
      page_size:        int (max 100 per docs)

    Each yielded transaction dict typically contains:
      transaction_type   ('ORDER_INCOME', 'WITHDRAWAL', 'ADJUSTMENT', ...)
      reason             (free-text Bahasa)
      amount, current_balance
      create_time        (Unix int)
      order_sn           (when applicable)
      money_flow         ('Money In' / 'Money Out')

    Raises RuntimeError on platform-level error, or RefreshTokenExpiredError
    if the refresh chain is dead.
    """
    path = "/api/v2/payment/get_wallet_transaction_list"
    page_no = 1
    page_size = config.SHOPEE_WALLET_PAGE_SIZE

    while True:
        body = {
            "page_no":          page_no,
            "page_size":        page_size,
            "create_time_from": int(start_ts_unix),
            "create_time_to":   int(end_ts_unix),
        }
        response = _call_signed("POST", path, body=body)
        _check_ok(response, context=f"wallet_transaction_list page={page_no}")

        payload = response.json().get("response") or {}
        rows = payload.get("transaction_list") or []
        for row in rows:
            yield row

        # Shopee returns 'more' (bool) to indicate further pages.
        # Some older docs use len(rows) < page_size; we honour both.
        more = payload.get("more")
        if more is False or (more is None and len(rows) < page_size):
            return

        page_no += 1
        time.sleep(config.DELAY_BETWEEN_CALLS_SECONDS)


# ============================================================
# Internal: signed transport (shop-level signature format)
# ============================================================

def _call_signed(
    method: str,
    path: str,
    *,
    body: dict | list | None = None,
    extra_query: dict[str, str] | None = None,
) -> requests.Response:
    """
    Signs and dispatches a shop-level Open API call.

    Signature format (NOT the auth-endpoint format — that one lives
    in shopee_auth.py):
      base = partner_id + path + timestamp + access_token + shop_id
      sign = HMAC-SHA256(partner_key, base).hexdigest()
    """
    access_token = shopee_auth.get_valid_access_token()
    timestamp = int(time.time())

    base = (
        f"{config.SHOPEE_PARTNER_ID}"
        f"{path}"
        f"{timestamp}"
        f"{access_token}"
        f"{config.SHOPEE_SHOP_ID}"
    )
    sign = hmac.new(
        config.SHOPEE_PARTNER_KEY.encode("utf-8"),
        base.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()

    query = {
        "partner_id":   config.SHOPEE_PARTNER_ID,
        "timestamp":    str(timestamp),
        "access_token": access_token,
        "shop_id":      config.SHOPEE_SHOP_ID,
        "sign":         sign,
    }
    if extra_query:
        query.update(extra_query)

    url = f"{config.SHOPEE_API_BASE_URL}{path}"

    headers = {"Content-Type": "application/json"}
    if method.upper() == "POST":
        return requests.post(
            url,
            params=query,
            data=_json.dumps(body or {}, separators=(",", ":")),
            headers=headers,
            timeout=60,
        )
    return requests.get(url, params=query, timeout=60)


def _check_ok(response: requests.Response, *, context: str) -> None:
    """Asserts a successful Shopee response or raises with platform detail."""
    if response.status_code != 200:
        raise RuntimeError(
            f"Shopee HTTP {response.status_code} on {context}: {response.text[:500]}"
        )

    data = response.json()
    err = data.get("error")
    if err:
        msg = data.get("message", "")
        raise RuntimeError(f"Shopee API error on {context}: {err}: {msg}")
