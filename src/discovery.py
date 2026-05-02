"""
discovery.py
------------
Reason-code discovery for both platforms. Public surface is consumed
by both src/main.py (orchestrator) and scripts/dump_reason_codes.py
(thin CLI wrapper).

Why two callers, one module:
  PR 1's only real "work" is walking finance APIs and aggregating
  distinct reason codes. We want both the script and main.py to
  share that logic without duplication.

Public functions:
  discover_shopee(start_ts, end_ts) -> dict
  discover_tiktokshop(start_ts, end_ts) -> dict
  print_summary(result) -> None
  save_dump(result, period, project_root) -> Path

Result dict shape (Shopee):
  {
    "platform":      "shopee",
    "total_rows":    int,
    "distinct_keys": int,
    "by_key": [
      {
        "transaction_type": str,
        "reason":           str,
        "count":            int,
        "sum_amount":       float,
        "samples":          [dict, ...],   # up to 3 sample rows, truncated
      },
      ...
    ],
  }

Result dict shape (TikTok):
  {
    "platform":           "tiktokshop",
    "statements_walked":  int,
    "total_rows":         int,
    "distinct_keys":      int,
    "field_frequencies":  {field_name: count_seen},
    "by_key": [
      {
        "type":           str,
        "adjustment_type": str,
        "count":           int,
        "sum_settlement":  float,
        "samples":         [dict, ...],
      },
      ...
    ],
  }
"""

from __future__ import annotations

import json
from collections import Counter, defaultdict
from pathlib import Path

from src import shopee_client, tiktokshop_client


# ============================================================
# Shopee
# ============================================================

def discover_shopee(start_ts: int, end_ts: int) -> dict:
    """
    Walks payment.get_wallet_transaction_list and aggregates by
    (transaction_type, reason). Each key gets a count, a summed amount,
    and up to 3 sample rows so we can see what real data looks like.
    """
    counter: Counter[tuple[str, str]] = Counter()
    samples: dict[tuple[str, str], list[dict]] = defaultdict(list)
    sum_amount: dict[tuple[str, str], float] = defaultdict(float)
    total_rows = 0

    print("\nWalking Shopee wallet transactions...")
    for row in shopee_client.get_wallet_transactions(start_ts, end_ts):
        total_rows += 1
        ttype  = str(row.get("transaction_type") or "(missing)")
        reason = str(row.get("reason") or "(missing)")
        key = (ttype, reason)
        counter[key] += 1

        try:
            sum_amount[key] += float(row.get("amount") or 0)
        except (TypeError, ValueError):
            pass

        if len(samples[key]) < 3:
            samples[key].append(_compact_row(row))

        if total_rows % 500 == 0:
            print(f"  ...{total_rows} rows so far ({len(counter)} distinct keys)")

    print(f"  Done: {total_rows} rows, {len(counter)} distinct (transaction_type, reason) pairs.")

    return {
        "platform":      "shopee",
        "total_rows":    total_rows,
        "distinct_keys": len(counter),
        "by_key": [
            {
                "transaction_type": ttype,
                "reason":           reason,
                "count":            counter[(ttype, reason)],
                "sum_amount":       round(sum_amount[(ttype, reason)], 2),
                "samples":          samples[(ttype, reason)],
            }
            for (ttype, reason), _ in counter.most_common()
        ],
    }


# ============================================================
# TikTok Shop
# ============================================================

def discover_tiktokshop(start_ts: int, end_ts: int) -> dict:
    """
    Walks finance/.../statements + statement_transactions and aggregates
    rows by (type, adjustment_type). Also tallies which keys appear
    in any row at all — handy for confirming the schema we'll code
    against in PR 2.
    """
    counter: Counter[tuple[str, str]] = Counter()
    samples: dict[tuple[str, str], list[dict]] = defaultdict(list)
    sum_settle: dict[tuple[str, str], float] = defaultdict(float)
    all_keys_seen: Counter[str] = Counter()
    statements_walked = 0
    rows_walked = 0

    print("\nWalking TikTok Shop statements...")
    for stmt in tiktokshop_client.iter_statements(start_ts, end_ts):
        statements_walked += 1
        stmt_id = stmt.get("id") or stmt.get("statement_id")
        if not stmt_id:
            print(f"  WARN: statement without id, skipping: {stmt}")
            continue

        for row in tiktokshop_client.iter_statement_transactions(str(stmt_id)):
            rows_walked += 1
            for k in row.keys():
                all_keys_seen[k] += 1

            ttype = str(row.get("type") or "(missing)")
            atype = str(row.get("adjustment_type") or "")
            key = (ttype, atype)
            counter[key] += 1

            try:
                sum_settle[key] += float(row.get("settlement_amount") or 0)
            except (TypeError, ValueError):
                pass

            if len(samples[key]) < 3:
                samples[key].append(_compact_row(row))

            if rows_walked % 200 == 0:
                print(f"  ...{rows_walked} rows ({statements_walked} statements, {len(counter)} keys)")

    print(
        f"  Done: {statements_walked} statements, {rows_walked} rows, "
        f"{len(counter)} distinct (type, adjustment_type) pairs."
    )

    return {
        "platform":          "tiktokshop",
        "statements_walked": statements_walked,
        "total_rows":        rows_walked,
        "distinct_keys":     len(counter),
        "field_frequencies": dict(all_keys_seen),
        "by_key": [
            {
                "type":            ttype,
                "adjustment_type": atype,
                "count":           counter[(ttype, atype)],
                "sum_settlement":  round(sum_settle[(ttype, atype)], 2),
                "samples":         samples[(ttype, atype)],
            }
            for (ttype, atype), _ in counter.most_common()
        ],
    }


# ============================================================
# Output
# ============================================================

def print_summary(result: dict) -> None:
    """Renders a summary table that pastes well into chat."""
    print("\n" + "─" * 70)
    print(f"SUMMARY  ({result['platform']})")
    print("─" * 70)
    print(f"Total rows:     {result.get('total_rows', 0)}")
    if "statements_walked" in result:
        print(f"Statements:     {result['statements_walked']}")
    print(f"Distinct keys:  {result['distinct_keys']}")
    print()

    if result["platform"] == "shopee":
        print(f"{'COUNT':>8}  {'SUM_AMOUNT':>14}  TRANSACTION_TYPE / REASON")
        print(f"{'-----':>8}  {'-----------':>14}  ---------------------------")
        for entry in result["by_key"]:
            print(
                f"{entry['count']:>8}  {entry['sum_amount']:>14,.2f}  "
                f"{entry['transaction_type']} / {entry['reason']}"
            )
    else:
        print(f"{'COUNT':>8}  {'SUM_SETTLE':>14}  TYPE / ADJUSTMENT_TYPE")
        print(f"{'-----':>8}  {'-----------':>14}  ----------------------")
        for entry in result["by_key"]:
            atype = entry["adjustment_type"] or "(none)"
            print(
                f"{entry['count']:>8}  {entry['sum_settlement']:>14,.2f}  "
                f"{entry['type']} / {atype}"
            )

        print("\nField frequencies (keys observed in transaction rows):")
        for k, n in sorted(result["field_frequencies"].items()):
            print(f"  {n:>6}  {k}")


def save_dump(result: dict, period: str, project_root: Path) -> Path:
    """Persists the full result as discovery_<platform>_<period>.json (gitignored)."""
    outfile = project_root / f"discovery_{result['platform']}_{period}.json"
    with open(outfile, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False, default=str)
    return outfile


# ============================================================
# Internal
# ============================================================

def _compact_row(row: dict) -> dict:
    """Truncates long string values to keep dumps human-readable."""
    out = {}
    for k, v in row.items():
        if isinstance(v, str) and len(v) > 120:
            out[k] = v[:117] + "..."
        else:
            out[k] = v
    return out
