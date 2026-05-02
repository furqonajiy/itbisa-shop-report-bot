"""
dump_reason_codes.py
--------------------
Thin CLI wrapper around src.discovery for one platform at a time.

For most workflows, prefer src/main.py — it supports --platform all
and integrates the auth health-check and Telegram heartbeat. This
script exists for the case where you want to walk just one platform
without invoking the full orchestrator.

Usage:
  python scripts/dump_reason_codes.py --platform shopee     --period 2026-01
  python scripts/dump_reason_codes.py --platform tiktokshop --period 2026-01

Output:
  - Pretty summary printed to stdout.
  - discovery_<platform>_<YYYY-MM>.json written to project root.
    Both are gitignored — they contain shop-internal financial data.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src import config, discovery


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--platform",
        required=True,
        choices=["shopee", "tiktokshop"],
        help="Which marketplace to dump.",
    )
    parser.add_argument(
        "--period",
        required=True,
        help="Month to walk, format YYYY-MM (e.g. 2026-01).",
    )
    args = parser.parse_args()

    year, month = config.parse_period_arg(args.period)
    start, end = config.month_bounds_utc(year, month)
    start_ts, end_ts = int(start.timestamp()), int(end.timestamp())

    print("=" * 70)
    print(f"Reason-code discovery — {args.platform} — {args.period}")
    print(f"  Window: {start.isoformat()}  →  {end.isoformat()} (exclusive)")
    print(f"  Unix:   {start_ts}            →  {end_ts}")
    print("=" * 70)

    if args.platform == "shopee":
        result = discovery.discover_shopee(start_ts, end_ts)
    else:
        result = discovery.discover_tiktokshop(start_ts, end_ts)

    discovery.print_summary(result)

    outfile = discovery.save_dump(result, args.period, PROJECT_ROOT)
    print(f"\nFull dump written to {outfile} (gitignored)")


if __name__ == "__main__":
    main()
