"""
main.py
-------
CLI orchestrator for itbisa-shop-report-bot. Runs locally with:

    python -m src.main --period 2026-01
    python -m src.main --period 2026-01 --platform shopee
    python -m src.main --period 2026-01 --platform tiktokshop
    python -m src.main --period 2026-01 --no-save
    python -m src.main --period 2026-01 --no-telegram

PR 1 behaviour:
  For each requested platform:
    1. Auth health-check (get_valid_access_token() — refreshes if needed).
    2. Walks finance API for the period.
    3. Aggregates by reason code.
    4. Prints summary table to stdout.
    5. Saves discovery_<platform>_<YYYY-MM>.json to project root (gitignored).
    6. Sends a one-line Bahasa heartbeat to Telegram at end of run.

PR 2 will replace step 2-4 with full xlsx generation. The CLI surface
(--platform, --period) stays the same so workflow_dispatch from the
Worker (PR 3) keeps working without changes.

Exit codes:
  0  all platforms walked successfully
  1  any platform failed (token, network, API error)
"""

from __future__ import annotations

import argparse
import sys
import traceback
from datetime import datetime, timezone
from pathlib import Path

# Allow `python -m src.main` AND `python src/main.py` to both work.
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src import config, discovery, telegram_sender


PLATFORM_ALL = "all"
PLATFORM_CHOICES = ["shopee", "tiktokshop", PLATFORM_ALL]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--platform",
        choices=PLATFORM_CHOICES,
        default=PLATFORM_ALL,
        help="Marketplace to walk. 'all' runs both sequentially (default).",
    )
    parser.add_argument(
        "--period",
        required=True,
        help="Month to walk, format YYYY-MM (e.g. 2026-01).",
    )
    parser.add_argument(
        "--no-save",
        action="store_true",
        help="Skip writing discovery_<platform>_<period>.json.",
    )
    parser.add_argument(
        "--no-telegram",
        action="store_true",
        help="Skip the start/end Telegram heartbeats.",
    )
    args = parser.parse_args()

    # Resolve period.
    try:
        year, month = config.parse_period_arg(args.period)
    except ValueError as exc:
        print(f"ERROR: {exc}")
        return 1

    start, end = config.month_bounds_utc(year, month)
    start_ts = int(start.timestamp())
    end_ts   = int(end.timestamp())

    platforms = (
        ["shopee", "tiktokshop"] if args.platform == PLATFORM_ALL else [args.platform]
    )

    print("=" * 70)
    print(f"itbisa-shop-report-bot — PR 1 (discovery mode)")
    print(f"  Period:    {args.period}")
    print(f"  Window:    {start.isoformat()}  →  {end.isoformat()} (exclusive)")
    print(f"  Platforms: {', '.join(platforms)}")
    print("=" * 70)

    if not args.no_telegram:
        telegram_sender.send_message(
            f"🔍 Mulai discovery reason code\n"
            f"   Periode: {args.period}\n"
            f"   Platform: {', '.join(platforms)}"
        )

    failures: list[tuple[str, str]] = []

    for platform in platforms:
        print(f"\n{'━' * 70}\n► {platform.upper()}\n{'━' * 70}")
        try:
            _run_one_platform(
                platform=platform,
                start_ts=start_ts,
                end_ts=end_ts,
                period=args.period,
                save=not args.no_save,
            )
        except Exception as exc:
            print(f"\n✗ {platform} failed: {exc}")
            traceback.print_exc()
            failures.append((platform, str(exc)))

    # Final status — to console and Telegram.
    print(f"\n{'=' * 70}")
    if failures:
        lines = ["✗ Discovery selesai dengan error:"]
        for plat, msg in failures:
            lines.append(f"   • {plat}: {msg[:200]}")
        summary = "\n".join(lines)
    else:
        summary = (
            f"✅ Discovery selesai — {args.period}\n"
            f"   Platform: {', '.join(platforms)}\n"
            f"   Lihat output JSON di project root."
        )
    print(summary)
    print("=" * 70)

    if not args.no_telegram:
        telegram_sender.send_message(summary)

    return 1 if failures else 0


def _run_one_platform(
    *,
    platform: str,
    start_ts: int,
    end_ts: int,
    period: str,
    save: bool,
) -> None:
    """Health-checks auth, walks the finance API, prints + persists."""
    # STEP 1: Auth health-check. Surfaces token issues fast (before any
    # walk that might run for minutes). Both auth modules log the
    # refresh transparently if needed.
    print(f"\n[{platform}] Verifying access token...")
    if platform == "shopee":
        from src import shopee_auth
        access_token = shopee_auth.get_valid_access_token()
    elif platform == "tiktokshop":
        from src import tiktokshop_auth
        access_token = tiktokshop_auth.get_valid_access_token()
    else:
        raise ValueError(f"Unknown platform: {platform}")
    print(f"  ✓ token ok ({access_token[:12]}...)")

    # STEP 2: Discovery walk.
    if platform == "shopee":
        result = discovery.discover_shopee(start_ts, end_ts)
    else:
        result = discovery.discover_tiktokshop(start_ts, end_ts)

    # STEP 3: Print summary.
    discovery.print_summary(result)

    # STEP 4: Persist.
    if save:
        outfile = discovery.save_dump(result, period, PROJECT_ROOT)
        print(f"\nFull dump written to {outfile} (gitignored)")
    else:
        print("\n(--no-save: skipping JSON dump)")


if __name__ == "__main__":
    sys.exit(main())
