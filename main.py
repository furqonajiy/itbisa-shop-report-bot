"""Main orchestrator for ITBisa sales analysis report.

Usage:
    python main.py                       # analyze current year
    python main.py --year 2024           # analyze specific year
    python main.py --year 2024 --data-dir /custom/path
"""
from __future__ import annotations
import argparse
import sys
from datetime import datetime
from pathlib import Path

from config import (DATA_DIR, OUTPUT_DIR, STOK_FILENAME, JUAL_FILENAME, OUTPUT_FILENAME)
from data_loader import clean_jual, load_jual, load_stok
from analysis import (aggregate_by_sku, calculate_hpp_wa, calculate_qty_setelah_restock,
                       enrich_with_profit, find_sku_without_hpp)
from tables import (build_table_borderline, build_table_diminati, build_table_kandidat,
                     build_table_platform, build_table_profit, build_table_rugi)
from excel_writer import write_report


def run_analysis(year: int, data_dir: Path = DATA_DIR,
                 output_dir: Path = OUTPUT_DIR) -> Path:
    """Run end-to-end analysis pipeline. Returns path to generated Excel."""
    stok_path = data_dir / STOK_FILENAME.format(year=year)
    jual_path = data_dir / JUAL_FILENAME.format(year=year)
    output_path = output_dir / OUTPUT_FILENAME.format(year=year)

    print(f"\n{'='*60}")
    print(f"ANALISA PENJUALAN ITBISA — TAHUN {year}")
    print(f"{'='*60}\n")

    stok = load_stok(stok_path)
    jual_raw = load_jual(jual_path)
    jual_clean, _ = clean_jual(jual_raw)

    hpp_agg = calculate_hpp_wa(stok)
    sku_no_hpp = find_sku_without_hpp(jual_clean, hpp_agg)
    jual_with_profit = enrich_with_profit(jual_clean, hpp_agg)

    print("✓ Aggregasi per SKU")
    sku_agg = aggregate_by_sku(jual_with_profit, hpp_agg, year)
    sku_agg = calculate_qty_setelah_restock(sku_agg, jual_with_profit)

    print("✓ Membangun tabel analisa")
    tables = {
        "diminati": build_table_diminati(sku_agg),
        "profit": build_table_profit(sku_agg),
        "rugi": build_table_rugi(sku_agg),
        "borderline": build_table_borderline(sku_agg),
        "kandidat": build_table_kandidat(sku_agg),
        "platform": build_table_platform(jual_with_profit),
    }

    write_report(output_path, year, jual_with_profit, sku_agg, tables, sku_no_hpp)

    total_profit = jual_with_profit["profit"].sum()
    margin = total_profit / jual_with_profit["omzet"].sum() * 100
    print(f"\n{'='*60}")
    print(f"✓ Selesai! Total profit {year}: Rp {total_profit:,.0f} (margin {margin:.1f}%)")
    print(f"  File output: {output_path}")
    print(f"{'='*60}\n")
    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate ITBisa sales analysis Excel report."
    )
    parser.add_argument("--year", type=int, default=datetime.now().year,
                        help="Year to analyze (default: current year)")
    parser.add_argument("--data-dir", type=Path, default=DATA_DIR,
                        help=f"Input data directory (default: {DATA_DIR})")
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR,
                        help=f"Output directory (default: {OUTPUT_DIR})")
    args = parser.parse_args()

    try:
        run_analysis(args.year, args.data_dir, args.output_dir)
        return 0
    except FileNotFoundError as e:
        print(f"❌ {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"❌ Error data: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
