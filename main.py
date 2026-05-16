"""Main orchestrator for ITBisa sales analysis."""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from analysis import (aggregate_by_sku, calculate_hpp_wa, calculate_qty_setelah_restock,
                      enrich_with_profit, find_sku_without_hpp)
from config import DATA_DIR, JUAL_GLOB, OUTPUT_DIR, OUTPUT_FILENAME, STOK_GLOB
from data_loader import clean_jual, load_jual_files, load_stok_files
from excel_writer import write_report
from tables import (build_supplier_analysis, build_table_borderline, build_table_diminati,
                    build_table_kandidat, build_table_platform, build_table_profit,
                    build_table_rugi)


def _load_all(data_dir: Path):
    """Load stok + jual once; compute shared aggregates.
    Returns (stok, jual_full_clean, hpp_agg, qty_jual_all_time)."""
    stok_files = sorted(data_dir.glob(STOK_GLOB))
    jual_files = sorted(data_dir.glob(JUAL_GLOB))
    stok = load_stok_files(stok_files)
    jual_raw = load_jual_files(jual_files)
    jual_full_clean, _ = clean_jual(jual_raw, year=None)
    hpp_agg = calculate_hpp_wa(stok)
    qty_jual_all_time = jual_full_clean.groupby("SKU")["qty_jual"].sum()
    return stok, jual_full_clean, hpp_agg, qty_jual_all_time


def _analyze_year(stok, jual_full_clean, hpp_agg, qty_jual_all_time,
                  year: int, output_dir: Path):
    """Run analysis for a single year using pre-loaded data.
    Returns (path, profit, margin) or None if no data for that year."""
    jual_year = jual_full_clean[jual_full_clean["tanggal_pesan"].dt.year == year].copy()
    print(f"\n--- Tahun {year}: {len(jual_year):,} transaksi ---")

    if len(jual_year) == 0:
        print(f"  ⚠ Tidak ada data, skip")
        return None

    sku_no_hpp = find_sku_without_hpp(jual_year, hpp_agg)
    jual_with_profit = enrich_with_profit(jual_year, hpp_agg)

    sku_agg = aggregate_by_sku(jual_with_profit, hpp_agg, year, qty_jual_all_time)
    sku_agg = calculate_qty_setelah_restock(sku_agg, jual_with_profit)

    tables = {
        "diminati": build_table_diminati(sku_agg),
        "profit": build_table_profit(sku_agg),
        "rugi": build_table_rugi(sku_agg),
        "borderline": build_table_borderline(sku_agg),
        "kandidat": build_table_kandidat(sku_agg),
        "platform": build_table_platform(jual_with_profit),
        "supplier": build_supplier_analysis(stok, sku_agg),
    }

    output_path = output_dir / OUTPUT_FILENAME.format(year=year)
    write_report(output_path, year, jual_with_profit, sku_agg, tables, sku_no_hpp)

    total_profit = jual_with_profit["profit"].sum()
    omzet = jual_with_profit["omzet"].sum()
    margin = total_profit / omzet * 100 if omzet > 0 else 0
    return output_path, total_profit, margin


def run_analysis(year: int, data_dir: Path = DATA_DIR,
                 output_dir: Path = OUTPUT_DIR) -> Path:
    """Single-year analysis."""
    print(f"\n{'=' * 60}")
    print(f"ANALISA PENJUALAN ITBISA — TAHUN {year}")
    print(f"{'=' * 60}\n")

    stok, jual_full_clean, hpp_agg, qty_jual_all_time = _load_all(data_dir)
    result = _analyze_year(stok, jual_full_clean, hpp_agg, qty_jual_all_time,
                           year, output_dir)

    if result is None:
        raise ValueError(f"Tidak ada data jual untuk tahun {year}")

    path, profit, margin = result
    print(f"\n{'=' * 60}")
    print(f"✓ Selesai! Total profit {year}: Rp {profit:,.0f} (margin {margin:.1f}%)")
    print(f"  File output: {path}")
    print(f"{'=' * 60}\n")
    return path


def run_all_years(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR) -> list:
    """Generate reports for all years found in jual data."""
    print(f"\n{'=' * 60}")
    print(f"ANALISA PENJUALAN ITBISA — ALL YEARS")
    print(f"{'=' * 60}\n")

    stok, jual_full_clean, hpp_agg, qty_jual_all_time = _load_all(data_dir)

    years = sorted(int(y) for y in jual_full_clean["tanggal_pesan"].dt.year.dropna().unique())
    print(f"\n✓ Tahun ditemukan ({len(years)}): {', '.join(str(y) for y in years)}")

    results = []
    for year in years:
        try:
            res = _analyze_year(stok, jual_full_clean, hpp_agg, qty_jual_all_time,
                                year, output_dir)
            if res is not None:
                results.append((year,) + res)
        except Exception as e:
            print(f"  ✗ Error tahun {year}: {e}")

    print(f"\n{'=' * 60}")
    print(f"✓ Selesai: {len(results)} laporan dibuat di {output_dir}")
    print(f"{'=' * 60}")
    print(f"{'Tahun':>6}  {'Profit':>18}  {'Margin':>8}  File")
    for year, _, profit, margin in results:
        print(f"  {year:>4}  Rp {profit:>15,.0f}  {margin:>6.1f}%  "
              f"Analisa_Penjualan_ITBisa_{year}.xlsx")
    print()
    return results


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate ITBisa sales analysis Excel report.")
    parser.add_argument("--year", type=int, default=datetime.now().year,
                        help="Year to analyze (default: current year). Diabaikan kalau --all dipakai.")
    parser.add_argument("--all", action="store_true",
                        help="Generate laporan untuk SEMUA tahun yang ditemukan di data")
    parser.add_argument("--data-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    args = parser.parse_args()

    try:
        if args.all:
            run_all_years(args.data_dir, args.output_dir)
        else:
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
