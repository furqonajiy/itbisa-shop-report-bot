"""Main orchestrator for ITBisa sales analysis."""
from __future__ import annotations
import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

from analysis import (aggregate_by_sku, calculate_hpp_wa, calculate_qty_setelah_restock,
                      compute_reorder_metrics, enrich_with_profit, find_sku_without_hpp)
from ab_testing import (analyze_ab_tests, create_template, load_ab_tests,
                        write_ab_test_report)
from config import (AB_TESTS_FILENAME, AB_TESTS_OUTPUT_FILENAME, DATA_DIR,
                    JUAL_GLOB, OUTPUT_DIR, OUTPUT_FILENAME,
                    REORDER_OUTPUT_FILENAME, STOK_GLOB)
from data_loader import clean_jual, load_jual_files, load_stok_files
from excel_writer import write_report, write_reorder_standalone
from tables import (build_reorder_tables, build_supplier_analysis,
                    build_table_borderline, build_table_diminati,
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
                  year: int, output_dir: Path, reorder_df=None, today=None):
    """Run analysis for a single year using pre-loaded data.
    `reorder_df` and `today` are optional and shared across years in --all mode.
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
    if reorder_df is not None:
        tables["reorder"] = build_reorder_tables(reorder_df)

    output_path = output_dir / OUTPUT_FILENAME.format(year=year)
    write_report(output_path, year, jual_with_profit, sku_agg, tables, sku_no_hpp,
                 today=today)

    total_profit = jual_with_profit["profit"].sum()
    omzet = jual_with_profit["omzet"].sum()
    margin = total_profit / omzet * 100 if omzet > 0 else 0
    return output_path, total_profit, margin


def run_analysis(year: int, data_dir: Path = DATA_DIR,
                 output_dir: Path = OUTPUT_DIR) -> Path:
    """Single-year analysis."""
    print(f"\n{'='*60}")
    print(f"ANALISA PENJUALAN ITBISA — TAHUN {year}")
    print(f"{'='*60}\n")

    stok, jual_full_clean, hpp_agg, qty_jual_all_time = _load_all(data_dir)
    today = pd.Timestamp(datetime.now().date())
    reorder_df = compute_reorder_metrics(stok, jual_full_clean, today)

    result = _analyze_year(stok, jual_full_clean, hpp_agg, qty_jual_all_time,
                           year, output_dir, reorder_df=reorder_df, today=today)

    if result is None:
        raise ValueError(f"Tidak ada data jual untuk tahun {year}")

    path, profit, margin = result
    print(f"\n{'='*60}")
    print(f"✓ Selesai! Total profit {year}: Rp {profit:,.0f} (margin {margin:.1f}%)")
    print(f"  File output: {path}")
    print(f"{'='*60}\n")
    return path


def run_all_years(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR) -> list:
    """Generate reports for all years found in jual data."""
    print(f"\n{'='*60}")
    print(f"ANALISA PENJUALAN ITBISA — ALL YEARS")
    print(f"{'='*60}\n")

    stok, jual_full_clean, hpp_agg, qty_jual_all_time = _load_all(data_dir)
    today = pd.Timestamp(datetime.now().date())
    reorder_df = compute_reorder_metrics(stok, jual_full_clean, today)

    years = sorted(int(y) for y in jual_full_clean["tanggal_pesan"].dt.year.dropna().unique())
    print(f"\n✓ Tahun ditemukan ({len(years)}): {', '.join(str(y) for y in years)}")

    results = []
    for year in years:
        try:
            res = _analyze_year(stok, jual_full_clean, hpp_agg, qty_jual_all_time,
                                year, output_dir, reorder_df=reorder_df, today=today)
            if res is not None:
                results.append((year,) + res)
        except Exception as e:
            print(f"  ✗ Error tahun {year}: {e}")

    print(f"\n{'='*60}")
    print(f"✓ Selesai: {len(results)} laporan dibuat di {output_dir}")
    print(f"{'='*60}")
    print(f"{'Tahun':>6}  {'Profit':>18}  {'Margin':>8}  File")
    for year, _, profit, margin in results:
        print(f"  {year:>4}  Rp {profit:>15,.0f}  {margin:>6.1f}%  "
              f"Analisa_Penjualan_ITBisa_{year}.xlsx")
    print()
    return results


def run_reorder(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR) -> Path:
    """Standalone reorder analysis report. Fast — skips yearly aggregation."""
    print(f"\n{'='*60}")
    print(f"ANALISA REORDER — KAPAN & BERAPA BANYAK")
    print(f"{'='*60}\n")

    stok, jual_full_clean, _hpp, _qty = _load_all(data_dir)
    today = pd.Timestamp(datetime.now().date())
    reorder_df = compute_reorder_metrics(stok, jual_full_clean, today)
    tables = build_reorder_tables(reorder_df)

    output_path = output_dir / REORDER_OUTPUT_FILENAME
    write_reorder_standalone(output_path, tables, today=today)

    print(f"\n{'='*60}")
    print(f"Ringkasan:")
    for label in ["🔴 STOCKOUT", "🔴 Reorder URGENT", "🟠 Reorder Now",
                  "🟡 Reorder Soon", "🟢 Healthy", "🔵 Overstock", "💤 Slow/Dead"]:
        n = (reorder_df["status"] == label).sum()
        print(f"  {label:>22}: {n:>4}")
    print(f"{'='*60}\n")
    return output_path


def run_ab_test(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR) -> Path:
    """Run A/B test analysis. Creates template if config file missing."""
    print(f"\n{'='*60}")
    print(f"ANALISA A/B TEST — PERUBAHAN HARGA")
    print(f"{'='*60}\n")

    ab_config_path = data_dir / AB_TESTS_FILENAME
    if not ab_config_path.exists():
        print(f"⚠ Config A/B test belum ada. Membuat template di {ab_config_path}")
        create_template(ab_config_path)
        print(f"\nEdit file tersebut, tambahkan SKU & tanggal perubahan, lalu run ulang.")
        return ab_config_path

    _stok, jual_full_clean, hpp_agg, _qty = _load_all(data_dir)
    ab_tests = load_ab_tests(ab_config_path)

    if len(ab_tests) == 0:
        print("⚠ Tidak ada test config valid di ab_tests.xlsx")
        return ab_config_path

    print(f"\n--- Menganalisa {len(ab_tests)} test ---")
    today = datetime.now()
    results = analyze_ab_tests(ab_tests, jual_full_clean, hpp_agg, today)

    output_path = output_dir / AB_TESTS_OUTPUT_FILENAME
    write_ab_test_report(output_path, results, today)

    print(f"\n{'='*60}")
    if len(results) > 0:
        print(f"Ringkasan {len(results)} test:")
        for _, r in results.iterrows():
            warn = f"  ({r['warning']})" if r['warning'] else ""
            print(f"  {r['sku']}: {r['verdict']}{warn}")
    print(f"{'='*60}\n")
    return output_path


def _run_ab_test_if_configured(data_dir: Path, output_dir: Path) -> Path | None:
    """For --all mode: run AB test only if template exists with rows. Skip silently otherwise."""
    ab_config_path = data_dir / AB_TESTS_FILENAME
    if not ab_config_path.exists():
        print(f"⊘ A/B test dilewati: {AB_TESTS_FILENAME} belum ada.")
        print(f"  Run 'python main.py --ab-test' untuk setup template.")
        return None

    ab_tests = load_ab_tests(ab_config_path)
    if len(ab_tests) == 0:
        print(f"⊘ A/B test dilewati: {AB_TESTS_FILENAME} kosong.")
        return None

    _stok, jual_full_clean, hpp_agg, _qty = _load_all(data_dir)
    print(f"\n--- Menganalisa {len(ab_tests)} test ---")
    today = datetime.now()
    results = analyze_ab_tests(ab_tests, jual_full_clean, hpp_agg, today)
    output_path = output_dir / AB_TESTS_OUTPUT_FILENAME
    write_ab_test_report(output_path, results, today)

    if len(results) > 0:
        print(f"\nRingkasan {len(results)} test:")
        for _, r in results.iterrows():
            warn = f"  ({r['warning']})" if r['warning'] else ""
            print(f"  {r['sku']}: {r['verdict']}{warn}")
    return output_path


def run_everything(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR) -> None:
    """Run sales all years + reorder standalone + ab-test (if configured)."""
    print(f"\n{'#'*60}")
    print(f"# RUN EVERYTHING — SALES (ALL YEARS) + REORDER + AB TEST")
    print(f"{'#'*60}")

    print(f"\n[1/3] Sales analysis untuk semua tahun")
    print(f"{'-'*60}")
    run_all_years(data_dir, output_dir)

    print(f"\n[2/3] Reorder analysis standalone")
    print(f"{'-'*60}")
    run_reorder(data_dir, output_dir)

    print(f"\n[3/3] A/B test")
    print(f"{'-'*60}")
    _run_ab_test_if_configured(data_dir, output_dir)

    print(f"\n{'#'*60}")
    print(f"# ✓ Selesai semua. Hasil di folder: {output_dir}")
    print(f"{'#'*60}\n")


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate ITBisa sales analysis Excel report.")
    parser.add_argument("--sales", nargs="?", const="all", default=None, metavar="YEAR",
                        help="Sales analysis. Tanpa argumen = semua tahun. "
                             "Dengan tahun (mis. --sales 2026) = tahun spesifik. "
                             "Tanpa flag = tahun berjalan.")
    parser.add_argument("--reorder", action="store_true",
                        help="Generate laporan reorder standalone (cepat, tanpa analisa tahunan).")
    parser.add_argument("--ab-test", action="store_true",
                        help="Generate laporan A/B test (perubahan harga). Otomatis bikin template kalau belum ada.")
    parser.add_argument("--all", action="store_true",
                        help="Run SEMUANYA: sales all years + reorder + ab-test (kalau template ada).")
    parser.add_argument("--data-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    args = parser.parse_args()

    try:
        if args.all:
            run_everything(args.data_dir, args.output_dir)
        elif args.ab_test:
            run_ab_test(args.data_dir, args.output_dir)
        elif args.reorder:
            run_reorder(args.data_dir, args.output_dir)
        elif args.sales is not None:
            if args.sales == "all":
                run_all_years(args.data_dir, args.output_dir)
            else:
                try:
                    year = int(args.sales)
                except ValueError:
                    print(f"❌ Tahun tidak valid: {args.sales}", file=sys.stderr)
                    return 1
                run_analysis(year, args.data_dir, args.output_dir)
        else:
            run_analysis(datetime.now().year, args.data_dir, args.output_dir)
        return 0
    except FileNotFoundError as e:
        print(f"❌ {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"❌ Error data: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())