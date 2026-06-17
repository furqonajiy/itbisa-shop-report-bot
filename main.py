"""Main orchestrator for ITBisa sales analysis."""
from __future__ import annotations
import argparse
import sys
from collections import namedtuple
from datetime import datetime
from pathlib import Path

import pandas as pd

from analysis import (aggregate_by_sku, build_stock_ledger, calculate_hpp_wa,
                      calculate_qty_setelah_restock, compute_harga_sekarang,
                      compute_price_change_status, compute_reorder_metrics,
                      enrich_with_profit, find_sku_without_hpp)
from ab_testing import (analyze_ab_tests, create_template, load_ab_change_dates,
                        load_ab_tests, write_ab_test_report)
from restock_pricing import (analyze_restock, compute_platform_fees,
                             compute_rmb_factor, create_restock_template,
                             load_restock_check, load_rmb_hpp_history,
                             write_restock_report)
from cashflow import (build_restock_plan, pivot_month_supplier, summarize_by_month,
                      write_cashflow_report, _months_axis)
from basket_analysis import analyze_baskets, write_basket_report
from deadstock_analysis import analyze_deadstock, write_deadstock_report
from momentum_analysis import analyze_momentum, write_momentum_report
from trend_analysis import analyze_trend, write_trend_report
from config import (AB_TESTS_FILENAME, AB_TESTS_OUTPUT_FILENAME,
                    BASKET_OUTPUT_FILENAME, CASHFLOW_HORIZON_MONTHS,
                    CASHFLOW_OUTPUT_FILENAME, DATA_DIR,
                    DEADSTOCK_OUTPUT_FILENAME, JUAL_GLOB,
                    MOMENTUM_OUTPUT_FILENAME, OUTPUT_DIR, OUTPUT_FILENAME,
                    REORDER_OUTPUT_FILENAME, RESTOCK_CHECK_FILENAME,
                    RESTOCK_OUTPUT_FILENAME, STOK_GLOB, TREND_OUTPUT_FILENAME)
from data_loader import (clean_jual, latest_file, load_current_jual_nonvoid,
                         load_current_stok_arrived, load_hilang, load_jual_files,
                         load_pindah, load_stok_files)
from excel_writer import write_report, write_reorder_standalone
from tables import (build_reorder_tables, build_supplier_analysis,
                    build_table_borderline, build_table_diminati,
                    build_table_kandidat, build_table_platform, build_table_profit,
                    build_table_rugi)


def _load_all(data_dir: Path):
    """Load stok + jual once; compute shared aggregates.
    Returns (stok, jual_full_clean, hpp_agg, qty_jual_all_time, sisa_by_sku, ledger_df).
    sisa_by_sku/ledger_df reconcile to BisaRekapBarang from the CURRENT workbook
    (latest stok + latest jual file by filename)."""
    stok_files = sorted(data_dir.glob(STOK_GLOB))
    jual_files = sorted(data_dir.glob(JUAL_GLOB))
    stok = load_stok_files(stok_files)
    jual_raw = load_jual_files(jual_files)
    jual_full_clean, _ = clean_jual(jual_raw, year=None)
    hpp_agg = calculate_hpp_wa(stok)
    qty_jual_all_time = jual_full_clean.groupby("SKU")["qty_jual"].sum()

    # --- Current-workbook stock ledger (authoritative sisa_stok) ---
    cur_stok_file = latest_file(stok_files)
    cur_jual_file = latest_file(jual_files)
    print(f"✓ Workbook berjalan: {cur_stok_file.name} + {cur_jual_file.name}")
    stok_arrived = load_current_stok_arrived(cur_stok_file)
    jual_cur = load_current_jual_nonvoid(cur_jual_file)
    hilang = load_hilang(cur_stok_file)
    pindah = load_pindah(cur_stok_file)
    ledger_df, sisa_by_sku = build_stock_ledger(stok_arrived, jual_cur, hilang, pindah)

    return stok, jual_full_clean, hpp_agg, qty_jual_all_time, sisa_by_sku, ledger_df


# Shared, already-computed inputs passed between steps in a full run so the Excel
# files are read once and reorder/AB-change dates are computed once (not per step).
_Loaded = namedtuple(
    "_Loaded",
    "stok jual hpp_agg qty_jual sisa ledger today reorder ab_changes")


def _load_shared(data_dir: Path) -> _Loaded:
    """Load stok+jual once and compute the shared aggregates (HPP, ledger, reorder,
    A/B change dates) a single time. A full run (`--all` / no flag) builds this once
    and hands it to every step, instead of each step re-reading the workbooks."""
    stok, jual, hpp_agg, qty_jual, sisa, ledger = _load_all(data_dir)
    today = pd.Timestamp(datetime.now().date())
    reorder_df = compute_reorder_metrics(stok, jual, today, sisa_by_sku=sisa)
    ab_changes = load_ab_change_dates(data_dir / AB_TESTS_FILENAME)
    return _Loaded(stok, jual, hpp_agg, qty_jual, sisa, ledger, today, reorder_df, ab_changes)


def _analyze_year(stok, jual_full_clean, hpp_agg, qty_jual_all_time,
                  year: int, output_dir: Path, today=None,
                  sisa_by_sku=None, ledger_df=None, ab_changes=None):
    """Run analysis for a single year using pre-loaded data.
    `today`, `sisa_by_sku`, `ledger_df`, `ab_changes` are shared across years in --all.
    The yearly file is pure sales history (sheets 00–08); reorder + per-gudang stock
    (a current snapshot, not year-specific) live only in Analisa_Reorder.xlsx.
    Returns (path, profit, margin) or None if no data for that year."""
    if today is None:
        today = pd.Timestamp(datetime.now().date())
    jual_year = jual_full_clean[jual_full_clean["tanggal_pesan"].dt.year == year].copy()
    print(f"\n--- Tahun {year}: {len(jual_year):,} transaksi ---")

    if len(jual_year) == 0:
        print(f"  ⚠ Tidak ada data, skip")
        return None

    sku_no_hpp = find_sku_without_hpp(jual_year, hpp_agg)
    jual_with_profit = enrich_with_profit(jual_year, hpp_agg)

    # 'Harga sekarang' uses the SKU's most recent selling day across ALL years
    # (like sisa/reorder), so the year filter doesn't truncate the latest price.
    harga_sekarang = compute_harga_sekarang(jual_full_clean)
    # Recent-price-increase guard: hold the Kandidat recommendation for SKUs whose
    # current price is a freshly-raised, under-validated price (qty earned at the old).
    price_change = compute_price_change_status(jual_full_clean, harga_sekarang,
                                               ab_changes, today, year)
    sku_agg = aggregate_by_sku(jual_with_profit, hpp_agg, year, qty_jual_all_time,
                               sisa_by_sku=sisa_by_sku, harga_sekarang=harga_sekarang,
                               price_change=price_change)
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
    write_report(output_path, year, jual_with_profit, sku_agg, tables, sku_no_hpp,
                 today=today, ledger_df=ledger_df)

    total_profit = jual_with_profit["profit"].sum()
    omzet = jual_with_profit["omzet"].sum()
    margin = total_profit / omzet * 100 if omzet > 0 else 0
    return output_path, total_profit, margin


def run_analysis(year: int, data_dir: Path = DATA_DIR,
                 output_dir: Path = OUTPUT_DIR, loaded: _Loaded | None = None) -> Path:
    """Single-year analysis."""
    print(f"\n{'='*60}")
    print(f"ANALISA PENJUALAN ITBISA — TAHUN {year}")
    print(f"{'='*60}\n")

    if loaded is None:
        loaded = _load_shared(data_dir)
    result = _analyze_year(loaded.stok, loaded.jual, loaded.hpp_agg, loaded.qty_jual,
                           year, output_dir, today=loaded.today,
                           sisa_by_sku=loaded.sisa, ledger_df=loaded.ledger,
                           ab_changes=loaded.ab_changes)

    if result is None:
        raise ValueError(f"Tidak ada data jual untuk tahun {year}")

    path, profit, margin = result
    print(f"\n{'='*60}")
    print(f"✓ Selesai! Total profit {year}: Rp {profit:,.0f} (margin {margin:.1f}%)")
    print(f"  File output: {path}")
    print(f"{'='*60}\n")
    return path


def run_all_years(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR,
                  loaded: _Loaded | None = None) -> list:
    """Generate reports for all years found in jual data."""
    print(f"\n{'='*60}")
    print(f"ANALISA PENJUALAN ITBISA — ALL YEARS")
    print(f"{'='*60}\n")

    if loaded is None:
        loaded = _load_shared(data_dir)
    jual_full_clean = loaded.jual

    years = sorted(int(y) for y in jual_full_clean["tanggal_pesan"].dt.year.dropna().unique())
    print(f"\n✓ Tahun ditemukan ({len(years)}): {', '.join(str(y) for y in years)}")

    results = []
    for year in years:
        try:
            res = _analyze_year(loaded.stok, jual_full_clean, loaded.hpp_agg, loaded.qty_jual,
                                year, output_dir, today=loaded.today,
                                sisa_by_sku=loaded.sisa, ledger_df=loaded.ledger,
                                ab_changes=loaded.ab_changes)
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


def run_reorder(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR,
                loaded: _Loaded | None = None) -> Path:
    """Standalone reorder analysis report. Fast — skips yearly aggregation."""
    print(f"\n{'='*60}")
    print(f"ANALISA REORDER — KAPAN & BERAPA BANYAK")
    print(f"{'='*60}\n")

    if loaded is None:
        loaded = _load_shared(data_dir)
    reorder_df = loaded.reorder
    tables = build_reorder_tables(reorder_df)

    output_path = output_dir / REORDER_OUTPUT_FILENAME
    write_reorder_standalone(output_path, tables, today=loaded.today, ledger_df=loaded.ledger)

    print(f"\n{'='*60}")
    print(f"Ringkasan:")
    for label in ["🔴 STOCKOUT", "🔴 Reorder URGENT", "🟠 Reorder Now",
                  "🟡 Reorder Soon", "🟢 Healthy", "🔵 Overstock", "💤 Slow/Dead"]:
        n = (reorder_df["status"] == label).sum()
        print(f"  {label:>22}: {n:>4}")
    print(f"{'='*60}\n")
    return output_path


def run_cashflow(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR,
                 loaded: _Loaded | None = None) -> Path:
    """Cash-flow restock plan: how much capital is needed, and when, to keep stock.
    Built from the reorder metrics + replacement HPP — no template needed."""
    print(f"\n{'='*60}")
    print(f"RENCANA CASH-FLOW RESTOCK — MODAL BELI")
    print(f"{'='*60}\n")

    if loaded is None:
        loaded = _load_shared(data_dir)
    today = loaded.today

    plan = build_restock_plan(loaded.reorder, loaded.hpp_agg, loaded.stok, today,
                              CASHFLOW_HORIZON_MONTHS)
    months = _months_axis(today, CASHFLOW_HORIZON_MONTHS)
    monthly = summarize_by_month(plan, months)
    pivot = pivot_month_supplier(plan, months)

    output_path = output_dir / CASHFLOW_OUTPUT_FILENAME
    write_cashflow_report(output_path, plan, monthly, pivot, months, today,
                          CASHFLOW_HORIZON_MONTHS)

    total = float(plan["order_cost"].sum(skipna=True)) if len(plan) else 0.0
    n_sku = int(plan["SKU"].nunique()) if len(plan) else 0
    print(f"\n{'='*60}")
    print(f"  Total modal restock {CASHFLOW_HORIZON_MONTHS} bln: Rp {total:,.0f} "
          f"({n_sku} SKU, {len(plan)} order)")
    this_month = months[0]
    due_now = (float(plan[plan['order_month'] == this_month]['order_cost'].sum(skipna=True))
               if len(plan) else 0.0)
    print(f"  Jatuh tempo bulan ini ({this_month}): Rp {due_now:,.0f}")
    print(f"{'='*60}\n")
    return output_path


def run_bundle(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR,
               loaded: _Loaded | None = None) -> Path:
    """Bundle / cross-sell market basket: SKUs frequently bought together."""
    print(f"\n{'='*60}")
    print(f"BUNDLE & CROSS-SELL — SERING DIBELI BERSAMA")
    print(f"{'='*60}\n")

    if loaded is None:
        loaded = _load_shared(data_dir)
    pairs, cross, stats = analyze_baskets(loaded.jual)
    output_path = output_dir / BASKET_OUTPUT_FILENAME
    write_basket_report(output_path, pairs, cross, stats, loaded.today)
    return output_path


def run_deadstock(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR,
                  loaded: _Loaded | None = None) -> Path:
    """Dead-stock / capital-release: capital frozen in slow/dead/overstock + how to free it."""
    print(f"\n{'='*60}")
    print(f"MODAL BEKU — KAPITAL TERTAHAN DI STOK LAMBAT/MATI")
    print(f"{'='*60}\n")

    if loaded is None:
        loaded = _load_shared(data_dir)
    df = analyze_deadstock(loaded.reorder, loaded.hpp_agg, loaded.jual, loaded.stok, loaded.today)
    output_path = output_dir / DEADSTOCK_OUTPUT_FILENAME
    write_deadstock_report(output_path, df, loaded.today)
    return output_path


def run_momentum(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR,
                 loaded: _Loaded | None = None) -> Path:
    """Momentum + ABC focus: which SKUs to push (rising winners) vs prune (declining tail)."""
    print(f"\n{'='*60}")
    print(f"MOMENTUM & ABC — APA YANG DIDORONG vs DIPANGKAS")
    print(f"{'='*60}\n")

    if loaded is None:
        loaded = _load_shared(data_dir)
    df = analyze_momentum(loaded.jual, loaded.hpp_agg, loaded.today)
    output_path = output_dir / MOMENTUM_OUTPUT_FILENAME
    write_momentum_report(output_path, df, loaded.today)
    return output_path


def run_trend(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR,
              loaded: _Loaded | None = None) -> Path:
    """Sales trend & seasonality: cross-year omzet/profit trend, YoY growth, seasonal index."""
    print(f"\n{'='*60}")
    print(f"TREN & MUSIMAN PENJUALAN — LINTAS TAHUN")
    print(f"{'='*60}\n")

    if loaded is None:
        loaded = _load_shared(data_dir)
    data = analyze_trend(loaded.jual, loaded.hpp_agg, loaded.today)
    output_path = output_dir / TREND_OUTPUT_FILENAME
    write_trend_report(output_path, data, loaded.today)
    return output_path


def run_restock_check(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR) -> Path:
    """Evaluate offered restock prices and recommend selling prices per marketplace.
    Reads data/restock_check.xlsx (auto-creates template if missing)."""
    print(f"\n{'='*60}")
    print(f"ANALISA HARGA RESTOCK — BELI & JUAL")
    print(f"{'='*60}\n")

    rc_path = data_dir / RESTOCK_CHECK_FILENAME
    if not rc_path.exists():
        print(f"⚠ Config restock belum ada. Membuat template di {rc_path}")
        create_restock_template(rc_path)
        print(f"\nEdit file tersebut (SKU, Harga RMB/HPP IDR, Toko, Kompetitor), lalu run ulang.")
        return rc_path

    checks = load_restock_check(rc_path)
    if len(checks) == 0:
        print(f"⚠ {RESTOCK_CHECK_FILENAME} kosong — tidak ada SKU untuk dicek.")
        return rc_path

    stok_files = sorted(data_dir.glob(STOK_GLOB))
    stok = load_stok_files(stok_files)
    jual = clean_jual(load_jual_files(sorted(data_dir.glob(JUAL_GLOB))), year=None)[0]
    hpp_agg = calculate_hpp_wa(stok)
    hist = load_rmb_hpp_history(stok_files)
    fees = compute_platform_fees(jual)
    _per_sku, factor_global = compute_rmb_factor(hist)
    print(f"  → Fee marketplace (dari data): "
          + ", ".join(f"{p} {fees[p]*100:.0f}%" for p in fees))
    print(f"  → Prediksi HPP landed: 1 RMB ≈ Rp{factor_global:,.0f} "
          f"(termasuk margin Martkita + ongkir + impor; kalibrasi histori Ocistok/Martkita)")

    results = analyze_restock(checks, hpp_agg, hist, fees)
    today = datetime.now()
    output_path = output_dir / RESTOCK_OUTPUT_FILENAME
    write_restock_report(output_path, results, fees, factor_global, today)

    print(f"\n{'='*60}")
    for _, r in results.iterrows():
        print(f"  {r['SKU']}: {r.get('keputusan','')}")
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

    _stok, jual_full_clean, hpp_agg, _qty, _sisa, _ledger = _load_all(data_dir)
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


def _run_ab_test_if_configured(data_dir: Path, output_dir: Path,
                               loaded: _Loaded | None = None) -> Path | None:
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

    if loaded is None:
        loaded = _load_shared(data_dir)
    jual_full_clean, hpp_agg = loaded.jual, loaded.hpp_agg
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


def _run_restock_check_if_configured(data_dir: Path, output_dir: Path,
                                     loaded: _Loaded | None = None) -> Path | None:
    """For --all mode: run restock-check only if the template exists with rows.
    Skip silently otherwise (don't create-a-template-and-halt like the standalone)."""
    rc_path = data_dir / RESTOCK_CHECK_FILENAME
    if not rc_path.exists():
        print(f"⊘ Restock-check dilewati: {RESTOCK_CHECK_FILENAME} belum ada.")
        print(f"  Run 'python main.py --restock-check' untuk setup template.")
        return None

    checks = load_restock_check(rc_path)
    if len(checks) == 0:
        print(f"⊘ Restock-check dilewati: {RESTOCK_CHECK_FILENAME} kosong.")
        return None

    stok_files = sorted(data_dir.glob(STOK_GLOB))
    if loaded is None:
        stok = load_stok_files(stok_files)
        jual = clean_jual(load_jual_files(sorted(data_dir.glob(JUAL_GLOB))), year=None)[0]
        hpp_agg = calculate_hpp_wa(stok)
    else:
        jual, hpp_agg = loaded.jual, loaded.hpp_agg
    hist = load_rmb_hpp_history(stok_files)
    fees = compute_platform_fees(jual)
    _per_sku, factor_global = compute_rmb_factor(hist)
    print(f"  → Fee marketplace (dari data): "
          + ", ".join(f"{p} {fees[p]*100:.0f}%" for p in fees))
    print(f"  → Prediksi HPP landed: 1 RMB ≈ Rp{factor_global:,.0f} "
          f"(termasuk margin Martkita + ongkir + impor; kalibrasi histori Ocistok/Martkita)")

    results = analyze_restock(checks, hpp_agg, hist, fees)
    today = datetime.now()
    output_path = output_dir / RESTOCK_OUTPUT_FILENAME
    write_restock_report(output_path, results, fees, factor_global, today)
    for _, r in results.iterrows():
        print(f"  {r['SKU']}: {r.get('keputusan','')}")
    return output_path


def run_everything(data_dir: Path = DATA_DIR, output_dir: Path = OUTPUT_DIR) -> None:
    """Run sales + trend + reorder + cash-flow + bundle + dead-stock + momentum
    + ab-test + restock-check. Loads the workbooks once and shares them."""
    print(f"\n{'#'*60}")
    print(f"# RUN EVERYTHING — SALES + TREND + REORDER + CASH-FLOW + BUNDLE")
    print(f"#   + DEAD-STOCK + MOMENTUM + AB + RESTOCK")
    print(f"{'#'*60}")

    print(f"\n[0/9] Memuat data (sekali untuk semua langkah)")
    print(f"{'-'*60}")
    loaded = _load_shared(data_dir)

    print(f"\n[1/9] Sales analysis untuk semua tahun")
    print(f"{'-'*60}")
    run_all_years(data_dir, output_dir, loaded=loaded)

    print(f"\n[2/9] Tren & musiman penjualan")
    print(f"{'-'*60}")
    run_trend(data_dir, output_dir, loaded=loaded)

    print(f"\n[3/9] Reorder analysis standalone")
    print(f"{'-'*60}")
    run_reorder(data_dir, output_dir, loaded=loaded)

    print(f"\n[4/9] Cash-flow restock plan")
    print(f"{'-'*60}")
    run_cashflow(data_dir, output_dir, loaded=loaded)

    print(f"\n[5/9] Bundle & cross-sell")
    print(f"{'-'*60}")
    run_bundle(data_dir, output_dir, loaded=loaded)

    print(f"\n[6/9] Modal beku (dead-stock / capital release)")
    print(f"{'-'*60}")
    run_deadstock(data_dir, output_dir, loaded=loaded)

    print(f"\n[7/9] Momentum & ABC focus")
    print(f"{'-'*60}")
    run_momentum(data_dir, output_dir, loaded=loaded)

    print(f"\n[8/9] A/B test")
    print(f"{'-'*60}")
    _run_ab_test_if_configured(data_dir, output_dir, loaded=loaded)

    print(f"\n[9/9] Restock price check")
    print(f"{'-'*60}")
    _run_restock_check_if_configured(data_dir, output_dir, loaded=loaded)

    print(f"\n{'#'*60}")
    print(f"# ✓ Selesai semua. Hasil di folder: {output_dir}")
    print(f"{'#'*60}\n")


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate ITBisa sales analysis Excel report.")
    parser.add_argument("--sales", nargs="?", const="all", default=None, metavar="YEAR",
                        help="Sales analysis. Tanpa argumen = semua tahun. "
                             "Dengan tahun (mis. --sales 2026) = tahun spesifik. "
                             "(Tanpa flag apa pun = jalankan semua, lihat --all.)")
    parser.add_argument("--reorder", action="store_true",
                        help="Generate laporan reorder standalone (cepat, tanpa analisa tahunan).")
    parser.add_argument("--cashflow", action="store_true",
                        help="Rencana cash-flow restock: modal beli yang dibutuhkan & kapan, per supplier.")
    parser.add_argument("--bundle", action="store_true",
                        help="Bundle & cross-sell: SKU yang sering dibeli bersama (market basket).")
    parser.add_argument("--deadstock", action="store_true",
                        help="Modal beku: kapital tertahan di stok lambat/mati + cara membebaskannya.")
    parser.add_argument("--momentum", action="store_true",
                        help="Momentum & ABC: SKU mana yang didorong (naik) vs dipangkas (turun).")
    parser.add_argument("--trend", action="store_true",
                        help="Tren & musiman: tren omzet/profit lintas tahun, pertumbuhan YoY, indeks musiman.")
    parser.add_argument("--ab-test", action="store_true",
                        help="Generate laporan A/B test (perubahan harga). Otomatis bikin template kalau belum ada.")
    parser.add_argument("--restock-check", action="store_true",
                        help="Analisa harga restock (beli vs jual per marketplace). Otomatis bikin template kalau belum ada.")
    parser.add_argument("--all", action="store_true",
                        help="Run SEMUANYA: sales all years + reorder + ab-test + restock-check "
                             "(ab-test & restock-check jalan kalau template-nya ada isinya).")
    parser.add_argument("--data-dir", type=Path, default=DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    args = parser.parse_args()

    try:
        if args.all:
            run_everything(args.data_dir, args.output_dir)
        elif args.restock_check:
            run_restock_check(args.data_dir, args.output_dir)
        elif args.ab_test:
            run_ab_test(args.data_dir, args.output_dir)
        elif args.cashflow:
            run_cashflow(args.data_dir, args.output_dir)
        elif args.bundle:
            run_bundle(args.data_dir, args.output_dir)
        elif args.deadstock:
            run_deadstock(args.data_dir, args.output_dir)
        elif args.momentum:
            run_momentum(args.data_dir, args.output_dir)
        elif args.trend:
            run_trend(args.data_dir, args.output_dir)
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
            # No flag → run the full suite (same as --all).
            run_everything(args.data_dir, args.output_dir)
        return 0
    except FileNotFoundError as e:
        print(f"❌ {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"❌ Error data: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
