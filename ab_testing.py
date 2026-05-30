"""A/B testing analyzer: compare sales metrics before vs after price changes."""
from __future__ import annotations
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from config import (
    AB_MIN_DAYS_POST, AB_TESTS_SHEET, ALERT_TEXT_COLOR, COL_AB_CATATAN,
    COL_AB_NAMA, COL_AB_SKU, COL_AB_TANGGAL, FMT_DEC, FMT_NUM, FMT_PCT, FMT_RP,
    FONT_NAME, GREEN_FILL_COLOR, HEADER_BG_COLOR, HEADER_TEXT_COLOR,
    LIGHT_GRAY_COLOR, RED_FILL_COLOR, TITLE_COLOR, YELLOW_FILL_COLOR,
)

HEADER_FONT = Font(name=FONT_NAME, bold=True, color=HEADER_TEXT_COLOR, size=11)
HEADER_FILL = PatternFill("solid", start_color=HEADER_BG_COLOR)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14, color=TITLE_COLOR)
BIG_TITLE_FONT = Font(name=FONT_NAME, bold=True, size=18, color=TITLE_COLOR)
SUB_FONT = Font(name=FONT_NAME, italic=True, size=10, color="555555")
NORMAL_FONT = Font(name=FONT_NAME, size=10)
BOLD_FONT = Font(name=FONT_NAME, bold=True, size=10)
LIGHT_FILL = PatternFill("solid", start_color=LIGHT_GRAY_COLOR)
GREEN_FILL = PatternFill("solid", start_color=GREEN_FILL_COLOR)
YELLOW_FILL = PatternFill("solid", start_color=YELLOW_FILL_COLOR)
RED_FILL = PatternFill("solid", start_color=RED_FILL_COLOR)


def load_ab_tests(filepath: Path) -> pd.DataFrame:
    """Load A/B test config. Returns empty DataFrame if file missing."""
    if not filepath.exists():
        print(f"  ⚠ File {filepath.name} tidak ditemukan, A/B test skipped")
        return pd.DataFrame()

    df = pd.read_excel(filepath, sheet_name=AB_TESTS_SHEET)
    df = df.rename(columns={
        COL_AB_SKU: "sku",
        COL_AB_TANGGAL: "tanggal_perubahan",
        COL_AB_NAMA: "nama_test",
        COL_AB_CATATAN: "catatan",
    })
    df["tanggal_perubahan"] = pd.to_datetime(df["tanggal_perubahan"], errors="coerce")
    df = df[df["sku"].notna() & df["tanggal_perubahan"].notna()].copy()
    print(f"✓ Loaded {len(df):,} test config dari {filepath.name}")
    return df


def _compute_period_metrics(jual_sub: pd.DataFrame, hpp_wa: float) -> dict:
    """Compute metrics for a slice of jual transactions."""
    if len(jual_sub) == 0:
        return {
            "n_trans": 0, "qty": 0, "omzet": 0.0, "profit": 0.0,
            "avg_price": float("nan"), "markup_pct": float("nan"),
            "margin_pct": float("nan"), "days": 0,
            "qty_per_day": 0.0, "omzet_per_day": 0.0, "profit_per_day": 0.0,
            "first_date": None, "last_date": None,
        }

    qty = jual_sub["qty_jual"].sum()
    omzet = jual_sub["omzet"].sum()
    biaya_admin = (jual_sub["tambahan"] + jual_sub["kode_unik"]).sum()
    hpp_total = hpp_wa * qty if pd.notna(hpp_wa) else 0
    profit = omzet - hpp_total + biaya_admin

    first_date = jual_sub["tanggal_pesan"].min()
    last_date = jual_sub["tanggal_pesan"].max()
    days = max(1, (last_date - first_date).days + 1)

    avg_price = omzet / qty if qty > 0 else float("nan")
    markup_pct = ((avg_price - hpp_wa) / hpp_wa * 100) if pd.notna(hpp_wa) and hpp_wa > 0 else float("nan")
    margin_pct = (profit / omzet * 100) if omzet > 0 else float("nan")

    return {
        "n_trans": jual_sub["Invoice"].nunique(),
        "qty": qty,
        "omzet": omzet,
        "profit": profit,
        "avg_price": avg_price,
        "markup_pct": markup_pct,
        "margin_pct": margin_pct,
        "days": days,
        "qty_per_day": qty / days,
        "omzet_per_day": omzet / days,
        "profit_per_day": profit / days,
        "first_date": first_date,
        "last_date": last_date,
    }


def _pct_change(old: float, new: float) -> float:
    if pd.isna(old) or pd.isna(new) or old == 0:
        return float("nan")
    return (new - old) / abs(old) * 100


def _verdict(delta_qty: float, delta_profit: float, delta_price: float) -> str:
    if pd.isna(delta_profit) or pd.isna(delta_qty):
        return "⚪ Inconclusive (data sedikit)"
    if delta_profit > 5 and delta_qty > -10:
        return "✅ Effective — profit naik, qty stabil"
    if delta_profit > 5 and delta_qty <= -10:
        return "🟡 Mixed — profit naik tapi qty turun"
    if delta_profit < -5:
        return "🔴 Bad — profit turun"
    if abs(delta_profit) <= 5:
        return "⚪ No significant change"
    return "⚪ Unclear"


def analyze_ab_tests(ab_tests: pd.DataFrame, jual_full_clean: pd.DataFrame,
                     hpp_agg: pd.DataFrame, today: datetime) -> pd.DataFrame:
    """For each test config, compute pre vs post metrics."""
    if len(ab_tests) == 0:
        return pd.DataFrame()

    hpp_lookup = hpp_agg.set_index("SKU")["hpp_wa"].to_dict()
    results = []

    for _, t in ab_tests.iterrows():
        sku = t["sku"]
        change_date = t["tanggal_perubahan"]
        hpp = hpp_lookup.get(sku, float("nan"))

        sku_jual = jual_full_clean[jual_full_clean["SKU"] == sku].copy()

        if len(sku_jual) == 0:
            print(f"  ⚠ {sku}: tidak ada transaksi, skip")
            continue

        pre = sku_jual[sku_jual["tanggal_pesan"] < change_date]
        post = sku_jual[sku_jual["tanggal_pesan"] >= change_date]

        pre_m = _compute_period_metrics(pre, hpp)
        post_m = _compute_period_metrics(post, hpp)

        warning = ""
        if post_m["days"] < AB_MIN_DAYS_POST:
            warning = f"⚠ Post period baru {post_m['days']} hari — data belum cukup"

        delta_qty = _pct_change(pre_m["qty_per_day"], post_m["qty_per_day"])
        delta_omzet = _pct_change(pre_m["omzet_per_day"], post_m["omzet_per_day"])
        delta_profit = _pct_change(pre_m["profit_per_day"], post_m["profit_per_day"])
        delta_price = _pct_change(pre_m["avg_price"], post_m["avg_price"])

        results.append({
            "sku": sku,
            "nama_test": t.get("nama_test", "") or "",
            "tanggal_perubahan": change_date,
            "hpp_wa": hpp,
            "pre_first": pre_m["first_date"],
            "pre_last": pre_m["last_date"],
            "pre_days": pre_m["days"] if pre_m["n_trans"] > 0 else 0,
            "pre_n_trans": pre_m["n_trans"],
            "pre_qty": pre_m["qty"],
            "pre_qty_per_day": pre_m["qty_per_day"],
            "pre_omzet_per_day": pre_m["omzet_per_day"],
            "pre_profit_per_day": pre_m["profit_per_day"],
            "pre_avg_price": pre_m["avg_price"],
            "pre_markup_pct": pre_m["markup_pct"],
            "pre_margin_pct": pre_m["margin_pct"],
            "post_first": post_m["first_date"],
            "post_last": post_m["last_date"],
            "post_days": post_m["days"] if post_m["n_trans"] > 0 else 0,
            "post_n_trans": post_m["n_trans"],
            "post_qty": post_m["qty"],
            "post_qty_per_day": post_m["qty_per_day"],
            "post_omzet_per_day": post_m["omzet_per_day"],
            "post_profit_per_day": post_m["profit_per_day"],
            "post_avg_price": post_m["avg_price"],
            "post_markup_pct": post_m["markup_pct"],
            "post_margin_pct": post_m["margin_pct"],
            "delta_qty_pct": delta_qty,
            "delta_omzet_pct": delta_omzet,
            "delta_profit_pct": delta_profit,
            "delta_price_pct": delta_price,
            "verdict": _verdict(delta_qty, delta_profit, delta_price),
            "warning": warning,
            "catatan": t.get("catatan", "") or "",
        })

    return pd.DataFrame(results)


def write_ab_test_report(output_path: Path, results: pd.DataFrame,
                          today: datetime) -> None:
    """Write A/B test analysis to a standalone Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "00_Summary"

    ws["A1"] = "LAPORAN ANALISA A/B TEST — ITBISA SHOP"
    ws["A1"].font = BIG_TITLE_FONT
    ws.merge_cells("A1:F1")
    ws["A2"] = (f"Generated: {today.strftime('%d %B %Y %H:%M')}  |  "
                f"Window: Full data pre vs post change date")
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:F2")

    if len(results) == 0:
        ws["A4"] = "Tidak ada test config. Isi data/ab_tests.xlsx untuk track perubahan harga."
        ws["A4"].font = BOLD_FONT
        ws.column_dimensions["A"].width = 70
        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        print(f"✓ Menulis laporan ke {output_path} (kosong)")
        return

    ws["A4"] = "RINGKASAN"
    ws["A4"].font = TITLE_FONT
    summary = [
        ("Total Test Tracked", len(results)),
        ("Effective (profit naik)", (results["verdict"].str.startswith("✅")).sum()),
        ("Mixed", (results["verdict"].str.startswith("🟡")).sum()),
        ("Bad (profit turun)", (results["verdict"].str.startswith("🔴")).sum()),
        ("Inconclusive / No change", (results["verdict"].str.startswith("⚪")).sum()),
    ]
    for i, (lbl, val) in enumerate(summary, start=5):
        ws.cell(row=i, column=1, value=lbl).font = NORMAL_FONT
        c = ws.cell(row=i, column=2, value=val)
        c.font = BOLD_FONT
        c.number_format = FMT_NUM

    ws["A11"] = "DAFTAR TEST"
    ws["A11"].font = TITLE_FONT
    for i, (_, r) in enumerate(results.iterrows(), start=12):
        ws.cell(row=i, column=1, value=r["sku"]).font = NORMAL_FONT
        ws.cell(row=i, column=2, value=r["nama_test"]).font = NORMAL_FONT
        c = ws.cell(row=i, column=3, value=r["verdict"])
        c.font = BOLD_FONT
        if r["warning"]:
            wc = ws.cell(row=i, column=4, value=r["warning"])
            wc.font = Font(name=FONT_NAME, italic=True, size=10, color=ALERT_TEXT_COLOR)

    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 38
    ws.column_dimensions["D"].width = 40

    _write_details_sheet(wb.create_sheet("01_Test_Results"), results)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"✓ Menulis laporan ke {output_path}")


def _write_details_sheet(ws, results: pd.DataFrame) -> None:
    ws["A1"] = "DETAIL A/B TEST — PRE vs POST"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:T1")
    ws["A2"] = ("Bandingan periode SEBELUM dan SESUDAH perubahan harga. "
                "Daily rates dipakai supaya fair karena window beda panjang.")
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:T2")

    headers = [
        "SKU", "Nama Test", "Tgl Perubahan", "HPP/Buah",
        "Pre Days", "Pre Qty/Day", "Pre Omzet/Day", "Pre Profit/Day",
        "Pre Avg Price", "Pre Markup %",
        "Post Days", "Post Qty/Day", "Post Omzet/Day", "Post Profit/Day",
        "Post Avg Price", "Post Markup %",
        "Δ Qty/Day", "Δ Omzet/Day", "Δ Profit/Day", "Δ Price",
        "Verdict", "Warning",
    ]
    widths = [38, 25, 13, 11,
              9, 12, 14, 14, 13, 11,
              9, 12, 14, 14, 13, 11,
              11, 11, 11, 11,
              38, 30]

    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.row_dimensions[4].height = 30

    pct_cols = {10, 16, 17, 18, 19, 20}
    rp_cols = {4, 7, 8, 9, 13, 14, 15}
    num_cols = {5, 6, 11, 12}

    for r_idx, (_, row) in enumerate(results.iterrows()):
        delta_profit = row["delta_profit_pct"]
        verdict = row["verdict"]
        for c_idx, h in enumerate(headers, start=1):
            field_map = {
                "SKU": "sku", "Nama Test": "nama_test", "Tgl Perubahan": "tanggal_perubahan",
                "HPP/Buah": "hpp_wa",
                "Pre Days": "pre_days", "Pre Qty/Day": "pre_qty_per_day",
                "Pre Omzet/Day": "pre_omzet_per_day", "Pre Profit/Day": "pre_profit_per_day",
                "Pre Avg Price": "pre_avg_price", "Pre Markup %": "pre_markup_pct",
                "Post Days": "post_days", "Post Qty/Day": "post_qty_per_day",
                "Post Omzet/Day": "post_omzet_per_day", "Post Profit/Day": "post_profit_per_day",
                "Post Avg Price": "post_avg_price", "Post Markup %": "post_markup_pct",
                "Δ Qty/Day": "delta_qty_pct", "Δ Omzet/Day": "delta_omzet_pct",
                "Δ Profit/Day": "delta_profit_pct", "Δ Price": "delta_price_pct",
                "Verdict": "verdict", "Warning": "warning",
            }
            val = row[field_map[h]]
            if pd.isna(val):
                val = None
            elif h in ("Pre Markup %", "Post Markup %", "Δ Qty/Day",
                       "Δ Omzet/Day", "Δ Profit/Day", "Δ Price"):
                val = val / 100 if val is not None else None
            cell = ws.cell(row=5 + r_idx, column=c_idx, value=val)
            cell.font = NORMAL_FONT
            if c_idx in pct_cols:
                cell.number_format = FMT_PCT
            elif c_idx in rp_cols:
                cell.number_format = FMT_RP
            elif c_idx in num_cols:
                cell.number_format = FMT_DEC
            if r_idx % 2 == 1:
                cell.fill = LIGHT_FILL

        verdict_cell = ws.cell(row=5 + r_idx, column=21)
        if verdict.startswith("✅"):
            verdict_cell.fill = GREEN_FILL
        elif verdict.startswith("🔴"):
            verdict_cell.fill = RED_FILL
        elif verdict.startswith("🟡"):
            verdict_cell.fill = YELLOW_FILL

    ws.freeze_panes = "B5"


def create_template(filepath: Path, example_sku: str = "ITBISA-IC-NE555P-DIP8") -> None:
    """Create the ab_tests.xlsx template with sample data."""
    wb = Workbook()
    ws = wb.active
    ws.title = AB_TESTS_SHEET

    headers = [COL_AB_SKU, COL_AB_TANGGAL, COL_AB_NAMA, COL_AB_CATATAN]
    widths = [38, 18, 30, 50]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.row_dimensions[1].height = 25

    example = [
        example_sku,
        datetime(2026, 5, 17),
        "NE555P Price Bump May 2026",
        "Harga naik dari 599 ke 699 (single), 50pcs ke 689, 1000pcs ke 679",
    ]
    for i, v in enumerate(example, start=1):
        c = ws.cell(row=2, column=i, value=v)
        c.font = NORMAL_FONT
        if i == 2:
            c.number_format = "yyyy-mm-dd"

    ws.cell(row=4, column=1, value="Format kolom:").font = BOLD_FONT
    notes = [
        "SKU: persis sama dengan SKU di data Jual (case-sensitive)",
        "Tanggal Perubahan: tanggal harga berubah (YYYY-MM-DD)",
        "Nama Test: nama bebas untuk identifikasi (opsional)",
        "Catatan: detail perubahan harga, alasan, dll. (opsional)",
        "",
        "Tambah baris baru untuk setiap perubahan harga yang ingin dilacak.",
    ]
    for i, n in enumerate(notes, start=5):
        c = ws.cell(row=i, column=1, value=n)
        c.font = SUB_FONT

    filepath.parent.mkdir(parents=True, exist_ok=True)
    wb.save(filepath)
    print(f"✓ Template dibuat: {filepath}")
