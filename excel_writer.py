"""Excel report writer with styling helpers."""
from __future__ import annotations
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from config import (
    ALERT_TEXT_COLOR, FONT_NAME, FMT_DEC, FMT_NUM, FMT_PCT, FMT_RP,
    GREEN_FILL_COLOR, HEADER_BG_COLOR, HEADER_TEXT_COLOR, LIGHT_GRAY_COLOR,
    MARGIN_THRESHOLD_KANDIDAT, PRICE_SCENARIOS, RED_FILL_COLOR,
    TARGET_MARGIN_KOREKSI, TITLE_COLOR, TOP_N_PER_PLATFORM, YELLOW_FILL_COLOR,
)
from tables import build_top_per_platform

# Reusable styles
HEADER_FONT = Font(name=FONT_NAME, bold=True, color=HEADER_TEXT_COLOR, size=11)
HEADER_FILL = PatternFill("solid", start_color=HEADER_BG_COLOR)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14, color=TITLE_COLOR)
BIG_TITLE_FONT = Font(name=FONT_NAME, bold=True, size=18, color=TITLE_COLOR)
SUB_FONT = Font(name=FONT_NAME, italic=True, size=10, color="555555")
NORMAL_FONT = Font(name=FONT_NAME, size=10)
BOLD_FONT = Font(name=FONT_NAME, bold=True, size=10)
ALERT_FONT = Font(name=FONT_NAME, italic=True, size=10, color=ALERT_TEXT_COLOR)
RED_FILL = PatternFill("solid", start_color=RED_FILL_COLOR)
GREEN_FILL = PatternFill("solid", start_color=GREEN_FILL_COLOR)
YELLOW_FILL = PatternFill("solid", start_color=YELLOW_FILL_COLOR)
LIGHT_FILL = PatternFill("solid", start_color=LIGHT_GRAY_COLOR)
THIN_BORDER = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)


def write_headers(ws: Worksheet, row: int, headers: list[str],
                  widths: list[float] | None = None) -> None:
    """Write a styled header row with optional column widths."""
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER
        if widths:
            ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.row_dimensions[row].height = 30


def write_data_rows(ws: Worksheet, start_row: int, df: pd.DataFrame,
                    formats: list[str | None] | None = None,
                    negative_highlight_col: int | None = None) -> None:
    """Write data rows; optionally highlight whole row red if a column < 0.

    `negative_highlight_col` is 1-indexed (Excel column number).
    """
    for r_idx, (_, row) in enumerate(df.iterrows()):
        for c_idx, val in enumerate(row, start=1):
            if pd.isna(val):
                val = None
            cell = ws.cell(row=start_row + r_idx, column=c_idx, value=val)
            cell.font = NORMAL_FONT
            cell.border = THIN_BORDER
            if formats and c_idx - 1 < len(formats) and formats[c_idx - 1]:
                cell.number_format = formats[c_idx - 1]
            if r_idx % 2 == 1:
                cell.fill = LIGHT_FILL
        if negative_highlight_col is not None:
            target = ws.cell(row=start_row + r_idx, column=negative_highlight_col)
            if isinstance(target.value, (int, float)) and target.value < 0:
                for c in range(1, len(row) + 1):
                    ws.cell(row=start_row + r_idx, column=c).fill = RED_FILL


def _build_findings(sku_agg: pd.DataFrame, tables: dict,
                    sku_no_hpp: list[str]) -> list[str]:
    """Generate data-driven findings narrative as a list of lines."""
    lines: list[str] = []

    rugi = tables["rugi"]
    if len(rugi) > 0:
        lines.append("🔴 BARANG RUGI YANG HARUS SEGERA DINAIKKAN HARGANYA:")
        for _, r in rugi.head(5).iterrows():
            cost = r["hpp_wa"] + abs(r["biaya_admin_per_buah"])
            rekom = cost / (1 - TARGET_MARGIN_KOREKSI)
            lines.append(
                f"   • {r['SKU']} — terjual {r['qty_terjual']:,.0f} pcs, "
                f"RUGI Rp {r['profit']:,.0f}. Cost Rp {cost:,.0f}, "
                f"dijual Rp {r['harga_jual_avg']:,.0f}. "
                f"Saran: Rp {rekom:,.0f} (margin {TARGET_MARGIN_KOREKSI*100:.0f}%)."
            )
        lines.append("")

    kandidat = tables["kandidat"]
    if len(kandidat) > 0:
        lines.append("🟢 TOP KANDIDAT NAIK HARGA (laris + margin sehat + stok ada):")
        for _, r in kandidat.head(5).iterrows():
            note = ""
            if r["restock_di_tahun"] and pd.notna(r["qty_setelah_restock"]) and r["qty_terjual"] > 0:
                pct = r["qty_setelah_restock"] / r["qty_terjual"] * 100
                note = f", restock tahun ini langsung {r['qty_setelah_restock']:,.0f} pcs terjual ({pct:.0f}%)"
            lines.append(
                f"   • {r['SKU']} — {r['qty_terjual']:,.0f} pcs, "
                f"margin {r['margin_pct']:.1f}%, stok {r['sisa_stok']:,.0f}{note}."
            )
        lines.append("")

    top_qty = sku_agg.nlargest(3, "qty_terjual")
    lines.append("⭐ BARANG PALING DIMINATI:")
    for i, (_, r) in enumerate(top_qty.iterrows(), 1):
        lines.append(
            f"   {i}. {r['SKU']} — {r['qty_terjual']:,.0f} pcs, "
            f"{r['jumlah_transaksi']:,} transaksi "
            f"(rata-rata {r['avg_qty_per_order']:.0f} pcs/order)"
        )
    lines.append("")

    top_profit = sku_agg.nlargest(5, "profit")
    lines.append("💰 PENYUMBANG PROFIT TERBESAR:")
    for i, (_, r) in enumerate(top_profit.iterrows(), 1):
        lines.append(f"   {i}. {r['SKU']} — Rp {r['profit']:,.0f}")
    lines.append("")

    plat = tables["platform"]
    total_omzet_plat = plat["omzet"].sum()
    lines.append("🏪 PER PLATFORM:")
    for _, r in plat.iterrows():
        share = r["omzet"] / total_omzet_plat * 100 if total_omzet_plat else 0
        lines.append(
            f"   • {r['akun_penjual']}: {share:.0f}% omzet, "
            f"margin {r['margin_pct']:.1f}%, biaya admin {r['admin_pct']:.1f}%"
        )
    lines.append("")

    if sku_no_hpp:
        lines.append(
            f"CATATAN: {len(sku_no_hpp)} SKU dijual tanpa HPP (di-exclude): "
            + ", ".join(sku_no_hpp)
        )
    return lines


def _write_summary(ws: Worksheet, year: int, jual: pd.DataFrame,
                   sku_agg: pd.DataFrame, tables: dict,
                   sku_no_hpp: list[str]) -> None:
    """Write executive summary sheet with totals and findings."""
    ws["A1"] = f"LAPORAN ANALISA PENJUALAN ITBISA SHOP {year}"
    ws["A1"].font = BIG_TITLE_FONT
    ws.merge_cells("A1:D1")
    ws["A2"] = (f"Generated: {datetime.now().strftime('%d %B %Y %H:%M')}  |  "
                f"Periode: {year}  |  Metode HPP: Weighted Average")
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:D2")

    total_omzet = jual["omzet"].sum()
    total_hpp = jual["hpp_total"].sum()
    total_admin = jual["biaya_admin"].sum()
    total_profit = jual["profit"].sum()
    total_qty = jual["qty_jual"].sum()
    total_trans = jual["Invoice"].nunique()
    margin = total_profit / total_omzet if total_omzet else 0

    ws["A4"] = "RINGKASAN PERFORMA"
    ws["A4"].font = TITLE_FONT
    summary = [
        ("Total Transaksi", total_trans, FMT_NUM),
        ("Total Qty Terjual (pcs)", total_qty, FMT_NUM),
        ("Unique SKU Terjual", len(sku_agg), FMT_NUM),
        ("", "", ""),
        ("Total Omzet", total_omzet, FMT_RP),
        ("Total HPP", total_hpp, FMT_RP),
        ("Total Biaya Admin Platform", total_admin, FMT_RP),
        ("TOTAL PROFIT", total_profit, FMT_RP),
        ("Margin Keseluruhan", margin, FMT_PCT),
        ("", "", ""),
        ("Barang RUGI Total", len(tables["rugi"]), FMT_NUM),
        ("Margin Borderline 0-5%", len(tables["borderline"]), FMT_NUM),
        ("Kandidat Naik Harga", len(tables["kandidat"]), FMT_NUM),
    ]
    for i, (lbl, val, fmt) in enumerate(summary, start=6):
        is_bold = "TOTAL" in str(lbl) or "Margin" in str(lbl)
        c1 = ws.cell(row=i, column=1, value=lbl)
        c1.font = BOLD_FONT if is_bold else NORMAL_FONT
        c2 = ws.cell(row=i, column=2, value=val)
        c2.font = BOLD_FONT if is_bold else NORMAL_FONT
        if fmt:
            c2.number_format = fmt
        if "TOTAL PROFIT" in str(lbl):
            c2.fill = GREEN_FILL

    ws["A21"] = "TEMUAN UTAMA & REKOMENDASI"
    ws["A21"].font = TITLE_FONT
    findings = _build_findings(sku_agg, tables, sku_no_hpp)
    for i, line in enumerate(findings, start=22):
        c = ws.cell(row=i, column=1, value=line)
        c.font = NORMAL_FONT
        if line and any(line.startswith(em) for em in ["🔴", "🟢", "⭐", "💰", "🏪"]):
            c.font = BOLD_FONT
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=10)

    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 22


def _write_diminati(ws: Worksheet, df: pd.DataFrame) -> None:
    ws["A1"] = "TABEL 1: BARANG PALING DIMINATI (sort by Qty Terjual)"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:I1")
    ws["A2"] = "Avg Qty/Order TINGGI = banyak reseller borongan. RENDAH = banyak transaksi retail."
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:I2")
    write_headers(ws, 4,
        ["SKU", "Qty Terjual", "Jumlah Transaksi", "Avg Qty/Order",
         "Harga Jual Avg", "Omzet", "HPP/buah", "Profit", "Margin %"],
        widths=[38, 13, 13, 13, 15, 18, 13, 18, 12],
    )
    df = df.copy()
    df["margin_frac"] = df["margin_pct"] / 100
    out = df[["SKU", "qty_terjual", "jumlah_transaksi", "avg_qty_per_order",
              "harga_jual_avg", "omzet", "hpp_wa", "profit", "margin_frac"]]
    write_data_rows(ws, 5, out,
        formats=[None, FMT_NUM, FMT_NUM, FMT_DEC, FMT_RP, FMT_RP, FMT_RP, FMT_RP, FMT_PCT],
        negative_highlight_col=8)
    ws.freeze_panes = "B5"


def _write_profit(ws: Worksheet, df: pd.DataFrame) -> None:
    ws["A1"] = "TABEL 2: BARANG PENYUMBANG PROFIT TERBESAR"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    ws["A2"] = "Sort by total Profit (Rupiah). Profit/buah penting untuk lihat efisiensi per unit."
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:H2")
    write_headers(ws, 4,
        ["SKU", "Total Profit", "Profit/Buah", "Qty Terjual",
         "Margin %", "Omzet", "Harga Jual Avg", "HPP/Buah"],
        widths=[38, 18, 13, 13, 12, 18, 15, 13])
    df = df.copy()
    df["margin_frac"] = df["margin_pct"] / 100
    out = df[["SKU", "profit", "profit_per_buah", "qty_terjual",
              "margin_frac", "omzet", "harga_jual_avg", "hpp_wa"]]
    write_data_rows(ws, 5, out,
        formats=[None, FMT_RP, FMT_RP, FMT_NUM, FMT_PCT, FMT_RP, FMT_RP, FMT_RP],
        negative_highlight_col=2)
    ws.freeze_panes = "B5"


def _write_rugi(ws: Worksheet, df: pd.DataFrame) -> None:
    ws["A1"] = "TABEL 3: BARANG RUGI / SALAH PASANG HARGA"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:I1")
    ws["A2"] = "Total Cost = HPP + |Admin/buah|. Selisih = Harga - Total Cost. Selisih NEGATIF = jual di bawah modal."
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:I2")

    if len(df) == 0:
        ws["A4"] = "(tidak ada barang rugi — kerja bagus!)"
        ws["A4"].font = BOLD_FONT
        return

    write_headers(ws, 4,
        ["SKU", "Qty Terjual", "HPP/buah", "Admin/buah", "Total Cost/buah",
         "Harga Jual Avg", "Selisih", "Profit/buah", "Total Profit"],
        widths=[38, 13, 13, 13, 15, 15, 13, 13, 18])
    out = df[["SKU", "qty_terjual", "hpp_wa", "biaya_admin_per_buah",
              "total_cost_per_buah", "harga_jual_avg", "selisih_harga",
              "profit_per_buah", "profit"]]
    write_data_rows(ws, 5, out,
        formats=[None, FMT_NUM, FMT_RP, FMT_RP, FMT_RP, FMT_RP, FMT_RP, FMT_RP, FMT_RP],
        negative_highlight_col=9)

    row_rec = 5 + len(df) + 2
    ws.cell(row=row_rec, column=1,
            value=f"REKOMENDASI HARGA KOREKSI (target margin {TARGET_MARGIN_KOREKSI*100:.0f}%):").font = BOLD_FONT
    ws.merge_cells(start_row=row_rec, start_column=1, end_row=row_rec, end_column=5)
    for i, (_, r) in enumerate(df.iterrows(), start=row_rec + 1):
        cost = r["hpp_wa"] + abs(r["biaya_admin_per_buah"])
        rekom = cost / (1 - TARGET_MARGIN_KOREKSI)
        ws.cell(row=i, column=1, value=r["SKU"]).font = NORMAL_FONT
        ws.cell(row=i, column=2, value=f"Sekarang: Rp {r['harga_jual_avg']:,.0f}").font = NORMAL_FONT
        c = ws.cell(row=i, column=4, value=f"Rekomendasi: Rp {rekom:,.0f}")
        c.font = BOLD_FONT
        c.fill = GREEN_FILL


def _write_borderline(ws: Worksheet, df: pd.DataFrame) -> None:
    ws["A1"] = "TABEL 4: MARGIN BORDERLINE 0-5% (RAWAN RUGI)"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    ws["A2"] = "Margin tipis: rentan rugi kalau biaya admin naik atau ada diskon. Naikkan ~10-15%."
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:H2")

    if len(df) == 0:
        ws["A4"] = "(tidak ada SKU dalam kategori ini)"
        ws["A4"].font = BOLD_FONT
        return

    write_headers(ws, 4,
        ["SKU", "Qty Terjual", "HPP/buah", "Admin/buah",
         "Harga Jual Avg", "Profit/buah", "Margin %", "Profit Total"],
        widths=[38, 13, 13, 13, 15, 13, 12, 18])
    df = df.copy()
    df["margin_frac"] = df["margin_pct"] / 100
    out = df[["SKU", "qty_terjual", "hpp_wa", "biaya_admin_per_buah",
              "harga_jual_avg", "profit_per_buah", "margin_frac", "profit"]]
    write_data_rows(ws, 5, out,
        formats=[None, FMT_NUM, FMT_RP, FMT_RP, FMT_RP, FMT_RP, FMT_PCT, FMT_RP])


def _write_kandidat(ws: Worksheet, df: pd.DataFrame) -> None:
    ws["A1"] = "TABEL 5: KANDIDAT NAIK HARGA"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:M1")
    ws["A2"] = (f"Kriteria: top 20% qty terjual + margin ≥{MARGIN_THRESHOLD_KANDIDAT:.0f}% "
                "+ stok ada. Score = 60% velocity + 40% margin headroom.")
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:M2")
    ws["A3"] = '⚠ "Qty Setelah Restock" tinggi = restock baru cepat habis = sinyal harga underpriced.'
    ws["A3"].font = ALERT_FONT
    ws.merge_cells("A3:M3")

    if len(df) == 0:
        ws["A5"] = "(belum ada kandidat memenuhi semua kriteria)"
        ws["A5"].font = BOLD_FONT
        return

    pcts = PRICE_SCENARIOS
    headers = ["SKU", "Score", "Qty Terjual", "Margin %", "Harga Sekarang",
               "Sisa Stok", "Restock Tahun?", "Qty Setelah Restock"]
    headers += [f"Harga +{int(p*100)}%" for p in pcts]
    headers += [f"Proyeksi Profit +{int(pcts[0]*100)}%", "Saran"]
    widths = [35, 8, 12, 10, 14, 11, 13, 14] + [13] * len(pcts) + [18, 35]
    write_headers(ws, 5, headers, widths=widths)

    df = df.copy()
    df["margin_frac"] = df["margin_pct"] / 100
    df["restock_str"] = df["restock_di_tahun"].map({True: "Ya", False: "Tidak"})

    out_cols = ["SKU", "score_total", "qty_terjual", "margin_frac", "harga_jual_avg",
                "sisa_stok", "restock_str", "qty_setelah_restock"]
    out_cols += [f"harga_+{int(p*100)}pct" for p in pcts]
    out_cols += [f"proyeksi_profit_+{int(pcts[0]*100)}pct", "saran"]
    out = df[out_cols]

    formats = [None, FMT_DEC, FMT_NUM, FMT_PCT, FMT_RP, FMT_NUM, None, FMT_NUM]
    formats += [FMT_RP] * len(pcts)
    formats += [FMT_RP, None]
    write_data_rows(ws, 6, out, formats=formats)

    # Highlight top 5 with yellow
    end_col = len(headers)
    n_top = min(5, len(df))
    for r in range(6, 6 + n_top):
        for c in range(1, end_col + 1):
            ws.cell(row=r, column=c).fill = YELLOW_FILL


def _write_platform(ws: Worksheet, plat_df: pd.DataFrame, jual: pd.DataFrame) -> None:
    ws["A1"] = "TABEL 6: BREAKDOWN PER PLATFORM"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:I1")
    ws["A2"] = "Lihat margin dan biaya admin per platform. Platform dengan admin% tinggi = perlu evaluasi."
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:I2")
    write_headers(ws, 4,
        ["Platform", "Transaksi", "Qty", "Omzet", "HPP",
         "Biaya Admin", "Profit", "Margin %", "Admin %"],
        widths=[15, 12, 13, 18, 18, 18, 18, 12, 12])
    plat_df = plat_df.copy()
    plat_df["margin_frac"] = plat_df["margin_pct"] / 100
    plat_df["admin_frac"] = plat_df["admin_pct"] / 100
    out = plat_df[["akun_penjual", "transaksi", "qty", "omzet", "hpp",
                   "biaya_admin", "profit", "margin_frac", "admin_frac"]]
    write_data_rows(ws, 5, out,
        formats=[None, FMT_NUM, FMT_NUM, FMT_RP, FMT_RP, FMT_RP, FMT_RP, FMT_PCT, FMT_PCT])

    section_row = 5 + len(plat_df) + 2
    ws.cell(row=section_row, column=1, value="TOP SKU BY PROFIT — PER PLATFORM").font = TITLE_FONT

    row_cur = section_row + 2
    for platform in plat_df["akun_penjual"].tolist():
        top = build_top_per_platform(jual, platform, TOP_N_PER_PLATFORM)
        if len(top) == 0:
            continue
        ws.cell(row=row_cur, column=1, value=platform).font = BOLD_FONT
        write_headers(ws, row_cur + 1, ["SKU", "Qty", "Omzet", "Profit"],
                      widths=[38, 12, 18, 18])
        write_data_rows(ws, row_cur + 2, top,
            formats=[None, FMT_NUM, FMT_RP, FMT_RP],
            negative_highlight_col=4)
        row_cur += 3 + len(top)


def _write_full_data(ws: Worksheet, sku_agg: pd.DataFrame) -> None:
    ws["A1"] = "DATA LENGKAP PER SKU"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:Q1")
    cols = ["SKU", "qty_terjual", "jumlah_transaksi", "avg_qty_per_order", "omzet",
            "hpp_wa", "hpp_cost", "biaya_admin", "biaya_admin_per_buah", "profit",
            "profit_per_buah", "harga_jual_avg", "margin_pct", "total_qty_beli",
            "sisa_stok", "jumlah_pembelian", "restock_di_tahun"]
    headers = ["SKU", "Qty Terjual", "Trans", "Avg Qty/Order", "Omzet", "HPP/Buah",
               "HPP Total", "Biaya Admin", "Admin/Buah", "Profit", "Profit/Buah",
               "Harga Jual Avg", "Margin %", "Total Dibeli", "Sisa Stok",
               "Lot Beli", "Restock Tahun"]
    widths = [38, 12, 8, 13, 16, 12, 16, 16, 12, 16, 12, 14, 11, 12, 12, 9, 11]
    write_headers(ws, 3, headers, widths=widths)

    df_full = sku_agg[cols].sort_values("profit", ascending=False).copy()
    df_full["margin_pct"] = df_full["margin_pct"] / 100
    df_full["restock_di_tahun"] = df_full["restock_di_tahun"].map({True: "Ya", False: "Tidak"})
    write_data_rows(ws, 4, df_full,
        formats=[None, FMT_NUM, FMT_NUM, FMT_DEC, FMT_RP, FMT_RP, FMT_RP, FMT_RP,
                 FMT_RP, FMT_RP, FMT_RP, FMT_RP, FMT_PCT, FMT_NUM, FMT_NUM, FMT_NUM, None],
        negative_highlight_col=10)
    ws.freeze_panes = "B4"


def write_report(output_path: Path, year: int, jual: pd.DataFrame,
                 sku_agg: pd.DataFrame, tables: dict, sku_no_hpp: list[str]) -> None:
    """Orchestrate writing of all 8 sheets to a single Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "00_Summary"
    _write_summary(ws, year, jual, sku_agg, tables, sku_no_hpp)

    _write_diminati(wb.create_sheet("01_Paling_Diminati"), tables["diminati"])
    _write_profit(wb.create_sheet("02_Profit_Tertinggi"), tables["profit"])
    _write_rugi(wb.create_sheet("03_Barang_Rugi"), tables["rugi"])
    _write_borderline(wb.create_sheet("04_Margin_Borderline"), tables["borderline"])
    _write_kandidat(wb.create_sheet("05_Kandidat_Naik_Harga"), tables["kandidat"])
    _write_platform(wb.create_sheet("06_Per_Platform"), tables["platform"], jual)
    _write_full_data(wb.create_sheet("07_Data_Lengkap_per_SKU"), sku_agg)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"✓ Menulis laporan ke {output_path}")
