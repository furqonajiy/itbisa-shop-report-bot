"""Build analytical tables from per-SKU aggregates."""
from __future__ import annotations

import numpy as np
import pandas as pd

from analysis import classify_supplier
from config import (
    HPP_VARIANCE_THRESHOLD, MARGIN_BORDERLINE_MAX, MARGIN_BORDERLINE_MIN,
    MARGIN_THRESHOLD_KANDIDAT, PRICE_SCENARIOS, QTY_PERCENTILE,
    SCORE_WEIGHT_MARGIN, SCORE_WEIGHT_VELOCITY, SUPPLIER_TOP_N_SINGLE_SOURCE,
    TOP_N_DIMINATI, TOP_N_PROFIT,
)


def _normalize(s: pd.Series) -> pd.Series:
    if s.max() == s.min():
        return pd.Series([50.0] * len(s), index=s.index)
    return (s - s.min()) / (s.max() - s.min()) * 100


def build_table_diminati(sku_agg: pd.DataFrame) -> pd.DataFrame:
    return (sku_agg.sort_values(["qty_terjual", "jumlah_transaksi"], ascending=False)
            .head(TOP_N_DIMINATI).reset_index(drop=True))


def build_table_profit(sku_agg: pd.DataFrame) -> pd.DataFrame:
    return sku_agg.sort_values("profit", ascending=False).head(TOP_N_PROFIT).reset_index(drop=True)


def build_table_rugi(sku_agg: pd.DataFrame) -> pd.DataFrame:
    df = sku_agg[sku_agg["profit"] < 0].copy().sort_values("profit")
    df["total_cost_per_buah"] = df["hpp_wa"] + df["biaya_admin_per_buah"].abs()
    df["selisih_harga"] = df["harga_jual_avg"] - df["total_cost_per_buah"]
    return df.reset_index(drop=True)


def build_table_borderline(sku_agg: pd.DataFrame) -> pd.DataFrame:
    return (sku_agg[
        (sku_agg["margin_pct"] >= MARGIN_BORDERLINE_MIN) &
        (sku_agg["margin_pct"] < MARGIN_BORDERLINE_MAX)
    ].copy().sort_values("qty_terjual", ascending=False).reset_index(drop=True))


def build_table_kandidat(sku_agg: pd.DataFrame) -> pd.DataFrame:
    qty_threshold = sku_agg["qty_terjual"].quantile(QTY_PERCENTILE)
    print(f"  Qty threshold (p{int(QTY_PERCENTILE*100)}): {qty_threshold:,.0f} pcs")

    kandidat = sku_agg[
        (sku_agg["qty_terjual"] >= qty_threshold) &
        (sku_agg["margin_pct"] >= MARGIN_THRESHOLD_KANDIDAT) &
        (sku_agg["sisa_stok"] > 0)
    ].copy()

    if len(kandidat) == 0:
        return kandidat

    kandidat["score_velocity"] = _normalize(kandidat["qty_terjual"])
    kandidat["score_margin"] = _normalize(kandidat["margin_pct"])
    kandidat["score_total"] = (
        kandidat["score_velocity"] * SCORE_WEIGHT_VELOCITY +
        kandidat["score_margin"] * SCORE_WEIGHT_MARGIN
    ).round(1)

    for pct in PRICE_SCENARIOS:
        kandidat[f"harga_+{int(pct*100)}pct"] = (kandidat["harga_jual_avg"] * (1 + pct)).round(0)

    base = PRICE_SCENARIOS[0]
    kandidat[f"proyeksi_profit_+{int(base*100)}pct"] = (
        kandidat["profit"] + kandidat["qty_terjual"] * kandidat["harga_jual_avg"] * base
    ).round(0)

    kandidat["saran"] = kandidat.apply(_saran_kandidat, axis=1)
    return kandidat.sort_values("score_total", ascending=False).reset_index(drop=True)


def _saran_kandidat(r) -> str:
    if (r["restock_di_tahun"] and pd.notna(r["qty_setelah_restock"])
            and r["qty_setelah_restock"] > r["qty_terjual"] * 0.5):
        return "🔥 Restock cepat habis — naik 15-20%"
    if r["margin_pct"] > 50:
        return "✅ Margin tinggi, naik 10-15%"
    if r["score_total"] > 50:
        return "⭐ Top performer, naik 10%"
    return "Naik 5-10% bertahap, monitor"


def build_table_platform(jual: pd.DataFrame) -> pd.DataFrame:
    plat = jual.groupby("akun_penjual").agg(
        transaksi=("Invoice", "nunique"),
        qty=("qty_jual", "sum"),
        omzet=("omzet", "sum"),
        hpp=("hpp_total", "sum"),
        biaya_admin=("biaya_admin", "sum"),
        profit=("profit", "sum"),
    ).reset_index()
    plat["margin_pct"] = plat["profit"] / plat["omzet"] * 100
    plat["admin_pct"] = plat["biaya_admin"].abs() / plat["omzet"] * 100
    return plat.sort_values("omzet", ascending=False).reset_index(drop=True)


def build_top_per_platform(jual: pd.DataFrame, platform: str, top_n: int) -> pd.DataFrame:
    sub = jual[jual["akun_penjual"] == platform]
    return (sub.groupby("SKU").agg(
        qty=("qty_jual", "sum"),
        omzet=("omzet", "sum"),
        profit=("profit", "sum"),
    ).reset_index().sort_values("profit", ascending=False)
        .head(top_n).reset_index(drop=True))


def build_supplier_analysis(stok: pd.DataFrame, sku_agg: pd.DataFrame) -> dict:
    """Build supplier analysis. Returns dict with 4 sub-tables for the Excel sheet."""
    s = stok.copy()
    s["supplier_type"] = s.apply(
        lambda r: classify_supplier(r.get("toko"), r.get("luar_negeri")), axis=1
    )
    s["hpp_per_buah_lot"] = s["total_hpp"] / s["qty_beli"]

    rows = []
    for sku, group in s.groupby("SKU"):
        row = {"SKU": sku}
        for stype in ["China", "Market", "Other"]:
            sub = group[group["supplier_type"] == stype]
            key = stype.lower()
            if len(sub) > 0:
                tq = sub["qty_beli"].sum()
                th = sub["total_hpp"].sum()
                row[f"qty_{key}"] = tq
                row[f"hpp_{key}"] = th / tq if tq > 0 else np.nan
                row[f"n_{key}"] = len(sub)
                row[f"hpp_min_{key}"] = sub["hpp_per_buah_lot"].min()
                row[f"hpp_max_{key}"] = sub["hpp_per_buah_lot"].max()
            else:
                row[f"qty_{key}"] = np.nan
                row[f"hpp_{key}"] = np.nan
                row[f"n_{key}"] = 0
                row[f"hpp_min_{key}"] = np.nan
                row[f"hpp_max_{key}"] = np.nan
        rows.append(row)

    df = pd.DataFrame(rows)
    df["selisih_market_vs_china"] = df["hpp_market"] - df["hpp_china"]
    df["cv_china"] = np.where(
        (df["hpp_china"] > 0) & (df["n_china"] > 1),
        (df["hpp_max_china"] - df["hpp_min_china"]) / df["hpp_china"],
        np.nan,
    )
    df["rekomendasi"] = df.apply(_saran_supplier, axis=1)

    df = df.merge(sku_agg[["SKU", "qty_terjual", "profit"]], on="SKU", how="left")
    df["qty_terjual"] = df["qty_terjual"].fillna(0)
    df["profit"] = df["profit"].fillna(0)

    has_china = df["hpp_china"].notna()
    has_market = df["hpp_market"].notna()
    has_both = has_china & has_market

    comparison = df[has_both].sort_values("qty_terjual", ascending=False).reset_index(drop=True)
    volatile = df[has_china & ~has_market & (df["cv_china"] > HPP_VARIANCE_THRESHOLD)].sort_values(
        "cv_china", ascending=False).reset_index(drop=True)
    china_only = df[has_china & ~has_market].sort_values(
        "qty_terjual", ascending=False).head(SUPPLIER_TOP_N_SINGLE_SOURCE).reset_index(drop=True)
    market_only = df[has_market & ~has_china].sort_values(
        "qty_terjual", ascending=False).head(SUPPLIER_TOP_N_SINGLE_SOURCE).reset_index(drop=True)

    return {
        "comparison": comparison,
        "volatile": volatile,
        "china_only": china_only,
        "market_only": market_only,
    }


def _saran_supplier(r) -> str:
    has_china = pd.notna(r.get("hpp_china"))
    has_market = pd.notna(r.get("hpp_market"))

    if has_china and has_market:
        if r["hpp_market"] < r["hpp_china"]:
            saving = (r["hpp_china"] - r["hpp_market"]) / r["hpp_china"] * 100
            return f"🔴 STOP China — Market {saving:.0f}% lebih murah"
        if r["hpp_china"] < r["hpp_market"] * 0.8:
            return "✅ China sangat efisien"
        return "China & Market mirip — restock fleksibel"
    if has_china:
        if pd.notna(r.get("cv_china")) and r["cv_china"] > HPP_VARIANCE_THRESHOLD:
            return "⚠ HPP China tidak konsisten — cek kurs/supplier"
        return "💡 Pertimbangkan test market buy"
    if has_market:
        return "💡 Pertimbangkan test import China"
    return "—"
