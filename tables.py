"""Build analytical tables from per-SKU aggregate data."""
from __future__ import annotations

import pandas as pd

from config import (
    QTY_PERCENTILE, MARGIN_THRESHOLD_KANDIDAT,
    MARGIN_BORDERLINE_MIN, MARGIN_BORDERLINE_MAX,
    SCORE_WEIGHT_VELOCITY, SCORE_WEIGHT_MARGIN,
    PRICE_SCENARIOS, TOP_N_DIMINATI, TOP_N_PROFIT,
)


def _normalize(series: pd.Series) -> pd.Series:
    """Min-max normalize to 0-100. Returns 50 for constant series."""
    if series.max() == series.min():
        return pd.Series([50.0] * len(series), index=series.index)
    return (series - series.min()) / (series.max() - series.min()) * 100


def build_table_diminati(sku_agg: pd.DataFrame) -> pd.DataFrame:
    """Top N SKUs by qty terjual (primary), jumlah transaksi (secondary)."""
    return (sku_agg.sort_values(["qty_terjual", "jumlah_transaksi"], ascending=False)
            .head(TOP_N_DIMINATI).reset_index(drop=True))


def build_table_profit(sku_agg: pd.DataFrame) -> pd.DataFrame:
    """Top N SKUs by total profit."""
    return (sku_agg.sort_values("profit", ascending=False)
            .head(TOP_N_PROFIT).reset_index(drop=True))


def build_table_rugi(sku_agg: pd.DataFrame) -> pd.DataFrame:
    """SKUs with negative total profit; adds total_cost and selisih columns."""
    df = sku_agg[sku_agg["profit"] < 0].copy().sort_values("profit")
    df["total_cost_per_buah"] = df["hpp_wa"] + df["biaya_admin_per_buah"].abs()
    df["selisih_harga"] = df["harga_jual_avg"] - df["total_cost_per_buah"]
    return df.reset_index(drop=True)


def build_table_borderline(sku_agg: pd.DataFrame) -> pd.DataFrame:
    """SKUs with margin between MARGIN_BORDERLINE_MIN and MARGIN_BORDERLINE_MAX."""
    df = sku_agg[
        (sku_agg["margin_pct"] >= MARGIN_BORDERLINE_MIN) &
        (sku_agg["margin_pct"] < MARGIN_BORDERLINE_MAX)
    ].copy().sort_values("qty_terjual", ascending=False)
    return df.reset_index(drop=True)


def build_table_kandidat(sku_agg: pd.DataFrame) -> pd.DataFrame:
    """Scored candidates for price increase. Returns empty df if no qualifiers."""
    qty_threshold = sku_agg["qty_terjual"].quantile(QTY_PERCENTILE)
    print(f"  Qty threshold (percentile {QTY_PERCENTILE*100:.0f}%): {qty_threshold:,.0f} pcs")

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
        kandidat["score_velocity"] * SCORE_WEIGHT_VELOCITY
        + kandidat["score_margin"] * SCORE_WEIGHT_MARGIN
    ).round(1)

    for pct in PRICE_SCENARIOS:
        col = f"harga_+{int(pct * 100)}pct"
        kandidat[col] = (kandidat["harga_jual_avg"] * (1 + pct)).round(0)

    base = PRICE_SCENARIOS[0]
    kandidat[f"proyeksi_profit_+{int(base * 100)}pct"] = (
        kandidat["profit"] + kandidat["qty_terjual"] * kandidat["harga_jual_avg"] * base
    ).round(0)

    kandidat["saran"] = kandidat.apply(_saran_kandidat, axis=1)
    return kandidat.sort_values("score_total", ascending=False).reset_index(drop=True)


def _saran_kandidat(row: pd.Series) -> str:
    """Build advice string for a kandidat row."""
    if (row["restock_di_tahun"]
            and pd.notna(row["qty_setelah_restock"])
            and row["qty_setelah_restock"] > row["qty_terjual"] * 0.5):
        return "🔥 Restock cepat habis — naik 15-20%"
    if row["margin_pct"] > 50:
        return "✅ Margin tinggi, naik 10-15%"
    if row["score_total"] > 50:
        return "⭐ Top performer, naik 10%"
    return "Naik 5-10% bertahap, monitor"


def build_table_platform(jual: pd.DataFrame) -> pd.DataFrame:
    """Aggregate metrics per platform (Akun Penjual)."""
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
    """Top N SKUs by profit for a specific platform."""
    sub = jual[jual["akun_penjual"] == platform]
    return (sub.groupby("SKU").agg(
        qty=("qty_jual", "sum"),
        omzet=("omzet", "sum"),
        profit=("profit", "sum"),
    ).reset_index().sort_values("profit", ascending=False)
        .head(top_n).reset_index(drop=True))
