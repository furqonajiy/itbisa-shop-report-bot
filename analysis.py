"""Calculate HPP, profit per transaction, and aggregate metrics by SKU."""
from __future__ import annotations
from datetime import datetime

import numpy as np
import pandas as pd


def calculate_hpp_wa(stok: pd.DataFrame) -> pd.DataFrame:
    """Calculate weighted-average HPP per SKU from historical purchases.

    Formula: HPP_WA = sum(total_hpp_rp) / sum(qty_beli) grouped by SKU.
    """
    agg = stok.groupby("SKU").agg(
        total_qty_beli=("qty_beli", "sum"),
        total_hpp_sum=("total_hpp", "sum"),
        jumlah_pembelian=("SKU", "count"),
        tanggal_pembelian_terakhir=("tanggal_bayar", "max"),
    ).reset_index()
    agg["hpp_wa"] = agg["total_hpp_sum"] / agg["total_qty_beli"]
    print(f"✓ HPP weighted average dihitung untuk {len(agg)} SKU")
    return agg


def find_sku_without_hpp(jual: pd.DataFrame, hpp_agg: pd.DataFrame) -> list[str]:
    """Return SKUs present in jual but missing from stok (no HPP data)."""
    sku_jual = set(jual["SKU"].unique())
    sku_hpp = set(hpp_agg["SKU"].unique())
    return sorted(sku_jual - sku_hpp)


def enrich_with_profit(jual: pd.DataFrame, hpp_agg: pd.DataFrame) -> pd.DataFrame:
    """Add profit columns to jual. Rows without HPP are excluded with a warning.

    Profit formula: omzet - (hpp_wa * qty_jual) + tambahan + kode_unik.
    Tambahan and kode_unik are already signed (negative = biaya).
    """
    sku_no_hpp = find_sku_without_hpp(jual, hpp_agg)
    if sku_no_hpp:
        excluded_qty = jual[jual["SKU"].isin(sku_no_hpp)]["qty_jual"].sum()
        print(f"⚠ {len(sku_no_hpp)} SKU dijual tanpa data HPP "
              f"(di-exclude, {excluded_qty:,.0f} pcs):")
        for s in sku_no_hpp:
            print(f"    - {s}")

    jual = jual[~jual["SKU"].isin(sku_no_hpp)].copy()
    jual = jual.merge(hpp_agg[["SKU", "hpp_wa"]], on="SKU", how="left")
    jual["hpp_total"] = jual["hpp_wa"] * jual["qty_jual"]
    jual["biaya_admin"] = jual["tambahan"] + jual["kode_unik"]
    jual["profit"] = jual["omzet"] - jual["hpp_total"] + jual["biaya_admin"]
    return jual


def aggregate_by_sku(jual: pd.DataFrame, hpp_agg: pd.DataFrame, year: int) -> pd.DataFrame:
    """Aggregate enriched jual into per-SKU metrics, joined with stock info."""
    sku_agg = jual.groupby("SKU").agg(
        qty_terjual=("qty_jual", "sum"),
        jumlah_transaksi=("Invoice", "nunique"),
        omzet=("omzet", "sum"),
        hpp_cost=("hpp_total", "sum"),
        biaya_admin=("biaya_admin", "sum"),
        profit=("profit", "sum"),
        tgl_pertama_jual=("tanggal_pesan", "min"),
        tgl_terakhir_jual=("tanggal_pesan", "max"),
    ).reset_index()

    sku_agg["avg_qty_per_order"] = sku_agg["qty_terjual"] / sku_agg["jumlah_transaksi"]
    sku_agg["harga_jual_avg"] = sku_agg["omzet"] / sku_agg["qty_terjual"]
    sku_agg["profit_per_buah"] = sku_agg["profit"] / sku_agg["qty_terjual"]
    sku_agg["margin_pct"] = np.where(
        sku_agg["omzet"] > 0, sku_agg["profit"] / sku_agg["omzet"] * 100, np.nan
    )
    sku_agg["biaya_admin_per_buah"] = sku_agg["biaya_admin"] / sku_agg["qty_terjual"]

    sku_agg = sku_agg.merge(
        hpp_agg[["SKU", "hpp_wa", "total_qty_beli",
                 "tanggal_pembelian_terakhir", "jumlah_pembelian"]],
        on="SKU", how="left",
    )
    sku_agg["sisa_stok"] = sku_agg["total_qty_beli"] - sku_agg["qty_terjual"]

    year_start = datetime(year, 1, 1)
    sku_agg["restock_di_tahun"] = sku_agg["tanggal_pembelian_terakhir"] >= year_start
    return sku_agg


def calculate_qty_setelah_restock(sku_agg: pd.DataFrame, jual: pd.DataFrame) -> pd.DataFrame:
    """For SKUs with restock in analysis year, count qty sold after last restock."""
    def compute(row: pd.Series) -> float:
        if not row["restock_di_tahun"] or pd.isna(row["tanggal_pembelian_terakhir"]):
            return np.nan
        last_restock = row["tanggal_pembelian_terakhir"]
        post = jual[(jual["SKU"] == row["SKU"]) & (jual["tanggal_pesan"] >= last_restock)]
        return float(post["qty_jual"].sum()) if len(post) > 0 else 0.0

    sku_agg["qty_setelah_restock"] = sku_agg.apply(compute, axis=1)
    return sku_agg
