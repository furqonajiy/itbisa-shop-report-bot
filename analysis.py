"""HPP, profit, and per-SKU aggregation."""
from __future__ import annotations
from datetime import datetime

import numpy as np
import pandas as pd

from config import CHINA_KEYWORDS, MARKET_KEYWORDS


def classify_supplier(toko, luar_negeri) -> str:
    """Return 'China', 'Market', or 'Other'."""
    if pd.notna(luar_negeri) and luar_negeri == 1:
        return "China"
    if pd.isna(toko):
        return "Other"
    t = str(toko).lower()
    if any(k in t for k in CHINA_KEYWORDS):
        return "China"
    if any(k in t for k in MARKET_KEYWORDS):
        return "Market"
    return "Other"


def calculate_hpp_wa(stok: pd.DataFrame) -> pd.DataFrame:
    """HPP weighted average per SKU with Ocistok-priority.
    Rules:
    - If SKU has any Ocistok/Martkita (China) pembelian → HPP_WA from China rows only
    - Otherwise → HPP_WA from all available pembelian
    Inventory metrics (total_qty_beli, jumlah_pembelian) always use all rows."""
    stok = stok.copy()
    stok["supplier_type"] = stok.apply(
        lambda r: classify_supplier(r.get("toko"), r.get("luar_negeri")), axis=1
    )

    skus_with_ocistok = set(stok[stok["supplier_type"] == "China"]["SKU"].unique())
    in_hpp = (
        (stok["SKU"].isin(skus_with_ocistok) & (stok["supplier_type"] == "China"))
        | (~stok["SKU"].isin(skus_with_ocistok))
    )

    hpp_data = stok[in_hpp]
    hpp_agg = hpp_data.groupby("SKU").agg(
        _qty=("qty_beli", "sum"),
        _hpp=("total_hpp", "sum"),
    ).reset_index()
    hpp_agg["hpp_wa"] = hpp_agg["_hpp"] / hpp_agg["_qty"]
    hpp_agg = hpp_agg.drop(columns=["_qty", "_hpp"])

    inv_agg = stok.groupby("SKU").agg(
        total_qty_beli=("qty_beli", "sum"),
        jumlah_pembelian=("SKU", "count"),
        tanggal_pembelian_terakhir=("tanggal_bayar", "max"),
    ).reset_index()

    result = hpp_agg.merge(inv_agg, on="SKU", how="outer")
    result["hpp_source"] = result["SKU"].apply(
        lambda s: "Ocistok" if s in skus_with_ocistok else "Other"
    )

    n_oci = len(skus_with_ocistok)
    print(f"✓ HPP weighted average: {len(result):,} SKU "
          f"({n_oci:,} Ocistok-priority, {len(result)-n_oci:,} other-mix)")
    return result


def find_sku_without_hpp(jual: pd.DataFrame, hpp_agg: pd.DataFrame) -> list[str]:
    return sorted(set(jual["SKU"].unique()) - set(hpp_agg["SKU"].unique()))


def enrich_with_profit(jual: pd.DataFrame, hpp_agg: pd.DataFrame) -> pd.DataFrame:
    """Add profit columns; drop rows whose SKU has no HPP."""
    sku_no_hpp = find_sku_without_hpp(jual, hpp_agg)
    if sku_no_hpp:
        excluded_qty = jual[jual["SKU"].isin(sku_no_hpp)]["qty_jual"].sum()
        print(f"⚠ {len(sku_no_hpp)} SKU dijual tanpa HPP (excluded, {excluded_qty:,.0f} pcs):")
        for s in sku_no_hpp:
            print(f"    - {s}")

    jual = jual[~jual["SKU"].isin(sku_no_hpp)].copy()
    jual = jual.merge(hpp_agg[["SKU", "hpp_wa"]], on="SKU", how="left")
    jual["hpp_total"] = jual["hpp_wa"] * jual["qty_jual"]
    jual["biaya_admin"] = jual["tambahan"] + jual["kode_unik"]
    jual["profit"] = jual["omzet"] - jual["hpp_total"] + jual["biaya_admin"]
    return jual


def aggregate_by_sku(jual: pd.DataFrame, hpp_agg: pd.DataFrame, year: int,
                     qty_jual_all_time: pd.Series | None = None) -> pd.DataFrame:
    """Per-SKU aggregation joined with stock info.
    qty_jual_all_time: per-SKU total qty sold across all years (for accurate sisa_stok).
    If None, falls back to year-only qty (less accurate for sisa_stok)."""
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
        hpp_agg[["SKU", "hpp_wa", "hpp_source", "total_qty_beli",
                 "tanggal_pembelian_terakhir", "jumlah_pembelian"]],
        on="SKU", how="left",
    )

    sku_agg["markup_pct"] = np.where(
        sku_agg["hpp_wa"] > 0,
        (sku_agg["harga_jual_avg"] - sku_agg["hpp_wa"]) / sku_agg["hpp_wa"] * 100,
        np.nan,
    )

    if qty_jual_all_time is not None:
        sku_agg["qty_terjual_all_time"] = sku_agg["SKU"].map(qty_jual_all_time).fillna(0)
    else:
        sku_agg["qty_terjual_all_time"] = sku_agg["qty_terjual"]

    sku_agg["sisa_stok"] = sku_agg["total_qty_beli"] - sku_agg["qty_terjual_all_time"]
    sku_agg["restock_di_tahun"] = sku_agg["tanggal_pembelian_terakhir"] >= datetime(year, 1, 1)
    return sku_agg


def calculate_qty_setelah_restock(sku_agg: pd.DataFrame, jual: pd.DataFrame) -> pd.DataFrame:
    """For SKUs restocked in the analysis year, count qty sold after last restock."""
    def compute(row):
        if not row["restock_di_tahun"] or pd.isna(row["tanggal_pembelian_terakhir"]):
            return np.nan
        post = jual[(jual["SKU"] == row["SKU"])
                    & (jual["tanggal_pesan"] >= row["tanggal_pembelian_terakhir"])]
        return float(post["qty_jual"].sum()) if len(post) > 0 else 0.0

    sku_agg["qty_setelah_restock"] = sku_agg.apply(compute, axis=1)
    return sku_agg
