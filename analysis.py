"""HPP, profit, per-SKU aggregation, and reorder metrics."""
from __future__ import annotations
from datetime import datetime

import numpy as np
import pandas as pd

from config import (
    BULK_CHINA_SHARE_THRESHOLD, CHINA_KEYWORDS, CV_MODERATE_MAX, CV_STABLE_MAX,
    LEAD_TIME_CHINA_MONTHS, LEAD_TIME_MARKET_MONTHS, MARKET_KEYWORDS,
    OVERSTOCK_MONTHS, ROP_NOW_RATIO, ROP_SOON_RATIO, ROP_URGENT_RATIO,
    SAFETY_MULT_MODERATE, SAFETY_MULT_STABLE, SAFETY_MULT_VOLATILE,
    SLOW_DEAD_MAX_VELOCITY, TARGET_MONTHS_POST_REORDER,
    VELOCITY_MIN_ACTIVE_MONTHS,
)


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


def _velocity_window(jual_sku: pd.DataFrame, today: pd.Timestamp,
                     months: int) -> tuple[float, float, int, float]:
    """Return (avg_monthly_qty, std_monthly_qty, n_active_months, max_single_order)
    for the trailing window. Missing months padded as 0."""
    cutoff = today - pd.DateOffset(months=months)
    win = jual_sku[jual_sku["tanggal_pesan"] >= cutoff]
    if len(win) == 0:
        return 0.0, 0.0, 0, 0.0
    monthly = win.groupby(win["tanggal_pesan"].dt.to_period("M"))["qty_jual"].sum()
    all_m = pd.period_range(cutoff.to_period("M"), today.to_period("M"), freq="M")
    monthly = monthly.reindex(all_m, fill_value=0)
    return (
        float(monthly.mean()),
        float(monthly.std()) if len(monthly) > 1 else 0.0,
        int((monthly > 0).sum()),
        float(win["qty_jual"].max()),
    )


def _classify_volatility(cv: float) -> tuple[str, float]:
    if cv < CV_STABLE_MAX:
        return "Stabil", SAFETY_MULT_STABLE
    if cv < CV_MODERATE_MAX:
        return "Moderate", SAFETY_MULT_MODERATE
    return "Volatile", SAFETY_MULT_VOLATILE


def _classify_status(sisa: float, vel: float, rop: float,
                     months_cover: float) -> tuple[str, float]:
    if vel < SLOW_DEAD_MAX_VELOCITY:
        return "💤 Slow/Dead", -50.0
    if sisa <= 0:
        return "🔴 STOCKOUT", 100.0
    if sisa < rop * ROP_URGENT_RATIO:
        return "🔴 Reorder URGENT", 90.0 - min(80.0, months_cover * 20)
    if sisa < rop * ROP_NOW_RATIO:
        return "🟠 Reorder Now", 60.0 - min(40.0, months_cover * 10)
    if sisa < rop * ROP_SOON_RATIO:
        return "🟡 Reorder Soon", 30.0
    if months_cover > OVERSTOCK_MONTHS:
        return "🔵 Overstock", -10.0
    return "🟢 Healthy", 0.0


def _compute_china_share(stok: pd.DataFrame) -> pd.Series:
    """Per-SKU fraction of qty_beli that came from China sources."""
    stok = stok.copy()
    stok["_is_china"] = stok.apply(
        lambda r: classify_supplier(r.get("toko"), r.get("luar_negeri")) == "China", axis=1
    )
    qty_china = stok[stok["_is_china"]].groupby("SKU")["qty_beli"].sum()
    qty_total = stok.groupby("SKU")["qty_beli"].sum()
    return (qty_china / qty_total).fillna(0)


def compute_reorder_metrics(stok: pd.DataFrame, jual: pd.DataFrame,
                            today: pd.Timestamp | None = None) -> pd.DataFrame:
    """Build per-SKU reorder analysis DataFrame.
    `today` defaults to datetime.now() — controls velocity windows and 'as of' date.
    `jual` must be the FULL cleaned jual (all years), not year-filtered."""
    if today is None:
        today = pd.Timestamp(datetime.now().date())

    qty_beli_total = stok.groupby("SKU")["qty_beli"].sum()
    qty_jual_total = jual.groupby("SKU")["qty_jual"].sum()
    last_purchase = stok.groupby("SKU")["tanggal_bayar"].max()
    n_purchases = stok.groupby("SKU").size()
    omzet_total = jual.groupby("SKU")["omzet"].sum()
    china_share = _compute_china_share(stok)
    is_bulk_china = china_share >= BULK_CHINA_SHARE_THRESHOLD

    all_skus = sorted(set(qty_beli_total.index) | set(qty_jual_total.index))

    rows = []
    for sku in all_skus:
        sub = jual[jual["SKU"] == sku]
        v3, _, _, _ = _velocity_window(sub, today, 3)
        v6, std6, n6, max6 = _velocity_window(sub, today, 6)
        v12, std12, n12, max12 = _velocity_window(sub, today, 12)
        v24, std24, n24, _ = _velocity_window(sub, today, 24)

        if n6 >= VELOCITY_MIN_ACTIVE_MONTHS:
            vel, basis, cv = v6, "6mo", (std6 / v6 if v6 > 0 else 0.0)
        elif n12 >= VELOCITY_MIN_ACTIVE_MONTHS:
            vel, basis, cv = v12, "12mo", (std12 / v12 if v12 > 0 else 0.0)
        elif n24 >= 1:
            vel, basis, cv = v24, "24mo", (std24 / v24 if v24 > 0 else 0.0)
        else:
            vel, basis, cv = 0.0, "—", 0.0

        bulk_risk = max12 if max12 > 0 else max6

        qty_b = float(qty_beli_total.get(sku, 0))
        qty_s = float(qty_jual_total.get(sku, 0))
        sisa = qty_b - qty_s

        vol_cat, safety_mult = _classify_volatility(cv)
        is_china_bulk = bool(is_bulk_china.get(sku, False))
        lead = LEAD_TIME_CHINA_MONTHS if is_china_bulk else LEAD_TIME_MARKET_MONTHS
        lead_demand = vel * lead

        rop_safety = lead_demand * safety_mult
        rop_bulk = lead_demand + bulk_risk
        rop_final = max(rop_safety, rop_bulk)

        target_qty = vel * TARGET_MONTHS_POST_REORDER
        months_cover = sisa / vel if vel > 0 else float("inf")

        status, urgency = _classify_status(sisa, vel, rop_final, months_cover)

        if vel > 0 and sisa < rop_final:
            qty_suggest = max(0.0, target_qty - sisa + lead_demand)
        else:
            qty_suggest = 0.0

        rows.append({
            "SKU": sku,
            "sisa_stok": sisa,
            "qty_beli_all_time": qty_b,
            "qty_jual_all_time": qty_s,
            "v3mo": v3, "v6mo": v6, "v12mo": v12, "v24mo": v24,
            "velocity_basis": basis,
            "velocity_used": vel,
            "cv": cv,
            "volatility": vol_cat,
            "safety_mult": safety_mult,
            "max_single_order": bulk_risk,
            "bulk_china_share": float(china_share.get(sku, 0)),
            "is_bulk_china": is_china_bulk,
            "lead_months": lead,
            "lead_demand": lead_demand,
            "rop_safety": rop_safety,
            "rop_bulk": rop_bulk,
            "rop_final": rop_final,
            "target_qty_post_reorder": target_qty,
            "months_cover": months_cover,
            "status": status,
            "urgency_score": urgency,
            "qty_order_suggest": qty_suggest,
            "last_purchase": last_purchase.get(sku),
            "n_purchases": int(n_purchases.get(sku, 0)),
            "omzet_all_time": float(omzet_total.get(sku, 0)),
        })

    df = pd.DataFrame(rows)
    counts = df["status"].value_counts().to_dict()
    print(f"✓ Reorder analysis: {len(df):,} SKU "
          f"(STOCKOUT: {counts.get('🔴 STOCKOUT', 0)}, "
          f"URGENT: {counts.get('🔴 Reorder URGENT', 0)}, "
          f"Now: {counts.get('🟠 Reorder Now', 0)}, "
          f"Soon: {counts.get('🟡 Reorder Soon', 0)}, "
          f"Healthy: {counts.get('🟢 Healthy', 0)}, "
          f"Overstock: {counts.get('🔵 Overstock', 0)}, "
          f"Dead: {counts.get('💤 Slow/Dead', 0)})")
    return df