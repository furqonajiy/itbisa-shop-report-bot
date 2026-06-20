"""HPP, profit, per-SKU aggregation, and reorder metrics."""
from __future__ import annotations
from datetime import datetime

import numpy as np
import pandas as pd

from config import (
    BULK_CHINA_SHARE_THRESHOLD, CHINA_KEYWORDS, CV_MODERATE_MAX, CV_STABLE_MAX,
    IMPORT_SHOP_KEYWORDS, LEAD_SHOP_MIN_SHARE, LEAD_TIME_CHINA_MONTHS,
    LEAD_TIME_MARKET_MONTHS, LEAD_TIME_MAX_DAYS, LEAD_TIME_MIN_LOTS,
    LEAD_TIME_PERCENTILE, MARKET_KEYWORDS, MIGRASI_PREFIX, OVERSTOCK_MONTHS,
    PRICE_CHANGE_AUTO_PRIOR_DAYS,
    PRICE_CHANGE_AUTO_RECENT_DAYS, PRICE_CHANGE_MIN_STEP,
    PRICE_CHANGE_PRE_WINDOW_DAYS, PRICE_CHANGE_RECENT_DAYS,
    PRICE_CHANGE_VALIDATION_MIN_SHARE, ROP_NOW_RATIO, ROP_SOON_RATIO,
    ROP_URGENT_RATIO, SAFETY_MULT_MODERATE, SAFETY_MULT_STABLE,
    SAFETY_MULT_VOLATILE, SLOW_DEAD_MAX_VELOCITY, TARGET_MONTHS_POST_REORDER,
    VELOCITY_MIN_ACTIVE_MONTHS,
)

_DAYS_PER_MONTH = 365.25 / 12.0


def _canon_shop(toko) -> str:
    """Canonical supplier shop for lead-time grouping: an import forwarder name from
    IMPORT_SHOP_KEYWORDS, else 'Local' (marketplaces / domestic distributors that
    ship in days). Ocistok = Martkita = 1688 collapse to one forwarder."""
    t = str(toko).lower()
    for shop, kws in IMPORT_SHOP_KEYWORDS.items():
        if any(k in t for k in kws):
            return shop
    return "Local"


def compute_lead_time_months(stok: pd.DataFrame) -> tuple[pd.Series, float]:
    """Reorder lead time (months) derived PER SHOP (the forwarder), then assigned to
    each SKU via its primary shop — because lead time is a property of who you buy
    from, not the item (AliExpress ships faster than the Ocistok/Martkita sea-freight
    forwarder; Ocistok=Martkita=1688 are one shop). Two steps:
      1) For each import shop, the LEAD_TIME_PERCENTILE (p75) of observed shipping
         time (Tanggal Bayar→Sampai, non-Migrasi); a shop with < LEAD_TIME_MIN_LOTS
         dated lots falls back to the global import percentile.
      2) Each SKU inherits the lead of the shop supplying most of its purchase qty.
         SKUs sourced mainly from local shops use LEAD_TIME_MARKET_MONTHS.

    A mixed-sourced SKU takes the lead of the SLOWEST shop supplying ≥
    LEAD_SHOP_MIN_SHARE of its qty (plan for the slow import, not the occasional fast
    local top-up); import SKUs are floored at the global import lead.

    Returns (per_sku_lead: Series, global_import_lead: float). Falls back to
    LEAD_TIME_CHINA_MONTHS if there is no usable import-shipping history at all."""
    cols = {"toko", "tanggal_bayar", "tanggal_sampai", "qty_beli"}
    if not cols.issubset(stok.columns):
        return pd.Series(dtype=float), LEAD_TIME_CHINA_MONTHS

    s = stok.copy()
    s = s[~s["toko"].astype(str).str.startswith(MIGRASI_PREFIX, na=False)].copy()
    s["_shop"] = s["toko"].map(_canon_shop)
    s["_lead"] = (s["tanggal_sampai"] - s["tanggal_bayar"]).dt.days

    # 1) per-import-shop lead (p75 of dated lots); thin shops → global import lead.
    dated = s[(s["_shop"] != "Local") & s["_lead"].between(0, LEAD_TIME_MAX_DAYS)]
    if len(dated) == 0:
        return pd.Series(dtype=float), LEAD_TIME_CHINA_MONTHS
    global_import = float(dated["_lead"].quantile(LEAD_TIME_PERCENTILE)) / _DAYS_PER_MONTH
    shop_lead = dated.groupby("_shop")["_lead"].quantile(LEAD_TIME_PERCENTILE) / _DAYS_PER_MONTH
    shop_n = dated.groupby("_shop")["_lead"].count()
    shop_lead = shop_lead.where(shop_n >= LEAD_TIME_MIN_LOTS, global_import)

    # 2) Per-SKU lead = the slowest shop with a meaningful qty share. Whether a SKU is
    #    an import is taken from the reliable Luar Negeri / China-keyword qty share
    #    (NOT the toko name — many import lots are booked under a local marketplace
    #    ACCOUNT like "Tokopedia Furqonajiy", Luar Negeri=1); import SKUs are floored
    #    at the global import lead so a hidden-forwarder import isn't planned as local.
    def _shop_lead(sh):
        return LEAD_TIME_MARKET_MONTHS if sh == "Local" else float(shop_lead.get(sh, global_import))

    is_import = _compute_china_share(stok) >= BULK_CHINA_SHARE_THRESHOLD
    valid = s[s["qty_beli"] > 0]
    shop_qty = valid.groupby(["SKU", "_shop"])["qty_beli"].sum()
    sku_total = valid.groupby("SKU")["qty_beli"].sum()

    per_sku = {}
    for sku in sku_total.index:
        shares = shop_qty[sku] / sku_total[sku]
        leads = [_shop_lead(sh) for sh, sh_share in shares.items() if sh_share >= LEAD_SHOP_MIN_SHARE]
        lead = max(leads) if leads else LEAD_TIME_MARKET_MONTHS
        if bool(is_import.get(sku, False)):
            lead = max(lead, global_import)
        per_sku[sku] = lead
    per_sku = pd.Series(per_sku, dtype=float)

    shops_txt = ", ".join(f"{sh}={lead:.2f}bln(n{int(shop_n.get(sh, 0))})"
                          for sh, lead in shop_lead.sort_values().items())
    print(f"✓ Lead time per-shop (p{int(LEAD_TIME_PERCENTILE*100)}): {shops_txt}; "
          f"global-impor {global_import:.2f} bln; {len(per_sku):,} SKU dipetakan (shop terlambat ≥{int(LEAD_SHOP_MIN_SHARE*100)}% qty)")
    return per_sku, global_import


def build_stock_ledger(stok_arrived: pd.DataFrame, jual_nonvoid: pd.DataFrame,
                       hilang: pd.DataFrame, pindah: pd.DataFrame
                       ) -> tuple[pd.DataFrame, pd.Series]:
    """Reproduce RekapBarang from the current workbook, per (SKU, gudang):

        stok = Σ beli(arrived) − Σ jual(non-void) + Σ ketemu − Σ hilang
               + Σ pindah_masuk − Σ pindah_keluar

    Returns (ledger_df, sisa_by_sku):
      - ledger_df: one row per SKU, one column per gudang, plus 'Total'
      - sisa_by_sku: Series SKU -> total on-hand (authoritative sisa_stok)
    Migrasi rows are included in `stok_arrived` (they are the opening balance).
    Pindah nets to zero at the SKU total; it only moves stock between gudang.
    """
    def by_gudang(df, gcol, vcol):
        if df is None or len(df) == 0:
            return pd.Series(dtype=float)
        s = df.groupby(["SKU", gcol])[vcol].sum()
        s.index = s.index.set_names(["SKU", "gudang"])
        return s

    beli = by_gudang(stok_arrived, "gudang", "qty")
    jual = by_gudang(jual_nonvoid, "gudang", "qty")
    ketemu = by_gudang(hilang, "gudang", "ketemu")
    hil = by_gudang(hilang, "gudang", "hilang")
    pin = by_gudang(pindah, "gudang_in", "qty")
    pout = by_gudang(pindah, "gudang_out", "qty")

    net = (beli.subtract(jual, fill_value=0)
               .add(ketemu, fill_value=0)
               .subtract(hil, fill_value=0)
               .add(pin, fill_value=0)
               .subtract(pout, fill_value=0))
    net.name = "qty"

    flat = net.reset_index()
    flat = flat[flat["gudang"].astype(str).str.strip().ne("")]
    # The rekap only tracks SKUs that exist in the stock master (have ≥1 purchase row,
    # incl. Migrasi opening). A SKU with only stray sales and no purchase (e.g. a
    # mis-keyed SKU) is not on-hand stock and is excluded — matching RekapBarang.
    valid_skus = set(stok_arrived["SKU"].unique()) if len(stok_arrived) else set()
    flat = flat[flat["SKU"].isin(valid_skus)]
    ledger = flat.pivot_table(index="SKU", columns="gudang", values="qty",
                              aggfunc="sum", fill_value=0)
    gudang_cols = sorted(ledger.columns)
    ledger = ledger[gudang_cols]
    ledger["Total"] = ledger.sum(axis=1)

    # A negative per-gudang balance is physically impossible — it means a sale or
    # transfer was tagged to a gudang that didn't hold the stock. The Sheets rekap
    # collapses this: floor the negative to 0 and keep the SKU total intact (the
    # stock physically sits in the other gudang). This reproduces RekapBarang's
    # per-gudang columns exactly while leaving the (authoritative) Total unchanged.
    flagged = []
    oversold = []
    raw_g = ledger[gudang_cols]
    neg_mask = (raw_g < 0).any(axis=1)
    if neg_mask.any():
        clamped = raw_g.clip(lower=0)
        for sku in ledger.index[neg_mask]:
            total_sku = ledger.loc[sku, "Total"]
            if total_sku >= 0:
                surplus = clamped.loc[sku].sum() - total_sku
                if surplus > 1e-9:
                    gmax = clamped.loc[sku].idxmax()
                    clamped.loc[sku, gmax] -= surplus
                flagged.append(sku)
            else:
                # Oversold (negative SKU total) — a genuine inventory discrepancy in
                # the source, not a tagging quirk. Keep the raw split so it stays
                # visible for the user to investigate.
                clamped.loc[sku] = raw_g.loc[sku]
                oversold.append(sku)
        ledger[gudang_cols] = clamped

    ledger = ledger.sort_index().reset_index()

    sisa_by_sku = ledger.set_index("SKU")["Total"]
    print(f"✓ Ledger stok (workbook berjalan): {len(ledger):,} SKU, "
          f"gudang {gudang_cols} — total on-hand {sisa_by_sku.sum():,.0f} pcs")
    if flagged:
        print(f"  ⚠ {len(flagged)} SKU saldo gudang negatif (jual/pindah salah tag) "
              f"— dikoreksi agar cocok rekap: " + ", ".join(flagged))
    if oversold:
        print(f"  ⚠ {len(oversold)} SKU OVERSOLD (total stok negatif — cek data sumber): "
              + ", ".join(oversold))
    return ledger, sisa_by_sku


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


def _latest_ln_price(stok: pd.DataFrame) -> pd.Series:
    """Per-unit HPP of each SKU's most recent overseas lot (Luar Negeri? == 1),
    by Tanggal Bayar. Basis for the pricing-decision markup (harga supplier LN
    terbaru), per request — no weighted average since LN price is consistent.
    Strict definition: only rows flagged Luar Negeri? == 1 (not keyword-China).
    Returns Series SKU -> unit price; SKUs without an LN lot are absent."""
    ln = stok[stok.get("luar_negeri") == 1].copy()
    ln = ln[(ln["qty_beli"] > 0) & ln["total_hpp"].notna()]
    if len(ln) == 0:
        return pd.Series(dtype=float)
    ln["unit_hpp"] = ln["total_hpp"] / ln["qty_beli"]
    # na_position='first' keeps dated lots last, so tail(1) picks the latest real
    # date; if every LN lot is undated, it falls back to the last row available.
    ln = ln.sort_values("tanggal_bayar", na_position="first")
    return ln.groupby("SKU").tail(1).set_index("SKU")["unit_hpp"]


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

    # Pricing-decision HPP ("apakah harus naik harga?"): latest overseas lot price
    # for SKUs with LN purchases, else fall back to hpp_wa. Used ONLY for markup %
    # and price recommendations — profit/margin keep hpp_wa.
    ln_latest = _latest_ln_price(stok)
    result["hpp_pricing"] = result["SKU"].map(ln_latest)
    result["hpp_pricing_source"] = np.where(
        result["hpp_pricing"].notna(), "LN-terakhir", "WA"
    )
    result["hpp_pricing"] = result["hpp_pricing"].fillna(result["hpp_wa"])

    n_oci = len(skus_with_ocistok)
    n_ln = int(ln_latest.index.isin(result["SKU"]).sum())
    print(f"✓ HPP weighted average: {len(result):,} SKU "
          f"({n_oci:,} Ocistok-priority, {len(result)-n_oci:,} other-mix)")
    print(f"  → HPP dasar harga: {n_ln:,} SKU pakai harga LN terakhir, "
          f"{len(result)-n_ln:,} fallback ke WA")
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


def compute_harga_sekarang(jual_full_clean: pd.DataFrame) -> pd.Series:
    """Per-SKU 'harga sekarang' for the pricing decision: the LOWEST unit price
    (Omzet/Qty) among non-CoD sales on the SKU's most recent selling day.

    Rationale: the lifetime weighted average mixes old and new prices; the latest
    selling day reflects the current price level. Taking the minimum across that
    day captures the lowest active tier (grosir) while ignoring (a) stale promo
    prices from earlier in the month and (b) the retail-vs-grosir lottery of which
    single order happened to be last. CoD is excluded (its prices differ).
    SKUs with no non-CoD sale are absent (caller falls back to harga_jual_avg)."""
    d = jual_full_clean
    if "_sheet_source" in d.columns:
        d = d[d["_sheet_source"] != "JualCoD"]
    d = d[(d["qty_jual"] > 0) & (d["omzet"] > 0) & d["tanggal_pesan"].notna()].copy()
    if len(d) == 0:
        return pd.Series(dtype=float)
    d["_unit"] = d["omzet"] / d["qty_jual"]
    d["_hari"] = d["tanggal_pesan"].dt.normalize()
    last_day = d.groupby("SKU")["_hari"].transform("max")
    return d[d["_hari"] == last_day].groupby("SKU")["_unit"].min()


def _wavg_price(jual_sub: pd.DataFrame) -> float:
    """Qty-weighted average unit price (Omzet/Qty) over a slice; NaN if empty."""
    q = jual_sub["qty_jual"].sum()
    return float(jual_sub["omzet"].sum() / q) if q > 0 else float("nan")


def compute_price_change_status(jual_full_clean: pd.DataFrame,
                                harga_sekarang: pd.Series,
                                ab_changes: pd.Series | None,
                                today: pd.Timestamp, year: int) -> pd.DataFrame:
    """Per-SKU detection of a RECENT, not-yet-validated price INCREASE.

    Why: 'harga sekarang' is a point-in-time figure (latest selling day), but
    qty_terjual/profit are cumulative. If the current price is a recent hike, the
    historical demand was earned at the OLD price — so recommending a further raise
    and projecting profit on the full-year qty at the new price is invalid. Such a
    SKU should be flagged and held, not recommended, until demand accrues at the
    new price. (See the Kandidat Naik Harga guard in tables.py.)

    Change date is taken from ab_tests.xlsx when logged (authoritative), else from a
    two-window auto step-detection (recent qty-weighted avg vs prior baseline). The
    flag fires only when (a) harga_sekarang is genuinely above the pre-change price
    (≥ MIN_STEP) and (b) the post-change demand is thin (< VALIDATION_MIN_SHARE of
    the year's qty), i.e. the new price level is under-validated.

    Returns a DataFrame indexed by SKU with: harga_lama, tgl_naik, qty_validasi,
    share_validasi, sumber_perubahan ('ab_test'/'auto'), harga_baru_flag.
    SKUs with no recent change are absent (caller treats them as not flagged)."""
    d = jual_full_clean
    if "_sheet_source" in d.columns:
        d = d[d["_sheet_source"] != "JualCoD"]
    d = d[(d["qty_jual"] > 0) & (d["omzet"] > 0) & d["tanggal_pesan"].notna()].copy()
    if len(d) == 0:
        return pd.DataFrame()
    d["_unit"] = d["omzet"] / d["qty_jual"]
    ab_changes = ab_changes if ab_changes is not None else pd.Series(dtype="datetime64[ns]")
    recent_cut = today - pd.Timedelta(days=PRICE_CHANGE_RECENT_DAYS)

    rows = {}
    for sku, s in d.groupby("SKU"):
        hs = harga_sekarang.get(sku)
        if pd.isna(hs):
            continue
        last = s["tanggal_pesan"].max()

        # 1) Change date: ab_tests (authoritative) → else auto two-window step.
        tgl_naik, sumber = None, ""
        abd = ab_changes.get(sku)
        if pd.notna(abd) and recent_cut <= pd.Timestamp(abd) <= today:
            tgl_naik, sumber = pd.Timestamp(abd).normalize(), "ab_test"
        else:
            rec = s[s["tanggal_pesan"] >= today - pd.Timedelta(days=PRICE_CHANGE_AUTO_RECENT_DAYS)]
            prior = s[(s["tanggal_pesan"] < today - pd.Timedelta(days=PRICE_CHANGE_AUTO_RECENT_DAYS))
                      & (s["tanggal_pesan"] >= today - pd.Timedelta(
                          days=PRICE_CHANGE_AUTO_RECENT_DAYS + PRICE_CHANGE_AUTO_PRIOR_DAYS))]
            p_rec, p_prior = _wavg_price(rec), _wavg_price(prior)
            if pd.notna(p_rec) and pd.notna(p_prior) and p_rec >= p_prior * (1 + PRICE_CHANGE_MIN_STEP):
                stepped = rec[rec["_unit"] >= p_prior * (1 + PRICE_CHANGE_MIN_STEP)]
                tgl_naik = (stepped["tanggal_pesan"].min().normalize() if len(stepped)
                            else (today - pd.Timedelta(days=PRICE_CHANGE_AUTO_RECENT_DAYS)).normalize())
                sumber = "auto"
        if tgl_naik is None:
            continue

        # 2) Pre-change price (harga_lama) and post-change validation share.
        pre = s[(s["tanggal_pesan"] >= tgl_naik - pd.Timedelta(days=PRICE_CHANGE_PRE_WINDOW_DAYS))
                & (s["tanggal_pesan"] < tgl_naik)]
        harga_lama = _wavg_price(pre)
        yr = s[s["tanggal_pesan"].dt.year == year]
        post_yr = yr[yr["tanggal_pesan"] >= tgl_naik]
        qty_year = yr["qty_jual"].sum()
        qty_validasi = float(post_yr["qty_jual"].sum())
        share = qty_validasi / qty_year if qty_year > 0 else float("nan")

        # 3) Flag only a genuine, under-validated increase.
        is_increase = pd.notna(harga_lama) and hs > harga_lama * (1 + PRICE_CHANGE_MIN_STEP)
        flag = bool(is_increase and pd.notna(share)
                    and share < PRICE_CHANGE_VALIDATION_MIN_SHARE)
        rows[sku] = {
            "harga_lama": harga_lama,
            "tgl_naik": tgl_naik,
            "qty_validasi": qty_validasi,
            "share_validasi": share,
            "sumber_perubahan": sumber,
            "harga_baru_flag": flag,
        }

    df = pd.DataFrame.from_dict(rows, orient="index")
    n_flag = int(df["harga_baru_flag"].sum()) if len(df) else 0
    print(f"✓ Cek harga baru naik: {len(df):,} SKU dgn perubahan terdeteksi, "
          f"{n_flag:,} ditandai 'baru naik & belum tervalidasi' (di-hold dari rekomendasi)")
    return df


def aggregate_by_sku(jual: pd.DataFrame, hpp_agg: pd.DataFrame, year: int,
                     qty_jual_all_time: pd.Series | None = None,
                     sisa_by_sku: pd.Series | None = None,
                     harga_sekarang: pd.Series | None = None,
                     price_change: pd.DataFrame | None = None) -> pd.DataFrame:
    """Per-SKU aggregation joined with stock info.
    qty_jual_all_time: per-SKU total qty sold across all years (for fallback sisa).
    sisa_by_sku: authoritative on-hand from the current-workbook ledger
    (RekapBarang). When provided, sisa_stok comes from it (SKUs not in the
    current workbook → 0). When None, falls back to all-time beli − jual."""
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
        hpp_agg[["SKU", "hpp_wa", "hpp_source", "hpp_pricing", "hpp_pricing_source",
                 "total_qty_beli", "tanggal_pembelian_terakhir", "jumlah_pembelian"]],
        on="SKU", how="left",
    )

    # 'Harga sekarang' = lowest non-CoD unit price on the SKU's most recent selling
    # day (see compute_harga_sekarang). Fallback to the lifetime weighted-average
    # sell price for SKUs with no non-CoD sale. Used for the pricing decision only.
    if harga_sekarang is not None:
        sku_agg["harga_sekarang"] = sku_agg["SKU"].map(harga_sekarang)
    else:
        sku_agg["harga_sekarang"] = np.nan
    sku_agg["harga_sekarang"] = sku_agg["harga_sekarang"].fillna(sku_agg["harga_jual_avg"])

    # Markup is a pricing signal ("naik harga?"), so it compares harga_sekarang
    # against hpp_pricing (latest LN lot price where available, else hpp_wa).
    # Profit/margin still use harga_jual_avg and hpp_wa (realized).
    sku_agg["markup_pct"] = np.where(
        sku_agg["hpp_pricing"] > 0,
        (sku_agg["harga_sekarang"] - sku_agg["hpp_pricing"]) / sku_agg["hpp_pricing"] * 100,
        np.nan,
        )

    if qty_jual_all_time is not None:
        sku_agg["qty_terjual_all_time"] = sku_agg["SKU"].map(qty_jual_all_time).fillna(0)
    else:
        sku_agg["qty_terjual_all_time"] = sku_agg["qty_terjual"]

    if sisa_by_sku is not None:
        # On-hand stock per RekapBarang (current workbook). SKUs absent from the
        # current workbook are not in stock → 0.
        sku_agg["sisa_stok"] = sku_agg["SKU"].map(sisa_by_sku).fillna(0)
    else:
        sku_agg["sisa_stok"] = sku_agg["total_qty_beli"] - sku_agg["qty_terjual_all_time"]
    sku_agg["restock_di_tahun"] = sku_agg["tanggal_pembelian_terakhir"] >= datetime(year, 1, 1)

    # Recent-price-increase guard (see compute_price_change_status). When a SKU's
    # current price is a freshly-raised, under-validated price, the pricing
    # recommendation built on pre-increase demand is held back downstream.
    pc_cols = ["harga_lama", "tgl_naik", "qty_validasi", "share_validasi",
               "sumber_perubahan", "harga_baru_flag"]
    if price_change is not None and len(price_change) > 0:
        pc = price_change.reindex(columns=pc_cols).reset_index().rename(columns={"index": "SKU"})
        sku_agg = sku_agg.merge(pc, on="SKU", how="left")
    else:
        for c in pc_cols:
            sku_agg[c] = np.nan
    sku_agg["harga_baru_flag"] = sku_agg["harga_baru_flag"].fillna(False).astype(bool)
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
    for the trailing `months`-month window (by date). Average = total qty ÷ `months`
    (the nominal window length), NOT ÷ the number of calendar buckets the window
    touches: a date cutoff straddles months+1 buckets, so dividing by the bucket
    count understates the monthly rate by ~1/(months+1). The monthly buckets are
    still used for the volatility (std) and active-month count."""
    cutoff = today - pd.DateOffset(months=months)
    win = jual_sku[jual_sku["tanggal_pesan"] >= cutoff]
    if len(win) == 0:
        return 0.0, 0.0, 0, 0.0
    monthly = win.groupby(win["tanggal_pesan"].dt.to_period("M"))["qty_jual"].sum()
    all_m = pd.period_range(cutoff.to_period("M"), today.to_period("M"), freq="M")
    monthly = monthly.reindex(all_m, fill_value=0)
    return (
        float(win["qty_jual"].sum()) / months,
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
                            today: pd.Timestamp | None = None,
                            sisa_by_sku: pd.Series | None = None) -> pd.DataFrame:
    """Build per-SKU reorder analysis DataFrame.
    `today` defaults to datetime.now() — controls velocity windows and 'as of' date.
    `jual` must be the FULL cleaned jual (all years), not year-filtered.
    `sisa_by_sku`: authoritative on-hand from the current-workbook ledger. When
    provided, sisa_stok comes from it (SKUs not in current workbook → 0); velocity
    still uses all-time jual. When None, falls back to all-time beli − jual."""
    if today is None:
        today = pd.Timestamp(datetime.now().date())

    qty_beli_total = stok.groupby("SKU")["qty_beli"].sum()
    qty_jual_total = jual.groupby("SKU")["qty_jual"].sum()
    last_purchase = stok.groupby("SKU")["tanggal_bayar"].max()
    n_purchases = stok.groupby("SKU").size()
    omzet_total = jual.groupby("SKU")["omzet"].sum()
    china_share = _compute_china_share(stok)
    is_bulk_china = china_share >= BULK_CHINA_SHARE_THRESHOLD
    # Per-shop observed lead time mapped to each SKU's primary supplier (replaces the
    # old flat LEAD_TIME_CHINA_MONTHS guess). See compute_lead_time_months.
    per_sku_lead, global_import_lead = compute_lead_time_months(stok)

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
        if sisa_by_sku is not None:
            sisa = float(sisa_by_sku.get(sku, 0.0))   # on-hand per RekapBarang
        else:
            sisa = qty_b - qty_s

        vol_cat, safety_mult = _classify_volatility(cv)
        is_china_bulk = bool(is_bulk_china.get(sku, False))
        # Lead = the SKU's primary supplier shop's observed shipping time. SKUs with
        # no purchase history fall back: global import lead if China-heavy, else market.
        lead = per_sku_lead.get(sku)
        if lead is None or pd.isna(lead):
            lead = global_import_lead if is_china_bulk else LEAD_TIME_MARKET_MONTHS
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
