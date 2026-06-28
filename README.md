# ITBisa Shop — Sales Analysis Bot

A monthly/yearly sales-analysis tool for **IT Bisa Shop** built from marketplace export
data (Shopee, TikTok Shop, CoD, + legacy Tokopedia/Bukalapak). The output is a multi-sheet
Excel workbook: best-sellers, top profit contributors, loss-making items, price-increase
candidates, supplier analysis, reorder, and a per-warehouse stock recap.

## Setup

```bash
pip install -r requirements.txt   # pandas, openpyxl
```

> New here? `Cara_Pakai.md` is a step-by-step usage guide in Bahasa Indonesia (PowerShell), aimed at the shop operator.

Put the export files in the `data/` folder (any filename, matched by glob):

- Stock : `*Stok*.xlsx`   (sheet `Stok`)
- Sales : `*Jual*.xlsx`   (sheets `JualShopee`, `JualTiktok`, `JualTokopedia`,
  `JualBukalapak`, `JualCoD` — missing sheets are skipped automatically)

When several files match, the **latest file** (by name) is used as the "current workbook"
for the stock ledger; all files are used for sales history & HPP.

## Usage

```bash
python main.py                  # RUN EVERYTHING (= --all): sales + trend + reorder + cash-flow + dead-stock + A/B test + restock-check
python main.py --sales 2026     # sales report for a single year
python main.py --sales          # all years present in the sales data (= --sales all)
python main.py --trend          # sales trend & seasonality: cross-year omzet/profit trend, YoY growth, seasonal index
python main.py --reorder        # standalone reorder analysis
python main.py --cashflow       # cash-flow restock plan: how much capital is needed & when (per supplier)
python main.py --deadstock      # dead-stock / capital release: Rupiah frozen in slow/dead/overstock + how to free it
python main.py --ab-test        # A/B price-change test analysis (reads data/ab_tests.xlsx)
python main.py --restock-check  # restock price check & selling-price recommendation (reads data/restock_check.xlsx)
python main.py --stock-opname   # reconcile a physical count vs the ledger -> BisaHilang adjustment rows (reads data/stock_opname.xlsx)
python main.py --all            # everything together (same as no flag)
python main.py --laporan        # generate Laporan in laporan/ (every marketplace); scope: --laporan shopee tiktok
```

> With no flag, `python main.py` runs the full suite (same as `--all`). The A/B test and
> restock-check steps run only when their template (`data/ab_tests.xlsx` / `data/restock_check.xlsx`)
> already exists with rows — otherwise the step is skipped automatically (it never halts the
> run). Use the dedicated `--ab-test` / `--restock-check` flag to create the template the
> first time.

Results are written to the `output/` folder.

> **Laporan generator (`laporan/`)** — `python main.py --laporan` runs the co-located
> generator (an importable `laporan` package, called in-process; also runnable
> standalone as `python -m laporan`) that turns raw marketplace exports
> (`Transaksi`/`Saldo`/`Fee`) into `Laporan` workbooks — each delivered as a `Jual`
> sheet + a `Remit` sheet (the combined `Final` reconciliation). Its `Jual` feeds this bot's
> `Jual` ledger through a **manual Google Sheets step** (copy Laporan → Jual Sheets →
> export `Jual*.xlsx` into `data/`), so the generate and analyze stages are run
> separately. Reads `laporan/data`, writes `laporan/reports`.

## Key concepts

### HPP_WA vs pricing-basis HPP (pricing decision basis)

There are **two** HPP figures with different roles:

| | HPP_WA — "HPP/buah (P&L)" | Pricing-basis HPP (`hpp_pricing`) — "HPP/buah" |
|---|---|---|
| Definition | Weighted average of all relevant purchases | **Tiered current replacement cost** (see below) |
| Used for | **Profit & Margin** (P&L / realized truth) | **Pricing analysis**: Markup %, Kandidat Naik Harga, Borderline, the floor recommendation in Barang Rugi, and the restock cost in the cash-flow plan |

Rationale: for the question "should I raise the price?" the relevant basis is the **current
replacement cost**, not a lifetime weighted average still weighed down by old, expensive lots.
Profit/margin, conversely, keep using HPP_WA because it represents the cost actually realized.

The pricing-basis HPP ("HPP/buah", column was "HPP Dasar Harga") is chosen per SKU in tiers:

1. **Overseas SKU** — has any `Luar Negeri? = 1` purchase → the **latest overseas lot** price
   (by `Tanggal Bayar`). Overseas (direct-import) prices are consistent, so no averaging; the
   most recent lot is the current import cost.
2. **Purely-domestic SKU** — the **qty-weighted average of its domestic lots in the most recent
   restock window that has data**: the last **3 months**, else **6**, else **12**
   (`PRICING_HPP_WINDOWS_MONTHS`, anchored at today; Migrasi opening-balance and undated lots
   excluded). This tracks the current domestic restock cost rather than stale lots.
3. **Fallback** — no dated domestic lot in any window → the all-time **HPP_WA**.

- **Overseas (`Luar Negeri`)** is defined strictly: only rows with `Luar Negeri? = 1`
  (not guessed from the supplier name).
- The per-SKU pricing-HPP tier is shown in Sheet 07, column **"Sumber HPP Harga"**
  (`LN-terakhir` / `WA 3bln` / `WA 6bln` / `WA 12bln` / `WA`), next to **"HPP/buah"**
  (the pricing basis) and **"HPP/buah (P&L)"** (the realized HPP_WA).

### Harga Sekarang (baseline for the price-increase analysis)

The **"Harga Sekarang"** column (Sheets 04 & 05, also Sheet 07) = the **lowest unit price
(Omzet/Qty) on a SKU's most recent non-CoD selling day**.

- Not the lifetime average (`harga_jual_avg`), which mixes in old prices; it is taken from the
  **last selling day** so it represents the price in force now.
- The **minimum on that day** is used so it captures the lowest still-active wholesale tier,
  while ignoring (a) stale promo prices from early in the month and (b) the retail-vs-wholesale
  "lottery" of a single transaction that happens to be the most recent.
- **CoD is excluded** (its pricing is a different channel).
- **Markup %** (which decides whether a SKU lands in Kandidat ≥30% / Borderline <30%) is
  computed from `(Harga Sekarang − pricing-basis HPP) / pricing-basis HPP`, as are the
  Harga +10/15/20% scenarios in the Kandidat sheet.
- A SKU with no non-CoD sales at all → falls back to `harga_jual_avg`.

### "Harga baru naik" guard (Kandidat Naik Harga)

`Harga Sekarang` is a point-in-time value (the last selling day), but `Qty Terjual`/`Profit`
are cumulative. If the price was **just raised**, that qty & profit were earned at the
**older, cheaper** price — so recommending another increase and projecting a full year's qty
at the new price is **invalid**. Sheet 05 detects this and **holds** its recommendation
(row greyed out, the `Harga +10/15/20%` & `Proyeksi Profit` columns blanked, and the Saran
replaced with `⏳ Harga baru naik [tgl] (Rp lama→Rp baru); baru X% qty di harga baru —
kumpulkan data dulu, jangan naik lagi`).

- **Change date**: from `ab_tests.xlsx` when the SKU is logged (authoritative), else an
  **auto-detected** price jump (weighted average of the recent window vs the prior baseline).
- **Flagged only when** (a) `Harga Sekarang` is genuinely above the old price
  (≥ `PRICE_CHANGE_MIN_STEP`) **and** (b) demand at the new price is still thin
  (< `PRICE_CHANGE_VALIDATION_MIN_SHARE` of that year's qty). A long-stable price
  (e.g. Rp999 for months) is **not** flagged even if it is far above the average.
- Tunables in `config.py`: `PRICE_CHANGE_RECENT_DAYS`, `PRICE_CHANGE_MIN_STEP`,
  `PRICE_CHANGE_VALIDATION_MIN_SHARE`, `PRICE_CHANGE_AUTO_RECENT_DAYS`,
  `PRICE_CHANGE_AUTO_PRIOR_DAYS`, `PRICE_CHANGE_PRE_WINDOW_DAYS`. See
  `compute_price_change_status` in `analysis.py`.

### Markup, not margin

The minimum-price rule = **pricing-basis HPP × 1.30** (a 30% markup above HPP), **not** a 30%
gross margin. Constants: `TARGET_MARKUP_KOREKSI = 0.30`, `MARKUP_THRESHOLD_KANDIDAT = 30.0`.

### Migrasi (stock carry-over)

Rows whose Toko column starts with `Migrasi` = end-of-year stock carried into the next year's
file — **reference only, not a real purchase**. They are dropped automatically when there is
non-Migrasi purchase data for the same SKU (this prevents double-counting HPP and inflated
remaining stock).

### Reorder lead time (per shop, from data)

The reorder point (ROP) needs to know **how long goods take to arrive** so you do not run out
while waiting for a shipment. Lead time is a **property of the SHOP/forwarder, not per-SKU** —
AliExpress (~1 mo) is far faster than the Ocistok/Martkita sea-freight forwarder (~2.5 mo).
How it works:

1. **Per shop**: each import shop is the **75th percentile** (`LEAD_TIME_PERCENTILE`) of the
   `Tanggal Bayar` → `Tanggal Sampai` gap (Migrasi rows excluded); a shop with
   < `LEAD_TIME_MIN_LOTS` lots uses the global import percentile. **Ocistok = Martkita = 1688
   count as one forwarder** (Ocistok rebranded to Martkita; 1688 goes through them) —
   `OCISTOK_KEYWORDS` / `IMPORT_SHOP_KEYWORDS`.
2. **Per SKU**: take the lead of the **slowest shop supplying ≥ `LEAD_SHOP_MIN_SHARE` of that
   SKU's qty** (plan for the slow import, not the occasional local top-up). The forwarder comes
   from the **standardized `Toko` column** (e.g. `Jasa Impor Tiongkok`, `Ocistok/Martkita`,
   `AliExpress`, `Osell`), while import status comes from the **`Luar Negeri?`/China-keyword**
   qty share (authoritative); import SKUs are floored to the global import lead. Locally-sourced
   SKUs use `LEAD_TIME_MARKET_MONTHS` (≈ 1 week).

> **Data note**: `Toko` = the standardized source/forwarder; the payment-account detail +
> invoice/resi move to `Keterangan Pembelian`; `Luar Negeri?` remains the import flag
> (e.g. "Tokopedia Furqonajiy" paying Jasa Impor Tiongkok is now `Toko = Jasa Impor Tiongkok`).
> The loader automatically accepts the legacy `Toko[spasi]Akun Pemesan` header.

See `compute_lead_time_months` in `analysis.py`.

### On-order / in-transit stock (already ordered, not yet arrived)

The stock ledger only counts **arrived** lots (`Tanggal Sampai` filled), so a SKU you have
**already re-purchased** but whose shipment hasn't landed used to show low on-hand and get
re-flagged URGENT — telling you to buy something already on its way.

Now the reorder decision uses the **inventory position = on-hand + on-order**:

- **On-order (in transit)** = non-Migrasi lots that are paid (`Tanggal Bayar` filled) but not
  arrived (`Tanggal Sampai` blank), paid within `ONORDER_MAX_AGE_MONTHS` (default 6, so old
  blank-`Sampai` data gaps don't false-flag).
- The physical status is still classified on on-hand `sisa_stok` (a 0-on-hand SKU is still
  physically stocked out now), but the **suggested order nets out the incoming** (it only asks
  for the remainder), and a buy-now SKU (STOCKOUT/URGENT/Now/Soon) whose incoming order covers
  the ROP is relabelled **`⏳ Sudah Dipesan`** with an **estimated arrival** (last order date +
  lead time) — don't re-buy, wait for it to land. If the incoming is still short, it stays in
  its urgent bucket but only requests the top-up.
- New columns appear in `01_Reorder_Action` / `02_Reorder_Data_Lengkap`: `Qty Dipesan`,
  `Posisi` (on-hand + on-order), `Tgl Pesan`, `Est. Tiba`, plus a dedicated `⏳ Sudah Dipesan`
  section. (`--cashflow` is not yet position-aware; it still plans off on-hand.)

See `_compute_on_order` in `analysis.py`.

## Output sheet structure (sales report)

The yearly file `Analisa_Penjualan_ITBisa_<year>.xlsx` is **pure sales history (9 sheets)**.
Reorder and per-gudang stock are a current-day snapshot (not year-specific), so they live
only in `Analisa_Reorder.xlsx` (`--reorder`) — not duplicated into every yearly file.

| Sheet | Contents |
|---|---|
| 00_Summary | Summary of total omzet, profit, margin |
| 01_Paling_Diminati | Items with the highest avg qty/order |
| 02_Profit_Tertinggi | Largest profit contributors |
| 03_Barang_Rugi | Sold below cost (loss measured from HPP_WA; recommendation from pricing-basis HPP) |
| 04_Margin_Borderline | Markup < 30%, below the floor — must review |
| 05_Kandidat_Naik_Harga | Top qty + healthy markup → price-increase candidates |
| 06_Per_Platform | Margin & cost breakdown per marketplace |
| 07_Data_Lengkap_per_SKU | All per-SKU metrics (including pricing-basis HPP & its source) |
| 08_Supplier_Analysis | Supplier analysis (China vs local) |

The reorder workbook `Analisa_Reorder.xlsx` contains `00_Reorder_Summary`, `01_Reorder_Action`,
`02_Reorder_Data_Lengkap`, and `03_Rekap_Stok_per_Gudang`.

## Cash-flow restock plan (`--cashflow`)

Answers: **how much capital do I need to keep everything in stock, and when?** It turns the
reorder analysis into a purchasing-budget calendar — useful for an importer who pays suppliers
upfront and waits ~2.5 months for sea freight. No template needed; it is built entirely from
the stock/sales data, so it always runs in `--all`.

For each SKU with demand it projects **every** order due within the window (not just the next
one), via an inventory-position simulation:

- **When to order** = whenever the projected inventory position crosses the reorder point. A SKU
  at or below its ROP (or in STOCKOUT) is due now; after each order the position is replenished
  and then depletes at `velocity` until the next crossing. Only orders within
  `CASHFLOW_HORIZON_MONTHS` (default 6) are budgeted, and `CASHFLOW_MAX_CYCLES` caps the
  simulation. So a fast mover like NE555P gets several orders in the window.
- **How much** = the reorder qty (`target cover + lead demand − stock at order time`).
- **Cost** = qty × **replacement HPP** = `hpp_pricing` (the "HPP/buah" basis — latest overseas
  lot price, or the domestic restock WA of the last 3/6/12 months — what you would actually pay
  to restock now), falling back to `hpp_wa`.
- **Supplier** = the SKU's dominant standardized `Toko` (by non-Migrasi purchase qty), so spend
  is grouped by who you buy from.

Output: `output/Analisa_Cashflow_Restock.xlsx` — `00_Ringkasan` (total capital over the horizon
+ how much is due this month + a per-month table), `01_Kalender_per_Bulan` (a supplier × month
Rupiah matrix), and `02_Detail_per_SKU` (one row per order, with the order number per SKU).
See `cashflow.py`.

## Dead-stock / capital release (`--deadstock`)

Answers: **how much of my capital is stuck in stock that isn't moving, and what do I do about it?**
Built from the reorder metrics; no template needed, so it always runs in `--all`. It looks at the
SKUs the reorder analysis flags `🔵 Overstock` and `💤 Slow/Dead`.

- **Held value** = `sisa_stok × HPP_WA` — the capital tied up, valued at what you paid.
- **Freeable** = `max(0, sisa_stok − target_qty_post_reorder) × HPP_WA` — the excess above the
  reorder target, i.e. the actionable opportunity.
- **Recommendation**: 🧹 Likuidasi (no demand — velocity ≈ 0 or no sale in `DEADSTOCK_DEAD_DAYS`),
  🏷️ Markdown (slow turnover — cut price to speed it up), or ⛔ Stop reorder (healthy demand but
  far above target — stop buying / bundle to clear).

Output: `output/Analisa_Modal_Beku.xlsx` — `00_Ringkasan` (held vs freeable totals + top
opportunities), `01_Modal_Beku_per_SKU`, and `02_Per_Supplier` (whose goods are piling up).
See `deadstock_analysis.py`.

## Sales trend & seasonality (`--trend`)

Answers: **are we growing, and which months consistently sell best?** A cross-year view the
isolated per-year files can't give. No template needed; it always runs in `--all`.

- **Yearly trend** — omzet / qty / profit / orders per year, with **YoY** omzet growth (the
  current year is flagged partial).
- **Monthly trend** — the full omzet/profit time series across every month.
- **Seasonality** — a per-calendar-month index: each `(year, month)` omzet ÷ that year's own
  monthly average (so a growing business doesn't make later months look "seasonal"), averaged
  over the **complete** years only (the partial current year is excluded). Index > 1 = a month
  that consistently beats its year's average → a good time to stock up and push ads.
  Months are classed 🔥 Puncak / 🟢 above-avg / ⚪ Normal / 🔻 Sepi (`TREND_PEAK_INDEX`,
  `TREND_LOW_INDEX`); months with fewer than `TREND_SEASONAL_MIN_YEARS` years are marked thin.
- The summary also shows **year-to-date vs the same period last year**.

`profit = omzet + admin − HPP_WA × qty` (SKUs without HPP still count toward omzet, not profit).
Output: `output/Analisa_Tren_Musiman.xlsx` — `00_Ringkasan`, `01_Tren_Tahunan`, `02_Tren_Bulanan`,
`03_Musiman`. See `trend_analysis.py`.

## Stock opname → BisaHilang (`--stock-opname`)

Answers: **after a physical stock count, what Hilang/Ketemu adjustment makes the books match
reality?** Input `data/stock_opname.xlsx` (auto-created template: `SKU`, `Stok Fisik`, optional
`Lokasi Gudang` / `Tanggal Pengecekan` / `Keterangan`).

- Per SKU: `selisih = stok_buku − stok_fisik`, where `stok_buku` is the bot's ledger `sisa_stok`.
  `> 0` → **Banyak Hilang** (shrinkage); `< 0` → **Banyak Ketemu** (found).
- **Value** (`Nilai Hilang` / `Nilai Ketemu`) = `qty × HPP`, where HPP is `STOCK_OPNAME_VALUE_BASIS`
  (default **`hpp_wa`** — weighted-average cost, the realized inventory cost and standard
  write-off basis; set `hpp_pricing` for the latest/"HPP per buah" basis instead. **FIFO is not
  supported** — the tool keeps no purchase layers).
- **Lokasi Gudang** = the template's column if filled, else the SKU's dominant gudang in the
  ledger (else the busiest warehouse).

Output `output/BisaHilang_Rekonsiliasi.xlsx`: a **`BisaHilang`** sheet (discrepancy rows only,
in the exact tab format) ready to paste into the Google Sheets BisaHilang tab, plus a
`Ringkasan` audit sheet (every counted SKU + totals). Standalone (not part of `--all`). Paste,
re-export `BisaStok`, and `sisa_stok` lines up with your count. See `stock_opname.py`.

## Restock price check (`--restock-check`)

Answers: **is this supplier expensive/cheap/fair, and if I restock, what should I sell it for?**
Input `data/restock_check.xlsx` (SKU, Toko, `Harga RMB` and/or `HPP IDR`, `Kompetitor Min`/`Max`).

- **Landed-HPP prediction**: given an RMB price, the final HPP (Rp) is predicted with a factor
  **calibrated from the Ocistok/Martkita channel history** — `HPP per Buah (Rp) ÷ the (x RMB)`
  price from the `Keterangan` column (per-SKU when there are ≥ `RESTOCK_RMB_MIN_LOTS` lots, else
  the global median ≈ Rp`RMB_TO_IDR_FALLBACK`/RMB). This factor **already includes the Martkita
  margin + shipping + import** (≈25% above the spot rate `RMB_SPOT_FX_IDR`, because it is taken
  from the final cost actually paid). If `HPP IDR` is filled in directly, that value is used.
- **Cost verdict**: landed HPP vs the SKU's historical `hpp_wa` (±`RESTOCK_COST_TOL`) →
  cheaper / fair / more expensive.
- **Selling price per marketplace**: `HPP × (1 + RESTOCK_TARGET_NET_MARKUP) / (1 − fee)` so net
  ≥ target **after the fee**. Each marketplace fee is **derived from the** `Jual` **data**
  (`|admin|/omzet`, fallback `PLATFORM_FEE_FALLBACK`).
- **Decision** vs the competitor range: 🟢 restock & sell (target met within the market range),
  🟡 thin (profitable but below target), 🔴 don't sell (loss even at the highest market price).

Output: `output/Analisa_Restock_Check.xlsx`. See `restock_pricing.py`.

## A/B price-change test (`--ab-test`)

Measures **whether a price increase truly affects profit** — not merely pre vs post, because
profit moves for many reasons (trend, seasonality, stockouts). What is compared:

- **Matched pre-window**: `AB_PRE_WINDOW_DAYS` days (default 60) before the change date, not the
  lifetime average (an all-time baseline inflates the delta). All metrics use a daily rate
  (per day) to stay fair even when window lengths differ.
- **Profit bridge**: `Δprofit/day = Price Effect + Volume Effect + Interaction + Admin Effect`.
  It separates the extra margin from the price increase (price effect, +) from the impact of
  the volume change (volume effect). If the volume effect is strongly negative and outweighs the
  price effect → the increase is hurting.
- **Break-even qty drop** = `1 − (margin_pre / margin_post)`: how much % volume may be lost
  before profit returns to its previous level. **Headroom** = the distance from actual qty to
  that threshold (+ = safe).
- **Elasticity** = %Δqty ÷ %Δprice (diagnostic). When **positive** (qty & price both rise) →
  another factor is at play; the price effect cannot be isolated → flagged in the Catatan column.
- **Confound flags**: post too short / too few transactions, thin pre baseline, post qty
  dominated by one wholesale order, positive elasticity.
- **2-month validity gate**: a test younger than `AB_MIN_VALID_DAYS` (**60 days ≈ 2 months**
  post-change) is **not yet conclusive** — its verdict is forced to `⏳ In Progress (<2 bln)`
  (overriding Effective/Bad/Mixed/Pending) because the post-change qty/profit are still too thin
  to separate the price effect from seasonality/restock/promo noise. The `00_Summary` counts
  these `⏳ In Progress` tests separately, and `01_Test_Results` shows each test's
  `Masa Uji (hari)` (= `days_since_change`). Only ≥2-month tests get a conclusive read.
- **Verdict** is descriptive (✅/🟡/🔴/⚪) based on profit direction + break-even, downgraded to
  🟡 when attribution is weak — unless the validity gate forces `⏳ In Progress`. See `_verdict`
  / `analyze_ab_tests` in `ab_testing.py`.

Not yet including Difference-in-Differences (control SKUs) & bootstrap CI — a later step.

Config `data/ab_tests.xlsx` (sheet `ABTest`): `SKU`, `Tanggal Perubahan`, `Nama Test`,
`Catatan`. Output `output/Analisa_AB_Test.xlsx` (sheets `00_Summary`, `01_Test_Results`).

## Code structure

- `config.py` — constants (glob, sheet, thresholds, Migrasi prefix)
- `data_loader.py` — read & clean the stock/sales data
- `analysis.py` — HPP_WA, pricing-basis HPP (latest overseas lot / domestic 3-6-12mo WA), per-SKU aggregation, profit, reorder
- `tables.py` — build the analysis tables
- `excel_writer.py` — render the Excel workbook
- `ab_testing.py` — A/B price-change test analysis
- `restock_pricing.py` — restock price check & selling-price recommendation
- `stock_opname.py` — stock-opname reconciliation → BisaHilang adjustment rows
- `cashflow.py` — cash-flow restock plan (purchasing-budget calendar)
- `deadstock_analysis.py` — dead-stock / capital-release analysis
- `trend_analysis.py` — sales trend & seasonality analysis
- `main.py` — CLI entry point
