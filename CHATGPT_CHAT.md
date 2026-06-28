# ChatGPT Chat — Project Instructions (itbisa-shop-report-bot)

> **ChatGPT Chat** project instructions — a ≤8000-char condensed brief (ChatGPT's limit). For **Claude** use the full **`CLAUDE.md`**; **`AGENTS.md`** (ChatGPT Codex) points to `CLAUDE.md`. Update only when explicitly requested.

## What this is
Standalone **offline** Python tool turning ITBisa sales/stock Excel exports into analysis workbooks (sales, pricing, supplier, reorder + operational/pricing-strategy reports). No API/network/tokens/Actions — local, idempotent. Python 3.13 (3.11+); deps `pandas`, `openpyxl`.

## Stack & files (flat layout, no `src/`)
- `main.py` (CLI/orchestration), `config.py` (constants), `data_loader.py` (loading + SKU normalization + current-workbook loaders), `analysis.py` (HPP, profit, aggregation, supplier class, reorder, ledger, lead-time, price-change), `tables.py`, `excel_writer.py`, `ab_testing.py`, `restock_pricing.py`, `cashflow.py`, `deadstock_analysis.py`, `trend_analysis.py`, `stock_opname.py`. `data/` input (gitignored), `output/` reports.
- `laporan/` — co-located **Laporan generator** (importable `laporan` package; `python main.py --laporan` in-process, or `python -m laporan`): raw exports → `Laporan` workbooks, delivered as `Jual`/`Remit`(=combined `Final`)/`Bonus` (readers tolerate evolving TikTok/Shopee export quirks). Its `Jual` feeds this bot via a **manual Google Sheets step** (separate stages).

## CLI (`python main.py`)
- (no flag) = **full suite** (`--all`, 7 steps, loads data once via `_load_shared`, reorder once): sales+trend+reorder+cashflow+deadstock+ab-test+restock-check. Per-report flags: `--sales [YEAR]`/`--trend`/`--reorder`/`--cashflow`/`--deadstock`/`--ab-test`/`--restock-check`/`--stock-opname` (standalone), `--data-dir`/`--output-dir`. Zero-config reports always run; ab-test/restock-check run only if their template has rows. `--laporan [shopee tiktok …]` runs the `laporan` package in-process (`laporan/data`→`laporan/reports`).

## Inputs (`data/`, by glob)
- `*Stok*.xlsx` (purchases, sheet `Stok`; latest also needs `Hilang`+`PindahBarang`) and `*Jual*.xlsx` (sales, ≥ `JualShopee`). Legacy `Bisa*` file/sheet names still read transparently via `resolve_sheet()` (de-branding kept backward-compat; `ITBisa` brand unchanged). Stok **`Toko`** = standardized supplier/forwarder (Ocistok/Martkita, AliExpress, Jasa Impor, marketplaces, local); payment account + resi in **`Keterangan Pembelian`**; **`Luar Negeri?`** = authoritative overseas flag. Loader auto-accepts the legacy `Toko[spasi]Akun Pemesan` header.

## Output: `Analisa_Penjualan_ITBisa_<year>.xlsx` — 9 sheets, sales history
`00_Summary`, `01_Paling_Diminati`, `02_Profit_Tertinggi`, `03_Barang_Rugi`, `04_Margin_Borderline`, `05_Kandidat_Naik_Harga`, `06_Per_Platform`, `07_Data_Lengkap_per_SKU`, `08_Supplier_Analysis`. Reorder + per-gudang stock = current snapshot (not year-specific) → only in `Analisa_Reorder.xlsx` (`--reorder`: `00_Reorder_Summary`/`01_Reorder_Action`/`02_Reorder_Data_Lengkap`/`03_Rekap_Stok_per_Gudang`).

## Core logic (do not regress)
- **Sisa stok = current-workbook ledger** matching `RekapBarang`: per (SKU, gudang) `arrived beli − nonvoid jual + ketemu − hilang ± pindah`, from the latest stok/jual file by filename. Migrasi = opening balance (kept). Negative gudang floors to 0; OVERSOLD stays negative + flagged.
- **HPP = weighted average, Ocistok-priority**: if a SKU has Ocistok/Martkita China-direct buys, average those only, else all. Migrasi dropped when real purchases exist. **Supplier class**: China = `Luar Negeri?=1` or `Toko`∈`CHINA_KEYWORDS`; Market = `MARKET_KEYWORDS`; else Other. **No dedup** (only drop-Migrasi); SKU `UPPER().strip()`.
- **Pricing-decision basis** (markup % & price recs): `markup_pct=(harga_sekarang−hpp_pricing)/hpp_pricing`. `hpp_pricing` ("HPP/buah"; tier in sheet 07): latest `Luar Negeri?=1` lot for overseas → else domestic qty-weighted WA over the most recent `PRICING_HPP_WINDOWS_MONTHS` window with data (3→6→12 mo) → else `hpp_wa`. `harga_sekarang` = lowest non-CoD price on the latest selling day (else `harga_jual_avg`). **Profit/margin keep `hpp_wa`+`harga_jual_avg`** (P&L).
- **Recent-price-increase guard (Kandidat Naik Harga)**: a freshly-raised, under-validated `harga_sekarang` is **held** (`Harga +%`/`Proyeksi Profit` blanked) — qty/profit were earned at the OLD price. Change date from `ab_tests.xlsx` else auto step-detection; flags only when ≥`PRICE_CHANGE_MIN_STEP` over old AND post-change qty `<PRICE_CHANGE_VALIDATION_MIN_SHARE`.
- **Reorder**: velocity = avg monthly qty (total ÷ N months, trailing 6mo / fallback 12→24mo), CV → safety multiplier; **lead time is per-SHOP** from observed `Tanggal Bayar`→`Tanggal Sampai` at p75 (Ocistok=Martkita=1688 = one forwarder). A SKU takes the **slowest shop supplying ≥`LEAD_SHOP_MIN_SHARE` of its qty** (local → `LEAD_TIME_MARKET_MONTHS`). Then ROP + order qty; buckets STOCKOUT/URGENT/Now/Soon/Overstock + Slow/Dead. **On-order/in-transit**: decision uses inventory position = on-hand + on-order (paid-but-not-arrived lots within `ONORDER_MAX_AGE_MONTHS`); a buy-now SKU whose incoming covers ROP → `⏳ Sudah Dipesan`, qty nets out incoming. See `_compute_on_order`.
- **A/B test (`--ab-test`)**: matched pre-window vs post (daily rates), profit bridge + break-even, confounds. **2-month validity gate**: a test younger than `AB_MIN_VALID_DAYS` (60d ≈ 2 bln) is forced to `⏳ In Progress (<2 bln)` (overrides Effective/Bad/Mixed/Pending), counted separately in `00_Summary`; `01_Test_Results` shows `Masa Uji (hari)`.
- **Stock opname (`--stock-opname`, standalone)**: physical count vs ledger `sisa_stok` → `BisaHilang` rows (Banyak/Nilai Hilang/Ketemu, valued at `STOCK_OPNAME_VALUE_BASIS`=`hpp_wa`).
- **Restock check (`--restock-check`)**: landed HPP from `Harga RMB` (calibrated factor ≈Rp`RMB_TO_IDR_FALLBACK`/RMB) or given HPP IDR; verdict vs `hpp_wa`. Per-marketplace sell = HPP×(1+`RESTOCK_TARGET_NET_MARKUP`)/(1−fee), vs competitor range.
- **Cash-flow (`--cashflow`)**: reorder metrics → budget calendar; inv-position sim plans **every** cycle in `CASHFLOW_HORIZON_MONTHS` (cost qty×`hpp_pricing`, by supplier/month).
- **Dead-stock (`--deadstock`)**: capital in Overstock+Slow/Dead; held=sisa×hpp_wa, freeable=excess over `target_qty_post_reorder`; aksi likuidasi/markdown/stop.
- **Trend (`--trend`)**: cross-year omzet/profit + YoY + per-month seasonal index (complete years only; >1=puncak); headline YTD vs last yr. Outputs: `Analisa_<Tren_Musiman|Reorder|Cashflow_Restock|Modal_Beku|AB_Test|Restock_Check>.xlsx`; opname → `BisaHilang_Rekonsiliasi.xlsx`.
- **Summary (00)** surfaces a partial-year flag + data-quality block (OVERSOLD + sold-without-HPP). Tie-prone sorts (supplier/reorder) use an `SKU` tiebreaker (deterministic).

## Conventions & workflow
- Constants in `config.py` (never hardcode). Console strings Bahasa Indonesia; docs = English prose, Indonesian domain terms verbatim. Minimal, targeted changes.
- Branch `feature/<desc>` off `main`; doc/marker updates ride in the **same PR**. PR → **merge commit (`--no-ff`)**, title ends with the PR number. Authored **`C - Furqon Aji Yudhistira <furqonajiy@gmail.com>`**; **no AI/assistant references anywhere** (no Claude/Anthropic/`Co-Authored-By`/"Generated with"/session links); strip the auto-appended PR footer. CLI to the user = **PowerShell**. Sync marker `YYYY-MM-DD_HHMM.txt` (WIB) at root, renamed each update.

## Flag before changing
The stock-ledger reconciliation (`RekapBarang` parity), Ocistok-priority HPP, the pricing-decision basis + recent-price guard, the A/B 2-month validity gate, the per-shop reorder lead time + on-order rule + `config.py` tunables, the standardized `Toko` contract, SKU `UPPER().strip()`, the dedup/drop-Migrasi behavior, the 9-sheet yearly layout, the `resolve_sheet()` legacy fallback, the `SKU` sort tiebreaker, and the input globs.
