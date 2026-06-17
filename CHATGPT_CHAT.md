# ChatGPT Chat — Project Instructions (itbisa-shop-report-bot)

> **ChatGPT Chat** project instructions — a ≤8000-char condensed brief (ChatGPT's limit). For **Claude** (Code *and* Chat) use the full **`CLAUDE.md`** directly; **`AGENTS.md`** (ChatGPT Codex) points to `CLAUDE.md`. Update only when explicitly requested.

## What this is
Standalone, **offline** Python tool that turns ITBisa sales/stock Excel exports into analysis workbooks (sales performance, pricing, supplier, reorder + operational & pricing-strategy reports). No API/network/tokens/GitHub Actions — runs locally, idempotent.

## Stack & files (flat layout, no `src/`)
- Python 3.10+. Deps: `pandas`, `openpyxl`.
- `main.py` CLI/orchestration · `config.py` all constants · `data_loader.py` loading + SKU normalization + current-workbook loaders · `analysis.py` HPP, profit, aggregation, supplier classification, reorder, `build_stock_ledger`, `compute_lead_time_months`, `compute_price_change_status` · `tables.py` table builders · `excel_writer.py` Excel output · `ab_testing.py` A/B analyzer · `restock_pricing.py` restock price evaluator · `cashflow.py` cash-flow restock plan · `basket_analysis.py` bundle/cross-sell · `deadstock_analysis.py` modal beku · `momentum_analysis.py` momentum+ABC · `trend_analysis.py` tren+musiman.
- `data/` input (gitignored), `output/` reports.

## CLI (`python main.py`)
- (no flag) = **full suite** (`--all`, 9 steps, loads data once via `_load_shared`, reorder once) · `--sales [YEAR]` · `--trend` · `--reorder` · `--cashflow` · `--bundle` · `--deadstock` · `--momentum` · `--ab-test` · `--restock-check` · `--data-dir`/`--output-dir`. Zero-config reports (trend/cash-flow/bundle/dead-stock/momentum) always run; ab-test/restock-check run only if their template has rows.

## Inputs (`data/`, by glob)
- `*BisaStok*.xlsx` (purchases, sheet `BisaStok`; latest also needs `BisaHilang`+`BisaPindahBarang`) and `*BisaJual*.xlsx` (sales, ≥ `BisaJualShopee`).
- BisaStok **`Toko`** = standardized supplier/forwarder (Ocistok/Martkita, AliExpress, Jasa Impor, marketplaces, local distributors); payment account + invoice/resi in **`Keterangan Pembelian`**; **`Luar Negeri?`** = authoritative overseas flag (splits local vs cross-border import). Loader auto-accepts the legacy `Toko[spasi]Akun Pemesan` header.

## Output: `Analisa_Penjualan_ITBisa_<year>.xlsx` — 9 sheets, pure sales history
`00_Summary`…`08_Supplier_Analysis` (Paling_Diminati, Profit_Tertinggi, Barang_Rugi, Margin_Borderline, Kandidat_Naik_Harga, Per_Platform, Data_Lengkap_per_SKU, Supplier_Analysis). Reorder + per-gudang stock = current snapshot (not year-specific) → only in `Analisa_Reorder.xlsx` (`--reorder`), not duplicated into yearly files.

## Core logic (do not regress)
- **Sisa stok = current-workbook ledger** matching Google Sheets `BisaRekapBarang`: per (SKU, gudang) `arrived beli − nonvoid jual + ketemu − hilang ± pindah`, from the latest stok/jual file by filename. Migrasi = opening balance (kept). Negative gudang floors to 0 (deficit shifted); OVERSOLD (negative total) stays negative + flagged.
- **HPP = weighted average, Ocistok-priority**: if a SKU has Ocistok/Martkita China-direct buys, average those only, else all. Migrasi dropped when real purchases exist.
- **Pricing-decision basis** (markup % & price recs): `markup_pct = (harga_sekarang − hpp_pricing)/hpp_pricing`; `hpp_pricing` = latest `Luar Negeri?=1` lot (else `hpp_wa`); `harga_sekarang` = lowest non-CoD price on the latest selling day (else `harga_jual_avg`). **Profit/margin keep `hpp_wa`+`harga_jual_avg`** (P&L).
- **Recent-price-increase guard (Kandidat Naik Harga)**: a freshly-raised, under-validated `harga_sekarang` is **held** (`Harga +%`/`Proyeksi Profit` blanked, Saran → `⏳ Harga baru naik …`) — qty/profit were earned at the OLD price. Change date from `ab_tests.xlsx` else auto step-detection; flags only when ≥`PRICE_CHANGE_MIN_STEP` over old price AND post-change qty `< PRICE_CHANGE_VALIDATION_MIN_SHARE` of the year.
- **Reorder**: velocity = avg monthly qty (total ÷ N months, trailing 6mo / fallback 12→24mo), CV → safety multiplier; **lead time is per-SHOP** from observed `Tanggal Bayar`→`Tanggal Sampai` at p75 (Ocistok=Martkita=1688 = one forwarder ≈2.5mo vs market ≈1mo). A SKU takes the **slowest shop supplying ≥`LEAD_SHOP_MIN_SHARE` of its qty** (import floored at global-import lead); local → `LEAD_TIME_MARKET_MONTHS`. Then ROP + order qty; buckets STOCKOUT/URGENT/Now/Soon/Overstock + Slow/Dead.
- **Supplier classification**: China = `Luar Negeri?=1` or `Toko`∈`CHINA_KEYWORDS`; Market = `MARKET_KEYWORDS`; else Other.
- **Restock check (`--restock-check`)**: landed HPP from `Harga RMB` (calibrated factor ≈Rp`RMB_TO_IDR_FALLBACK`/RMB) or given HPP IDR; verdict vs `hpp_wa`. Per-marketplace sell price = HPP×(1+`RESTOCK_TARGET_NET_MARKUP`)/(1−fee); decision vs competitor range 🟢/🟡/🔴. Input `restock_check.xlsx` → `Analisa_Restock_Check.xlsx`.
- **Cash-flow (`--cashflow`)**: reorder metrics → purchasing-budget calendar; inv-position sim plans **every** cycle in `CASHFLOW_HORIZON_MONTHS` (cost qty×`hpp_pricing`, by supplier×month). → `Analisa_Cashflow_Restock.xlsx`.
- **Bundle (`--bundle`)**: market basket per `Invoice` — support/confidence/lift, pairs ≥ `BASKET_MIN_PAIR_SUPPORT`. → `Analisa_Bundle_CrossSell.xlsx`.
- **Dead-stock (`--deadstock`)**: capital frozen in `🔵 Overstock`+`💤 Slow/Dead`; held = sisa×hpp_wa, freeable = max(0, sisa−`target_qty_post_reorder`)×hpp_wa; aksi 🧹 likuidasi / 🏷️ markdown / ⛔ stop-reorder. → `Analisa_Modal_Beku.xlsx`.
- **Trend (`--trend`)**: cross-year omzet/profit trend + YoY growth + per-month seasonal index (omzet bln ÷ rata2 bulanan tahunnya, tahun penuh saja; >1 = puncak); headline YTD vs tahun lalu. → `Analisa_Tren_Musiman.xlsx`.
- **Momentum + ABC (`--momentum`)**: momentum (recent vs prior qty: 🚀/📉 ±`MOMENTUM_GROWTH_THRESHOLD`, ➡️/🆕/💤) × ABC Pareto-by-trailing-profit (A ≤`ABC_A_SHARE`, B ≤`ABC_B_SHARE`); rek. dorong/lindungi/pangkas. → `Analisa_Momentum_ABC.xlsx`.
- **Summary (00)** surfaces a partial/current-year flag + a data-quality block (OVERSOLD + sold-without-HPP).
- **SKU normalization** `UPPER().strip()` everywhere. **No dedup** (only drop-Migrasi).

## Conventions
- All constants in `config.py` (never hardcode). Console strings Bahasa Indonesia. Minimal changes.

## Workflow (process standard)
- Branch `feature/<desc>` off `main`; doc/marker updates ride in the **same PR**. PR → **merge commit (`--no-ff`)**, title ends with the PR number. Authored **`C - Furqon Aji Yudhistira <furqonajiy@gmail.com>`**; **no AI/assistant references anywhere** (no "Claude"/"Anthropic"/`Co-Authored-By`/"Generated with"/session links); strip the auto-appended PR footer.
- CLI handed to the user = **PowerShell**. Sync marker `YYYY-MM-DD_HHMM.txt` (WIB) at root → rename on every update.

## Flag before changing
The current-workbook stock-ledger reconciliation (parity with `BisaRekapBarang`), the Ocistok-priority HPP rule, the pricing-decision basis + recent-price guard (`PRICE_CHANGE_*`), the per-shop reorder lead time + `config.py` tunables, the standardized `Toko` contract (supplier + lead keying), SKU `UPPER().strip()` normalization, the removed dedup / drop-Migrasi behavior, the 9-sheet yearly layout (reorder/rekap → own file), and the input glob patterns.
