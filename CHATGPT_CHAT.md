# ChatGPT Chat — Project Instructions (itbisa-shop-report-bot)

> **ChatGPT Chat** project instructions — a ≤8000-char condensed brief (ChatGPT's limit). For **Claude** (Code *and* Chat) use the full **`CLAUDE.md`** directly; **`AGENTS.md`** (ChatGPT Codex) points to `CLAUDE.md`. Update only when explicitly requested.

## What this is
Standalone, **offline** Python tool that turns ITBisa sales/stock Excel exports into analysis workbooks (sales performance, pricing, supplier, reorder, cash-flow, channel, bundle + A/B price-test + restock-check reports). No API/network/tokens/GitHub Actions — runs locally, idempotent.

## Stack & files (flat layout, no `src/`)
- Python 3.10+. Deps: `pandas`, `openpyxl`.
- `main.py` CLI/orchestration · `config.py` all constants · `data_loader.py` loading + SKU normalization + current-workbook loaders · `analysis.py` HPP, profit, aggregation, supplier classification, reorder, `build_stock_ledger`, `compute_lead_time_months`, `compute_price_change_status` · `tables.py` table builders · `excel_writer.py` Excel output · `ab_testing.py` A/B analyzer · `restock_pricing.py` restock price evaluator · `cashflow.py` cash-flow restock plan · `channel_analysis.py` per-SKU channel optimizer · `basket_analysis.py` bundle/cross-sell.
- `data/` input (gitignored), `output/` reports.

## CLI (`python main.py`)
- (no flag) = **full suite** (`--all`, 7 steps, loads data once via `_load_shared`, reorder once) · `--sales [YEAR]` all/one year · `--reorder` · `--cashflow` · `--channel` · `--bundle` · `--ab-test` · `--restock-check` · `--data-dir`/`--output-dir`. Cash-flow/channel/bundle are zero-config (always run); ab-test/restock-check run only if their template has rows.

## Inputs (`data/`, by glob)
- `*BisaStok*.xlsx` (purchases, sheet `BisaStok`; latest also needs `BisaHilang`+`BisaPindahBarang`) and `*BisaJual*.xlsx` (sales, ≥ `BisaJualShopee`).
- BisaStok **`Toko`** = standardized supplier/forwarder (Ocistok/Martkita, AliExpress, Jasa Impor Tiongkok, Osell, Shopee, Tokopedia, local distributors); payment account + invoice/resi in **`Keterangan Pembelian`**; **`Luar Negeri?`** = authoritative overseas flag (splits e.g. Shopee into local vs cross-border import). Loader auto-accepts the legacy `Toko[spasi]Akun Pemesan` header.

## Output: `Analisa_Penjualan_ITBisa_<year>.xlsx` (12 sheets)
`00_Summary`, `01_Paling_Diminati`, `02_Profit_Tertinggi`, `03_Barang_Rugi`, `04_Margin_Borderline`, `05_Kandidat_Naik_Harga`, `06_Per_Platform`, `07_Data_Lengkap_per_SKU`, `08_Supplier_Analysis`, `09_Reorder_Analysis`, `10_Reorder_Data_Lengkap`, `11_Rekap_Stok_per_Gudang`. `--reorder` emits its own workbook with `03_Rekap_Stok_per_Gudang`.

## Core logic (do not regress)
- **Sisa stok = current-workbook ledger** matching Google Sheets `BisaRekapBarang`: per (SKU, gudang) `arrived beli − nonvoid jual + ketemu − hilang ± pindah`, from the latest stok/jual file by filename. Migrasi = opening balance (kept). Negative gudang floors to 0 (deficit shifted); OVERSOLD (negative total) stays negative + flagged.
- **HPP = weighted average, Ocistok-priority**: if a SKU has Ocistok/Martkita China-direct buys, average those only, else all. Migrasi dropped when real purchases exist.
- **Pricing-decision basis** (markup %, Kandidat/Borderline/Rugi recs): `markup_pct = (harga_sekarang − hpp_pricing)/hpp_pricing`, `hpp_pricing` = latest `Luar Negeri?=1` lot (else `hpp_wa`), `harga_sekarang` = lowest non-CoD unit price on the SKU's most recent selling day (else `harga_jual_avg`). **Profit/margin keep `hpp_wa` + `harga_jual_avg`** (realized P&L).
- **Recent-price-increase guard (Kandidat Naik Harga)**: a freshly-raised, under-validated `harga_sekarang` is **held** — `Harga +%`/`Proyeksi Profit` blanked, Saran → `⏳ Harga baru naik …` — because its qty/profit were earned at the OLD price. Change date from `ab_tests.xlsx` (authoritative) else auto two-window step-detection; flags only when ≥`PRICE_CHANGE_MIN_STEP` over the old price AND post-change qty `< PRICE_CHANGE_VALIDATION_MIN_SHARE` of the year. See `compute_price_change_status`.
- **Reorder**: velocity = avg monthly qty (total ÷ N months, trailing 6mo / fallback 12→24mo), CV → safety multiplier; **lead time is per-SHOP** from observed `Tanggal Bayar`→`Tanggal Sampai` at p75 (Ocistok/Martkita ≈2.5mo vs AliExpress ≈1mo; Ocistok=Martkita=1688 are one forwarder). A SKU takes the **slowest shop supplying ≥`LEAD_SHOP_MIN_SHARE` of its qty** (import status from `Luar Negeri?`/China-keyword share; import floored at global-import); local → `LEAD_TIME_MARKET_MONTHS`. Then ROP, suggested order; buckets STOCKOUT/URGENT/Now/Soon/Overstock + Slow/Dead. See `compute_lead_time_months`.
- **Supplier classification** (standardized `Toko` + `Luar Negeri?`): China = `Luar Negeri?=1` or `Toko` in `CHINA_KEYWORDS`; Market = `MARKET_KEYWORDS` (Shopee/Tokopedia/Bukalapak/Blibli/Tiktok); else Other.
- **Restock check (`--restock-check`)**: input data/restock_check.xlsx (SKU, Toko, Harga RMB and/or HPP IDR, Kompetitor Min/Max). Landed HPP from RMB via a history-calibrated factor (≈Rp`RMB_TO_IDR_FALLBACK`/RMB, per-SKU when available) or given HPP IDR; verdict vs `hpp_wa`. Per-marketplace sell price = HPP×(1+`RESTOCK_TARGET_NET_MARKUP`)/(1−fee) so net ≥ target after the fee (fees from BisaJual). Decision vs competitor range: 🟢 sell / 🟡 thin / 🔴 don't. See `restock_pricing.py`.
- **Cash-flow (`--cashflow`)**: reorder metrics → purchasing-budget calendar. Inventory-position sim plans **every** reorder cycle in `CASHFLOW_HORIZON_MONTHS` (order up to `target+lead_demand` when position crosses `rop_final`, deplete at velocity; cap `CASHFLOW_MAX_CYCLES`); STOCKOUT/below-ROP due now. Cost = qty × `hpp_pricing` (fallback `hpp_wa`); supplier = dominant `Toko`. Output `Analisa_Cashflow_Restock.xlsx`. Zero-config.
- **Channel optimizer (`--channel`)**: per SKU, net margin/unit per marketplace = `(omzet+admin)/qty − hpp_wa` (admin negative); dominant-volume channel vs best established channel (qty ≥ `CHANNEL_MIN_QTY`); flag 🔁 shift when gap ≥ `CHANNEL_SHIFT_MIN_GAP × hpp`/pcs. Output `Analisa_Channel_per_SKU.xlsx`. Zero-config.
- **Bundle/cross-sell (`--bundle`)**: market basket per `Invoice`; pair `support`/`confidence`/`lift`; report pairs support ≥ `BASKET_MIN_PAIR_SUPPORT` (top `BASKET_TOP_N`). Output `Analisa_Bundle_CrossSell.xlsx`. Zero-config.
- **Summary (00)** surfaces a partial/current-year flag + a data-quality block (OVERSOLD + sold-without-HPP).
- **SKU normalization** `UPPER().strip()` everywhere. **No dedup** (only drop-Migrasi).

## Conventions
- All constants in `config.py` — never hardcode thresholds/column names. Console strings Bahasa Indonesia. Minimal, targeted changes only.

## Workflow (process standard)
- Branch `feature/<desc>` off `main`; doc/marker updates ride in the **same PR as the code**. PR → **merge commit (`--no-ff`)**, title representative + ends with the PR number.
- Authored **`C - Furqon Aji Yudhistira <furqonajiy@gmail.com>`**; **no AI/assistant references anywhere** (no "Claude"/"Anthropic", no `Co-Authored-By`, no "Generated with", no session links). After opening a PR, strip the auto-appended footer.
- CLI handed to the user = **PowerShell** syntax. Branch deletion is blocked in the sandbox — hand the user a PowerShell `git push origin --delete …`.
- Sync marker `YYYY-MM-DD_HHMM.txt` (WIB) at root → rename to current WIB on every update.

## Flag before changing
The current-workbook stock-ledger reconciliation (parity with `BisaRekapBarang`), the Ocistok-priority HPP rule, the pricing-decision basis + recent-price guard (`PRICE_CHANGE_*`), the per-shop reorder lead time + `config.py` tunables, the standardized `Toko` contract (supplier + lead keying), SKU `UPPER().strip()` normalization, the removed dedup / drop-Migrasi behavior, the 12-sheet output layout, and the input glob patterns.
