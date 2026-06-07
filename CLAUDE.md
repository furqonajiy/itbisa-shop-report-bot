# CLAUDE.md — itbisa-shop-report-bot

Standalone, **offline** Python tool that turns ITBisa sales/stock Excel exports into an analysis workbook (sales performance, pricing, supplier, and reorder analysis + A/B price-test reports). No API calls, no network, no tokens, no GitHub Actions — runs locally and is idempotent.

> Unlike the sibling ITBisa repos, this one is **not** a GitHub-Actions bot. It does not call Shopee/TikTok Shop APIs, has no `bot-state` branch, no workflows, no Telegram, and no secrets.

## Stack & files (flat layout, no `src/`)
- Python 3.10+. Deps: `pandas`, `openpyxl` (`requirements.txt`).
- `main.py` — CLI entry point / orchestration.
- `config.py` — all constants: globs, sheet/column names, thresholds, colors, supplier keywords, reorder + A/B params.
- `data_loader.py` — multi-file glob loading, SKU normalization, current-workbook loaders (arrived beli / jual / hilang / pindah).
- `analysis.py` — HPP weighted-average, profit, per-SKU aggregation, supplier classification, reorder metrics, `build_stock_ledger`.
- `tables.py` — table builders (diminati, profit, rugi, kandidat, supplier, reorder).
- `excel_writer.py` — Excel output, incl. the per-gudang stock-reconciliation sheet.
- `ab_testing.py` — A/B price-test analyzer + template creation.
- `data/` — input Excel (gitignored). `output/` — generated reports.

## CLI (`python main.py`)
- (no flag) → sales analysis for the current year.
- `--sales` → all years (one file per year); `--sales YEAR` → specific year.
- `--reorder` → standalone `Analisa_Reorder.xlsx`.
- `--ab-test` → `Analisa_AB_Test.xlsx` (auto-creates `data/ab_tests.xlsx` template if missing).
- `--all` → all years + reorder + ab-test.
- `--data-dir` / `--output-dir` override the defaults.

## Inputs (`data/`, matched by glob)
- `STOK_GLOB = "*BisaStok*.xlsx"` (purchases) and `JUAL_GLOB = "*BisaJual*.xlsx"` (sales). Filenames must contain `BisaStok` / `BisaJual`.
- Stok files need sheet `BisaStok`; the latest file also needs `BisaHilang` and `BisaPindahBarang` for stock reconciliation.
- Jual files need at least `BisaJualShopee`; other `BisaJual*` sheets load if present.

## Output: `output/Analisa_Penjualan_ITBisa_<year>.xlsx`
12 sheets: `00_Summary`, `01_Paling_Diminati`, `02_Profit_Tertinggi`, `03_Barang_Rugi`, `04_Margin_Borderline`, `05_Kandidat_Naik_Harga`, `06_Per_Platform`, `07_Data_Lengkap_per_SKU`, `08_Supplier_Analysis`, `09_Reorder_Analysis`, `10_Reorder_Data_Lengkap`, `11_Rekap_Stok_per_Gudang`. The reorder report (`--reorder`) emits its own workbook with a `03_Rekap_Stok_per_Gudang` sheet.

## Core logic (do not regress — README has the full spec)
- **Sisa stok = current-workbook ledger**, recomputed to match the Google Sheets `BisaRekapBarang`: per (SKU, gudang) `arrived beli − nonvoid jual + ketemu − hilang ± pindah`. "Current workbook" = the latest `*BisaStok*`/`*BisaJual*` file by filename sort. Migrasi rows are the opening balance (kept here). Negative gudang balances floor to 0 and shift the deficit; truly OVERSOLD SKUs stay negative and are flagged to console.
- **HPP = weighted average with Ocistok-priority**: if a SKU has Ocistok/Martkita (China-direct) purchases, average those only; else average all. Combined across all stok files. For HPP/total-beli, Migrasi rows are dropped when real purchases exist (avoid double-count).
- **SKU normalization**: `UPPER().strip()` on load everywhere (matches case-insensitive `SUMIF`).
- **No dedup**: the old `drop_duplicates` was removed (it discarded genuine duplicate purchase lots); only drop-Migrasi remains.
- **Reorder**: velocity (6mo window, fallback 12→24mo), volatility CV → safety multiplier, lead time (China 2mo / market 0.25mo), ROP, suggested order; action buckets STOCKOUT/URGENT/Now/Soon/Overstock + Slow/Dead. Tunables in `config.py`.
- **Supplier classification**: China = `Luar Negeri? = 1` or supplier name in `CHINA_KEYWORDS`; Market = `MARKET_KEYWORDS`; else Other.

## Conventions
- All constants live in `config.py` — never hardcode thresholds/column names in logic.
- User-facing console strings are Bahasa Indonesia.
- Minimal, targeted changes only; preserve existing behavior unless explicitly in scope.

## Development workflow (process standard)
- Branch from `main` using `feature/<short-description>` (e.g. `feature/document-dev-workflow`).
- Always open a PR into `main` and **merge with a merge commit (`--no-ff`)** — never squash, never fast-forward — so the feature branch stays an ancestor of `main`.
- The merge commit title must be **representative and end with the PR number**, e.g. `Update Project Instructions to the Latest State (#47)` — never the default `Merge pull request …`.
- Commits and PRs are authored as **`C - Furqon Aji Yudhistira <furqonajiy@gmail.com>`** (never "Claude").
- Keep changes minimal and targeted; update `CLAUDE.md` / `README.md` in the same PR whenever behavior or process changes.
- `PROJECT_INSTRUCTIONS.md` is the synced source for the Claude & ChatGPT project instructions (≤ 8000 chars, ChatGPT limit). Update it **only when explicitly asked**, not on every change.
- Sync marker: a file named `YYYY-MM-DD_HHMM.txt` (WIB) sits at the repo root. **On every update to this repo, rename it to the current WIB timestamp** — it signals whether the repo / Claude / ChatGPT instructions are in sync.
- Doc/marker updates (this file, `PROJECT_INSTRUCTIONS.md`, the sync marker) ride in the **same feature branch and PR as the related code change** — never a separate doc-only branch (avoids noise).

## Flag before changing
The current-workbook stock-ledger reconciliation (and its parity with `BisaRekapBarang`), the Ocistok-priority HPP rule, SKU `UPPER().strip()` normalization, the removed dedup / drop-Migrasi behavior, the reorder methodology and its `config.py` tunables, the 12-sheet output layout, and the input glob patterns.
