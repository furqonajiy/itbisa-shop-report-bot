# CLAUDE.md — itbisa-bisalaporan

> **Source of truth for the Claude surfaces** — read by **Claude Code** and pasted into the **Claude Chat** project instructions. `AGENTS.md` (ChatGPT Codex) points here.

Standalone, **offline** Python tool that turns raw marketplace exports (Shopee, Tokopedia, Tiktok, Bukalapak) into standardized **BisaLaporan** workbooks (`BisaInvoice` / `BisaJual` / `BisaRemit` / `BisaBonus` + a combined `Final` sheet per period). No API calls, no network, no secrets, no GitHub Actions — runs locally and is idempotent. The generated **`BisaJual`** sheet is the upstream feed consumed by the sibling repo **itbisa-shop-report-bot**.

## Stack & files
- Python 3.8–3.11. Deps: `pandas` (**1.x**, pinned `>=1.4,<2.0`), `openpyxl` (`requirements.txt`).
  - **pandas must be < 2.0**: the Excel writers use the pandas-1.x `ExcelWriter.book` setter (removed in 2.0). `utility/generic.py` imports `SettingWithCopyWarning` from `pandas.errors` (pandas ≥1.5) with a fallback to `pandas.core.common` (older).
- `main.py` (repo root) — **CLI entry point** (`argparse`). Puts `generator/` on `sys.path`, parses flags, sets the data/report dirs, then calls `generator.run(...)`.
- `generator/main.py` — orchestration: `MARKETPLACE_PROCESSORS` (marketplace → ordered processor modules) and `run(list_report, marketplaces=None)`. After a marketplace's processors finish, `run` calls `generate_final(<Marketplace>)` so the `Final` sheet sees every period's workbook.
- `generator/process/preprocess.py` — `generate_report_list`: **recursive** glob of `data/` (`**/*.xls*`, `**/*.csv`).
- `generator/process/<marketplace>/<vN>.py` — per-marketplace/version readers: drop invalid-status rows, validate keywords, dispatch to the sheet generators.
- `generator/bisainvoice|bisajual|bisaremit|bisabonus|bisafee/` — each has a `generic.py` with the shared `*_to_excel` writer plus per-marketplace/version builders.
- `generator/bisafinal/generic.py` — `generate_final(marketplace)`: builds the per-workbook `Final` sheet by **reading back** the already-written `BisaInvoice`/`BisaJual`/`BisaRemit` sheets (marketplace-agnostic, no per-marketplace builders).
- `generator/bisarekonsiliasi/generic.py` — `generate_reconciliation(marketplaces)`: **read-only** `--reconcile` audit. Re-reads the raw `BisaSaldo`/`BisaFee` inputs, classifies every saldo row into Remit/Bonus/Penarikan/**Tidak Tercatat** using the `keywordchecker/` lists (single source of truth), and writes `reports/<marketplace>/Rekonsiliasi <Marketplace>.xlsx`. **Never changes any generated number.**
- `generator/keywordchecker/` — validates marketplace status / saldo keywords (raises on unknown keyword).
- `generator/utility/constant.py` — `DEFAULT_DATA_DIR` / `DEFAULT_REPORTS_DIR` (repo-relative), `set_dirs`/`get_data_dir`/`get_reports_dir`, and `MARKETPLACE_FOLDERS` (marketplace name → reports subfolder).
- `generator/utility/generic.py` — `ignore_warning`, `create_directory`, `detect_marketplace`, **`build_report_path`**.
- `generator/utility/sku.py` — `standardize_sku` (canonical SKU rewrites).
- `data/` — input exports (gitignored). `reports/` — generated output (gitignored).

## CLI (`python main.py`)
- (no flag) → process **every** marketplace (Bukalapak, Tokopedia, Shopee, Tiktok — the `MARKETPLACE_PROCESSORS` order).
- `--shopee` / `--tokopedia` / `--tiktok` / `--bukalapak` → process only the selected marketplace(s); flags combine.
- `--data-dir <dir>` / `--output-dir <dir>` → override `./data` / `./reports`.
- `--show-files` → log every discovered input file. `-v` / `--verbose` → debug logging.
- `--reconcile` → **read-only audit mode**: writes `Rekonsiliasi <Marketplace>.xlsx` (does **not** generate reports). Respects the marketplace + `--data-dir`/`--output-dir` flags.
- With no input files found, the CLI logs `Tidak ada file ditemukan …` and exits cleanly.

## Inputs (`data/`, matched recursively)
- Files may sit directly in `data/` or in any subfolder — discovery is recursive, and each file is classified by the **marketplace + version + type token in its filename** (not by folder), e.g. `BisaTransaksi v2 Shopee`, `BisaSaldo v2 Shopee`, `BisaFee v1 Tiktok`. Temp lock files (`~…`) are skipped by the per-processor `cond` checks.
- Three input types: **`BisaTransaksi`** (orders → BisaInvoice + BisaJual), **`BisaSaldo`** (balance mutations → BisaRemit + BisaBonus), **`BisaFee`** (settlement fees → feeds Shopee BisaRemit math; is the BisaRemit source for **Tiktok**).

## Output: `reports/<marketplace>/<name> BisaLaporan <Marketplace>.xlsx`
- `<marketplace>` ∈ `shopee`, `tiktokshop`, `tokopedia`, `bukalapak` (lowercase; Tiktok → `tiktokshop`). See `MARKETPLACE_FOLDERS`.
- Sheets per workbook: `BisaInvoice <MP>`, `BisaJual <MP>` (from `BisaTransaksi`); `BisaRemit <MP>` (from `BisaSaldo`, or `BisaFee` for Tiktok); `BisaBonus <MP>` (from `BisaSaldo`); plus a `Final` sheet (every workbook). Coverage: Shopee v2/v3 = all four; Tokopedia v1 = invoice + jual; Tokopedia v2 = all four; Tiktok v1 = invoice + jual + remit; Bukalapak v2 = invoice + jual + remit. `Final` is always present.

## Core logic (do not regress)
- **Path routing (`build_report_path`)**: each sheet generator still computes its report **filename** the original way — string-replace the input basename's type token (`BisaTransaksi`/`BisaSaldo`/`BisaFee` → `BisaLaporan`), strip ` v1/ v2/ v3` (Bukalapak keeps its ` v2`), `.csv`→`.xlsx`. `build_report_path` then **re-roots that filename** into `reports/<marketplace>/` (detected from the filename) and creates the folder. The directory part of the old string-replace result is discarded; only the basename is kept.
- **Workbook convergence (append model)**: for one period the `BisaTransaksi` and `BisaSaldo`/`BisaFee` files collapse to the **same** `… BisaLaporan …` basename, so all sheets land in one workbook. `BisaInvoice` is written first in **write** mode (creates the file); `BisaJual`/`BisaRemit`/`BisaBonus` **append** (`mode='a', if_sheet_exists='replace'`, with a create-on-missing fallback). The per-marketplace `process()` runs the `BisaTransaksi` pass before the `BisaSaldo`/`BisaFee` pass so the file exists when the appenders run. **Do not break this filename convergence or the write→append order.**
- **Recursive discovery**: `generate_report_list` returns `sorted(set(...))` of everything under `data/`; the per-processor `cond1 = '<token>' in file` filters select the right files. This is OS-independent (no Windows-only `H:\` paths or backslash globs).
- **Marketplace detection** (`detect_marketplace`): substring match of the report filename against `MARKETPLACE_FOLDERS` keys (`Shopee`/`Tiktok`/`Tokopedia`/`Bukalapak`); raises if none match.
- **Configurable dirs**: consumers call `get_data_dir()` / `get_reports_dir()` at run time (never bind the value at import) so the CLI's `set_dirs()` override (and `--data-dir`/`--output-dir`) takes effect.
- **Reconciliation (`bisarekonsiliasi/generic.py`, `--reconcile`)**: a **read-only audit** — it never edits the generated sheets. It mirrors each marketplace's `BisaSaldo` read (skiprows/cols) and the **per-marketplace remit row filter** (Shopee v2 `#`/`Penambahan dana`, Shopee v3 `No. Pesanan` without `-`, Tokopedia `INV`, Bukalapak `#`), then buckets every row via the imported `keywordchecker/` lists: a row is **Tidak Tercatat** when it matches a remit keyword but the invoice filter excludes it (e.g. `Pencairan SPinjam untuk Penjual`), is an invoice row with no remit keyword, or is a bonus keyword where no `BisaBonus` is generated (Bukalapak has none). `not_used` is derived as `VALID_SALDO_KEYWORD − components − bonus` (stays in sync). Shopee fee check de-duplicates overlapping fee files (a yearly v2 + monthly v3 file list the same order) before matching `(Invoice, Total Penghasilan)` to the remit. Besides `Ringkasan`/`Saldo Tidak Tercatat`/`BisaFee Tidak Cocok`, the workbook has a **`Rincian per Deskripsi`** sheet (every row rolled up to its matched keyword → category/bucket/count/total/example, uncaptured groups flagged), a **`Rincian Saldo`** row-level detail, and (Shopee) a **`Cek Remit Saldo vs Fee`** sheet listing every invoice's `BisaSaldo` remit vs `BisaFee` `Total Penghasilan` side by side with a Cocok/Beda status, plus a **`Cek Omzet vs Fee`** sheet that pairs re-derived `BisaJual` Omzet with the real money (`BisaSaldo` net + `BisaFee` `Total Penghasilan`/`Kerugian`) and flags orders whose Omzet isn't real money — *Retur - rugi = omzet* (return loss only in `BisaFee`, net 0, e.g. `231215A1CX1JNX`), *Omzet tidak settle*, *Belum ada penghasilan* (pending, not a loss). Keep this in lockstep with the generators' filters/keywords.
- **`Final` sheet (`bisafinal/generic.py`)**: one reconciliation row per `Invoice`, joining the order side with the remit side. `Omzet Barang` = Σ `BisaJual` `Omzet` per `Invoice`; `Nominal Invoice` = `Omzet Barang` + `Ongkir` + `Asuransi`. It is a **left join from `BisaInvoice`** (every order is a row), and the remit is looked up against the **union of every `BisaLaporan` workbook in `reports/<marketplace>/`** — so an order placed one period and remitted the next still matches (multi-remit → amounts summed, latest `Tanggal Remit` kept); when no remit is found the remit columns stay blank. `Cek Remit`/`Untung Lainnya`/`Rugi Lainnya`/`Keterangan` are always blank (manual). Because of the cross-period lookup, `generate_final` runs **after** all of a marketplace's processors (so every period's `BisaRemit` exists), reads the written sheets back (money is stored as text → coerced), and appends/replaces the `Final` sheet (`if_sheet_exists='replace'`, create-on-missing fallback). **Do not move `Final` before the remit pass or break the cross-period union.**

## Conventions
- Constants/paths live in `utility/constant.py` — never hardcode `data/`, `reports/`, or marketplace folders in logic.
- User-facing console strings are **Bahasa Indonesia**; **markdown docs are English prose** keeping domain terms verbatim (`BisaTransaksi`, `BisaJual`, `Toko`, `omzet`, marketplace/version tokens, any quoted console string).
- **The user runs Windows PowerShell** — give CLI snippets in PowerShell syntax.
- Minimal, targeted changes only; preserve existing behavior unless explicitly in scope.

## Development workflow (process standard)
- Branch from `main` using `feature/<short-description>` (e.g. `feature/repo-relative-paths-and-cli`).
- Always open a PR into `main` and **merge with a merge commit (`--no-ff`)** — never squash, never fast-forward — so the feature branch stays an ancestor of `main`.
- The merge commit title must be **representative and end with the PR number**, e.g. `Generate reports under the repo and add a python main.py CLI (#1)` — never the default `Merge pull request …`.
- Commits and PRs are authored as **`C - Furqon Aji Yudhistira <furqonajiy@gmail.com>`** — at the start of every session set `git config user.name "C - Furqon Aji Yudhistira"` and `git config user.email "furqonajiy@gmail.com"`.
- **No AI/assistant references anywhere** in branch names, commit messages, PR titles/bodies, or code/comments: no "Claude"/"Anthropic", no `Co-Authored-By` trailer, no "Generated with …" line, no model identifiers, no session links. History must read as human-authored.
- **Strip the auto-appended PR footer**: the PR-creation tool appends a `_Generated by …_` line to the PR **body** — immediately edit the description to remove it, then re-read to confirm. Avoid raw `<…>` placeholders in PR/issue bodies (GitHub strips them as HTML); use `[tgl]`-style instead.
- **Branch deletion from this environment may be blocked** (sandbox git proxy) — after a merge, hand the user a PowerShell `git push origin --delete <branch>` command so they can delete the merged feature branch.
- Keep `README.md` / `CLAUDE.md` / `RUNBOOK.md` updated in the **same PR** whenever behavior or process changes.

## Flag before changing
The recursive `data/` discovery glob and the per-processor filename-token classification, the `build_report_path` re-rooting + the workbook **filename convergence and write→append order**, the `MARKETPLACE_FOLDERS` mapping (incl. Tiktok → `tiktokshop`), the input filename token conventions, the **`Final` sheet's cross-period remit union + post-processor ordering** (and its column formulas: `Omzet Barang` = Σ `BisaJual` `Omzet`, `Nominal Invoice` = `Omzet Barang` + `Ongkir` + `Asuransi`, blank `Cek Remit`/`Lainnya`/`Keterangan`), and the **pandas < 2.0** constraint (`ExcelWriter.book` + `SettingWithCopyWarning` import).
