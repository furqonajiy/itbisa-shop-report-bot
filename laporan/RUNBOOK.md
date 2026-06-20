# RUNBOOK — laporan (Laporan generator)

Operating procedure for generating the **Laporan** workbooks. Commands are in
**Windows PowerShell** (the supported environment). See [`README.md`](README.md) for
the conceptual overview. This generator is the importable `laporan` package inside
**itbisa-shop-report-bot**; run it from the repo root via `python main.py --laporan`
(called in-process), or standalone with `python -m laporan`.

## 1. One-time setup

```powershell
cd <path-to>\itbisa-shop-report-bot
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

> Use a Python **3.13** interpreter. The repo-root `requirements.txt` pins
> `pandas>=2.0,<3.0` (pandas **3.0** is not yet supported — its new `str` dtype +
> Copy-on-Write default need a separate port).

## 2. Drop the marketplace exports into `data/`

Put the raw exports anywhere under `data\` (directly, or in subfolders — discovery is
recursive). Files are identified by the **token in their filename**, so keep the
exporter's naming, e.g.:

```
data\
  ... Transaksi v2 Shopee.xlsx
  ... Saldo v2 Shopee.csv
  ... Fee v2 Shopee.xlsx
  ... Transaksi v3 Shopee.xlsx
  ... Transaksi v1 Tokopedia.xlsx
  ... Transaksi v2 Tokopedia.xlsx
  ... Saldo v2 Tokopedia.xlsx
  ... Transaksi v1 Tiktok.xlsx
  ... Fee v1 Tiktok.xlsx
  ... Transaksi v2 Bukalapak.xlsx
  ... Saldo v2 Bukalapak.csv
```

`data\` is git-ignored — your exports are never committed.

## 3. Generate the reports

Run these from the **repo root** (the `laporan` package is invoked as a module). The
generator's own CLI (`--shopee`, `--reconcile`, …) is the same whether you call
`python -m laporan` or the bot's `python main.py --laporan`.

```powershell
# everything
python -m laporan

# a single marketplace (flags combine)
python -m laporan --shopee
python -m laporan --shopee --tiktok

# preview which input files were found, with debug logs
python -m laporan --show-files -v

# custom folders
python -m laporan --data-dir D:\exports --output-dir D:\reports
```

## 4. Collect the output

Reports are written to `reports\<marketplace>\`:

```
reports\
  shopee\      ... Laporan Shopee.xlsx
  tiktokshop\  ... Laporan Tiktok.xlsx
  tokopedia\   ... Laporan Tokopedia.xlsx
  bukalapak\   ... Laporan v2 Bukalapak.xlsx
```

Each workbook contains the `Invoice`, `Jual`, `Remit`, and/or `Bonus`
sheets for that period, plus a combined **`Final`** sheet (one reconciliation row per
`Invoice`; see the coverage table and the `Final` section in `README.md`). The run is
**idempotent** — re-running overwrites the matching sheets in place.

> The `Final` sheet looks up each order's remit across **all** of that marketplace's
> `Laporan` workbooks, so process a marketplace's periods together (an order placed
> one month and remitted the next only fills in once both periods' files are present).

## 4b. (Optional) Reconcile Saldo / Fee

```powershell
# audit only — writes Rekonsiliasi <Marketplace>.xlsx, generates no reports
python -m laporan --reconcile
python -m laporan --reconcile --shopee     # one marketplace

# use the itbisa-shop-report-bot Jual ledger for the Cek Omzet vs Fee sheet
python -m laporan --reconcile --shopee --jual-dir "..\itbisa-shop-report-bot\data"
```

This **read-only** pass re-reads the raw `Saldo` / `Fee` files and reports any
balance movement that is **not** captured into `Remit`/`Bonus`. Open the
`Rekonsiliasi <Marketplace>.xlsx` and check the **Ringkasan** tab for red **Perlu
Dicek** rows. **Rincian per Deskripsi** is the review list — every distinct description
(by matched keyword) with its category, bucket, count and total — so you can decide how
each description should be treated; **Rincian Saldo** is the full row-level detail. The
**Saldo Tidak Tercatat** tab lists each uncaptured row and why (e.g. a `Pencairan
SPinjam untuk Penjual` loan excluded by the invoice filter), and for Shopee **Fee
Tidak Cocok** lists fee/remit amounts that don't reconcile. Nothing here changes your
`Laporan` numbers — it's a checklist of things to review.

## 5. Hand the `Jual` to itbisa-shop-report-bot

The **`Jual <Marketplace>`** sheet is the upstream feed for the sibling project
[`itbisa-shop-report-bot`](https://github.com/furqonajiy/itbisa-shop-report-bot). Copy
the relevant workbook(s) into that project's `data\` per its own README.

## Troubleshooting

| Symptom | Cause / fix |
| --- | --- |
| `Tidak ada file ditemukan di …` | No `*.xls*` / `*.csv` under the data dir. Check `--data-dir` and that files actually landed in `data\`. Run `python -m laporan --show-files` to list what was discovered. |
| A file is ignored | Its filename is missing the marketplace/version/type token (e.g. `Transaksi v2 Shopee`). Rename to the exporter's convention. |
| `ValueError: Tidak dapat menentukan marketplace dari file: …` | The report filename has no `Shopee`/`Tiktok`/`Tokopedia`/`Bukalapak` substring — fix the source filename. |
| `ImportError` / `numpy.dtype size changed` on `import pandas` | pandas/numpy mismatch. Reinstall the pinned set: `pip install -r requirements.txt`. |
| `TypeError: Invalid value '0' for dtype 'str'` / `fillna() got an unexpected keyword 'method'` | pandas **3.0** is installed (unsupported). Pin to 2.x: `pip install "pandas>=2.0,<3.0"`. |
| `ValueError: Check … Keyword failed in …` | The export contains a status/saldo keyword the validator doesn't know (`keywordchecker/`). A new marketplace wording appeared — add it to the relevant `VALID_*` list. |

## Verifying a change before committing

```powershell
# compile-check every module in the package
python -m py_compile (Get-ChildItem -Recurse laporan\*.py | % FullName)

# smoke-run against a scratch folder (should report "Tidak ada file ditemukan")
python -m laporan --data-dir .\_empty --output-dir .\_out
```
