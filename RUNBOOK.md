# RUNBOOK — itbisa-bisalaporan

Operating procedure for generating the **BisaLaporan** workbooks. Commands are in
**Windows PowerShell** (the supported environment). See [`README.md`](README.md) for
the conceptual overview.

## 1. One-time setup

```powershell
cd <path-to>\itbisa-bisalaporan
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

> Use a Python **3.8–3.11** interpreter. `requirements.txt` pins `pandas>=1.4,<2.0`
> (the Excel writers rely on the pandas-1.x `ExcelWriter.book` API); pandas 1.x does
> not support Python 3.12+.

## 2. Drop the marketplace exports into `data/`

Put the raw exports anywhere under `data\` (directly, or in subfolders — discovery is
recursive). Files are identified by the **token in their filename**, so keep the
exporter's naming, e.g.:

```
data\
  ... BisaTransaksi v2 Shopee.xlsx
  ... BisaSaldo v2 Shopee.csv
  ... BisaFee v2 Shopee.xlsx
  ... BisaTransaksi v3 Shopee.xlsx
  ... BisaTransaksi v1 Tokopedia.xlsx
  ... BisaTransaksi v2 Tokopedia.xlsx
  ... BisaSaldo v2 Tokopedia.xlsx
  ... BisaTransaksi v1 Tiktok.xlsx
  ... BisaFee v1 Tiktok.xlsx
  ... BisaTransaksi v2 Bukalapak.xlsx
  ... BisaSaldo v2 Bukalapak.csv
```

`data\` is git-ignored — your exports are never committed.

## 3. Generate the reports

```powershell
# everything
python main.py

# a single marketplace (flags combine)
python main.py --shopee
python main.py --shopee --tiktok

# preview which input files were found, with debug logs
python main.py --show-files -v

# custom folders
python main.py --data-dir D:\exports --output-dir D:\reports
```

## 4. Collect the output

Reports are written to `reports\<marketplace>\`:

```
reports\
  shopee\      ... BisaLaporan Shopee.xlsx
  tiktokshop\  ... BisaLaporan Tiktok.xlsx
  tokopedia\   ... BisaLaporan Tokopedia.xlsx
  bukalapak\   ... BisaLaporan v2 Bukalapak.xlsx
```

Each workbook contains the `BisaInvoice`, `BisaJual`, `BisaRemit`, and/or `BisaBonus`
sheets for that period (see the coverage table in `README.md`). The run is
**idempotent** — re-running overwrites the matching sheets in place.

## 5. Hand the `BisaJual` to itbisa-shop-report-bot

The **`BisaJual <Marketplace>`** sheet is the upstream feed for the sibling project
[`itbisa-shop-report-bot`](https://github.com/furqonajiy/itbisa-shop-report-bot). Copy
the relevant workbook(s) into that project's `data\` per its own README.

## Troubleshooting

| Symptom | Cause / fix |
| --- | --- |
| `Tidak ada file ditemukan di …` | No `*.xls*` / `*.csv` under the data dir. Check `--data-dir` and that files actually landed in `data\`. Run `python main.py --show-files` to list what was discovered. |
| A file is ignored | Its filename is missing the marketplace/version/type token (e.g. `BisaTransaksi v2 Shopee`). Rename to the exporter's convention. |
| `ValueError: Tidak dapat menentukan marketplace dari file: …` | The report filename has no `Shopee`/`Tiktok`/`Tokopedia`/`Bukalapak` substring — fix the source filename. |
| `ImportError` / `numpy.dtype size changed` on `import pandas` | pandas/numpy mismatch. Reinstall the pinned set: `pip install -r requirements.txt` (pandas 1.x ↔ numpy 1.x). |
| `AttributeError: property 'book' … has no setter` | pandas ≥ 2.0 is installed. Downgrade: `pip install "pandas>=1.4,<2.0"`. |
| `ValueError: Check … Keyword failed in …` | The export contains a status/saldo keyword the validator doesn't know (`keywordchecker/`). A new marketplace wording appeared — add it to the relevant `VALID_*` list. |

## Verifying a change before committing

```powershell
# compile-check every module
python -m py_compile main.py (Get-ChildItem -Recurse generator\*.py | % FullName)

# smoke-run against a scratch folder (should report "Tidak ada file ditemukan")
python main.py --data-dir .\_empty --output-dir .\_out
```
