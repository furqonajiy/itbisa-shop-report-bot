# itbisa-bisalaporan

Offline Python tool that turns raw marketplace exports (Shopee, Tokopedia, Tiktok,
Bukalapak) into standardized **BisaLaporan** workbooks. Each generated workbook holds
the bookkeeping sheets **BisaInvoice**, **BisaJual**, **BisaRemit**, **BisaBonus**,
and a combined **Final** sheet for one period. The **BisaJual** sheet is the feed
consumed by the sibling project
[`itbisa-shop-report-bot`](https://github.com/furqonajiy/itbisa-shop-report-bot).

It is fully offline and idempotent: no API calls, no network, no secrets. Drop the
marketplace exports into `data/`, run `python main.py`, and collect the reports from
`reports/<marketplace>/`.

## Quick start

```powershell
# Windows PowerShell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

# 1) put the marketplace exports under .\data\  (see "Inputs" below)
# 2) generate every report
python main.py

# reports land in .\reports\shopee\ , .\reports\tiktokshop\ , etc.
```

> **Python / pandas:** this tool targets **pandas 1.x** (`requirements.txt` pins
> `pandas>=1.4,<2.0`), because the Excel writers use the pandas-1.x
> `ExcelWriter.book` API. pandas 1.5 supports Python **3.8–3.11**, so use one of
> those interpreters.

## CLI (`python main.py`)

| Command | Effect |
| --- | --- |
| `python main.py` | Process **every** marketplace (Bukalapak, Tokopedia, Shopee, Tiktok). |
| `python main.py --shopee` | Process Shopee only. (`--tokopedia`, `--tiktok`, `--bukalapak` likewise; flags combine.) |
| `python main.py --shopee --tiktok` | Process the selected subset. |
| `python main.py --data-dir <dir>` | Read inputs from `<dir>` instead of `./data`. |
| `python main.py --output-dir <dir>` | Write reports under `<dir>` instead of `./reports`. |
| `python main.py --show-files` | Log every input file discovered, then run. |
| `python main.py -v` / `--verbose` | Enable debug logging. |

## Inputs (`data/`)

Place the raw exports anywhere under `data/` — directly in `data/`, or in
subfolders; the loader globs `data/` **recursively**. Each file is classified by the
**marketplace + version + type token in its filename**, so the filename must contain
one of:

| Type | Token examples |
| --- | --- |
| `BisaTransaksi` (orders) | `… BisaTransaksi v2 Shopee.xlsx`, `… BisaTransaksi v3 Shopee.xlsx`, `… BisaTransaksi v1 Tokopedia.xlsx`, `… BisaTransaksi v2 Tokopedia.xlsx`, `… BisaTransaksi v1 Tiktok.xlsx`, `… BisaTransaksi v2 Bukalapak.xlsx` |
| `BisaSaldo` (balance mutations) | `… BisaSaldo v2 Shopee.csv`, `… BisaSaldo v3 Shopee.xlsx`, `… BisaSaldo v2 Tokopedia.xlsx`, `… BisaSaldo v2 Bukalapak.csv` |
| `BisaFee` (settlement fees) | `… BisaFee v2 Shopee.xlsx`, `… BisaFee v3 Shopee.xlsx`, `… BisaFee v1 Tiktok.xlsx` |

`data/` and `reports/` are git-ignored — they hold your own (often large) Excel files
and the generated output, which are not version-controlled.

## Outputs (`reports/<marketplace>/`)

Reports are written into a per-marketplace subfolder:

```
reports/
  shopee/      Jan 2024 BisaLaporan Shopee.xlsx
  tiktokshop/  Jan 2024 BisaLaporan Tiktok.xlsx
  tokopedia/   Jan 2024 BisaLaporan Tokopedia.xlsx
  bukalapak/   Jan 2024 BisaLaporan v2 Bukalapak.xlsx
```

The report filename mirrors the source filename, with the type token rewritten to
`BisaLaporan`. For one period, the orders file (`BisaTransaksi`) and the balance/fee
file (`BisaSaldo` / `BisaFee`) collapse to the **same** `… BisaLaporan …` filename, so
their sheets accumulate into a single workbook:

| Sheet | Built from | Marketplaces |
| --- | --- | --- |
| `BisaInvoice <MP>` | `BisaTransaksi` | all |
| `BisaJual <MP>` | `BisaTransaksi` | all *(feeds itbisa-shop-report-bot)* |
| `BisaRemit <MP>` | `BisaSaldo` (Tiktok: `BisaFee`) | Shopee, Tokopedia v2, Tiktok, Bukalapak |
| `BisaBonus <MP>` | `BisaSaldo` | Shopee, Tokopedia v2 |
| `Final` | `BisaInvoice` + `BisaJual` + `BisaRemit` | all (remit columns blank where no remit) |

`BisaInvoice` is written first (it creates the workbook); the other sheets are
appended to it.

### The `Final` sheet

`Final` is one reconciliation row per `Invoice`, joining the order side with the
remit side. Columns: `Tanggal Pesan`, `Marketplace`, `Invoice`, `Ongkir`,
`Asuransi`, `Omzet Barang`, `Nominal Invoice`, `Tanggal Remit`, `Potongan
Pembayaran`, `Nominal Remit`, `Keuntungan Tambahan`, `Kerugian Tambahan`, `Cek
Remit`, `Untung Lainnya`, `Rugi Lainnya`, `Keterangan`.

- `Omzet Barang` = sum of `BisaJual` `Omzet` for that `Invoice`; `Nominal
  Invoice` = `Omzet Barang` + `Ongkir` + `Asuransi`.
- Every `BisaInvoice` order is listed (left join). The remit is looked up across
  **all** of that marketplace's `BisaLaporan` workbooks — so an order placed this
  month but remitted next month still finds its remit; if none is found the remit
  columns stay blank.
- `Cek Remit`, `Untung Lainnya`, `Rugi Lainnya`, and `Keterangan` are left blank
  for manual entry.

The sheet mirrors the `BisaInvoice` styling (row-number gutter, bold centered
headers) and **color-codes the three column groups** by source: the order side
(`Tanggal Pesan` → `Nominal Invoice`) is blue, the remit side (`Tanggal Remit` →
`Cek Remit`) is green, and the manual columns (`Untung Lainnya` → `Keterangan`)
are amber. Money columns use a thousands separator.

Because the remit lookup spans periods, `Final` is built **after** all of a
marketplace's workbooks have been generated.

## Marketplace coverage

| Marketplace | Versions | Sheets generated |
| --- | --- | --- |
| Shopee | v2, v3 | BisaInvoice, BisaJual, BisaRemit, BisaBonus |
| Tokopedia | v1, v2 | v1: BisaInvoice, BisaJual · v2: + BisaRemit, BisaBonus |
| Tiktok | v1 | BisaInvoice, BisaJual, BisaRemit (from `BisaFee`) |
| Bukalapak | v2 | BisaInvoice, BisaJual, BisaRemit |

Every workbook also gets a `Final` sheet (the sheets listed above plus `Final`).

## Project layout

```
main.py                      # CLI entry point (argparse) -> generator.run(...)
requirements.txt
generator/
  main.py                    # orchestration: MARKETPLACE_PROCESSORS + run() + Final
  process/
    preprocess.py            # recursive discovery of data/ inputs
    <marketplace>/<vN>.py    # per-marketplace/version readers + filters
  bisainvoice/ bisajual/ bisaremit/ bisabonus/ bisafee/
    generic.py               # the shared *_to_excel writer for each sheet type
    <marketplace>/<vN>.py    # per-marketplace sheet builders
  bisafinal/
    generic.py               # builds the Final sheet (marketplace-agnostic)
  keywordchecker/            # validates marketplace status / saldo keywords
  utility/
    constant.py              # data/report dirs + marketplace->folder mapping
    generic.py               # warnings, create_directory, build_report_path
    sku.py                   # SKU standardization
```

See [`RUNBOOK.md`](RUNBOOK.md) for step-by-step operating instructions and
[`CLAUDE.md`](CLAUDE.md) for the engineering/contribution conventions.

## Spreadsheet helpers

VLOOKUP used to pull columns from a generated `BisaLaporan` sheet into another sheet:

```
=IFERROR(VLOOKUP($I2,$B$2:$G$1600,J$1,FALSE),"")
```
