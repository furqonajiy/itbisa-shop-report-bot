# Laporan generator (`laporan/`)

> Co-located subproject of **itbisa-shop-report-bot** — an importable `laporan` package.
> Run it from the repo root with `python main.py --laporan` (called in-process), or
> standalone with `python -m laporan`. Dependencies are the repo's root
> `requirements.txt` (`pandas>=2.0,<3.0`, `openpyxl`).

Offline Python tool that turns raw marketplace exports (Shopee, Tokopedia, Tiktok,
Bukalapak) into standardized **Laporan** workbooks. Each generated workbook holds
the bookkeeping sheets **Invoice**, **Jual**, **Remit**, **Bonus**,
and a combined **Final** sheet for one period. The **Jual** sheet is the feed
consumed by the sibling project
[`itbisa-shop-report-bot`](https://github.com/furqonajiy/itbisa-shop-report-bot).

It is fully offline and idempotent: no API calls, no network, no secrets. Drop the
marketplace exports into `data/`, run `python -m laporan`, and collect the reports from
`reports/<marketplace>/`.

## Quick start

```powershell
# Windows PowerShell — from the repo root
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt          # repo-root deps cover this generator too

# 1) put the marketplace exports under .\laporan\data\  (see "Inputs" below)

# 2a) run the generator via the bot (called in-process)
python main.py --laporan                 # every marketplace; scope: --laporan shopee tiktok

# 2b) or run it standalone as a module (same output)
python -m laporan

# reports land in laporan\reports\shopee\ , laporan\reports\tiktokshop\ , etc.
```

> **Python / pandas:** this tool targets **pandas 2.x** (the repo-root `requirements.txt`
> pins `pandas>=2.0,<3.0`). pandas **3.0** is not yet supported (its new `str` dtype and
> Copy-on-Write default need a separate port). Use Python **3.13**.

## CLI (`python -m laporan`)

| Command | Effect |
| --- | --- |
| `python -m laporan` | Process **every** marketplace (Bukalapak, Tokopedia, Shopee, Tiktok). |
| `python -m laporan --shopee` | Process Shopee only. (`--tokopedia`, `--tiktok`, `--bukalapak` likewise; flags combine.) |
| `python -m laporan --shopee --tiktok` | Process the selected subset. |
| `python -m laporan --data-dir <dir>` | Read inputs from `<dir>` instead of `./data`. |
| `python -m laporan --output-dir <dir>` | Write reports under `<dir>` instead of `./reports`. |
| `python -m laporan --show-files` | Log every input file discovered, then run. |
| `python -m laporan --reconcile` | Write `Rekonsiliasi <Marketplace>.xlsx` (read-only audit; see below). Generates no reports. |
| `python -m laporan --reconcile --jual-dir <dir>` | As above, but read the itbisa-shop-report-bot `*Jual*.xlsx` ledger from `<dir>` for the `Cek Omzet vs Fee` sheet. |
| `python -m laporan --cek-jual --jual-dir <dir>` | Reconcile a list of invoices against the Jual ledger to find entry bugs (see below). |
| `python -m laporan --cek-jual --invoices <file>` | Same, for the invoices listed in `<file>` (one per line). |
| `python -m laporan -v` / `--verbose` | Enable debug logging. |

## Reconciliation (`--reconcile`)

`python -m laporan --reconcile` writes a **read-only** `reports/<marketplace>/Rekonsiliasi
<Marketplace>.xlsx` for each marketplace that has a `Saldo` (Shopee, Tokopedia,
Bukalapak). It **changes no generated numbers** — it re-reads the raw `Saldo` /
`Fee` inputs and audits what the generator captures, so you can spot money that
silently falls out of `Laporan`. Each workbook has:

- **Ringkasan** — per period: net balance change split into Remit / Bonus /
  Penarikan-Transfer / **Tidak Tercatat** (uncaptured), with a red **Perlu Dicek**
  flag when anything is uncaptured.
- **Rincian per Deskripsi** — every distinct `Saldo` description (rolled up to its
  matched keyword) with its category, bucket, row count, total, and an example raw
  description. The review list: it shows exactly how each description is currently
  treated; uncaptured groups are highlighted red.
- **Rincian Saldo** — the full row-level detail (period, date, description, amount,
  category, bucket) to drill into any period.
- **Saldo Tidak Tercatat** — every `Saldo` row not captured into `Remit`/
  `Bonus`, with the reason (e.g. *matched a remit keyword but was excluded by the
  invoice filter* — this is where a `Pencairan SPinjam untuk Penjual` loan row lands).
- **Cek Omzet vs Fee** (Shopee) — per invoice, the booked `Jual` Omzet vs the real
  money received (`Saldo`) and `Fee` (`Total Penghasilan` + the refund/fee
  `Kerugian`). Flags orders whose Omzet is **not** real money: *Retur — rugi = omzet*
  (a return whose loss only lives in `Fee`, so net is 0, e.g. a returned item — and
  in particular an order **not voided** in the ledger but with no real money),
  *Omzet tidak settle*, and *Belum ada penghasilan* (booked but not yet remitted —
  not a loss). This is how you confirm `Jual` represents real money. Omzet is taken
  from the **itbisa-shop-report-bot** `*Jual*.xlsx` ledger (non-void `Omzet Barang`)
  when you point `--jual-dir` at it (or drop the file in `data/`); otherwise it's
  re-derived from raw `Transaksi`.
- **Cek Remit Saldo vs Fee** (Shopee) — every invoice side by side: the remit amount
  from `Saldo` vs the `Total Penghasilan` from `Fee`, with a **Cocok**/**Beda**
  status, so you can confirm `Remit` is correct against both sources at a glance.
- **Fee Tidak Cocok** (Shopee) — the filtered problem list: fee rows with no
  matching remit, and remit rows whose amount doesn't match the fee (so the
  `(Invoice, Nominal Remit)` join drops the fee). Overlapping fee files are
  de-duplicated first.

Run it for one marketplace with the usual flags, e.g. `python -m laporan --reconcile --shopee`.

## Find Jual entry bugs (`--cek-jual`)

`python -m laporan --cek-jual` reconciles a list of invoices against the
**itbisa-shop-report-bot** `Jual` ledger to answer one question: *how was each
order entered, vs the real money it made?* For each invoice it pairs the booked Omzet
(non-void, across every `Jual*` sheet) with the real money in `Saldo` (full
net) and `Fee`, then flags the entry bugs:

| Verdict | Meaning |
| --- | --- |
| `BUG: entry hilang` | money came in, but the order isn't in Jual |
| `BUG: … -> Void` | Omzet booked, but no money — a return left un-voided |
| `BUG: di-Void tapi ada uang` | voided in Jual, yet money was received |
| `BUG: omzet != uang diterima` | booked Omzet doesn't match the money received |
| `OK: cocok / retur / void` | the entry already matches reality |

It writes `reports/shopee/Cek Jual Shopee.xlsx` (BUG rows red, at the top). The
invoice list defaults to a built-in set; pass `--invoices <file>` (one invoice per
line, `#` comments allowed) for any other — the list is meant to grow over time. Point
`--jual-dir` at the bot repo's `data/` so it reads the real ledger. Read-only.

## Inputs (`data/`)

Place the raw exports anywhere under `data/` — directly in `data/`, or in
subfolders; the loader globs `data/` **recursively**. Each file is classified by the
**marketplace + version + type token in its filename**, so the filename must contain
one of:

| Type | Token examples |
| --- | --- |
| `Transaksi` (orders) | `… Transaksi v2 Shopee.xlsx`, `… Transaksi v3 Shopee.xlsx`, `… Transaksi v1 Tokopedia.xlsx`, `… Transaksi v2 Tokopedia.xlsx`, `… Transaksi v1 Tiktok.xlsx`, `… Transaksi v2 Bukalapak.xlsx` |
| `Saldo` (balance mutations) | `… Saldo v2 Shopee.csv`, `… Saldo v3 Shopee.xlsx`, `… Saldo v2 Tokopedia.xlsx`, `… Saldo v2 Bukalapak.csv` |
| `Fee` (settlement fees) | `… Fee v2 Shopee.xlsx`, `… Fee v3 Shopee.xlsx`, `… Fee v1 Tiktok.xlsx` |

`data/` and `reports/` are git-ignored — they hold your own (often large) Excel files
and the generated output, which are not version-controlled.

## Outputs (`reports/<marketplace>/`)

Reports are written into a per-marketplace subfolder:

```
reports/
  shopee/      Jan 2024 Laporan Shopee.xlsx
  tiktokshop/  Jan 2024 Laporan Tiktok.xlsx
  tokopedia/   Jan 2024 Laporan Tokopedia.xlsx
  bukalapak/   Jan 2024 Laporan v2 Bukalapak.xlsx
```

The report filename mirrors the source filename, with the type token rewritten to
`Laporan`. For one period, the orders file (`Transaksi`) and the balance/fee
file (`Saldo` / `Fee`) collapse to the **same** `… Laporan …` filename, so
their sheets accumulate into a single workbook:

| Sheet | Built from | Marketplaces |
| --- | --- | --- |
| `Invoice <MP>` | `Transaksi` | all |
| `Jual <MP>` | `Transaksi` | all *(feeds itbisa-shop-report-bot)* |
| `Remit <MP>` | `Saldo` (Tiktok: `Fee`) | Shopee, Tokopedia v2, Tiktok, Bukalapak |
| `Bonus <MP>` | `Saldo` | Shopee, Tokopedia v2 |
| `Final` | `Invoice` + `Jual` + `Remit` | all (remit columns blank where no remit) |

`Invoice` is written first (it creates the workbook); the other sheets are
appended to it.

### The `Final` sheet

`Final` is one reconciliation row per `Invoice`, joining the order side with the
remit side. Columns: `Tanggal Pesan`, `Marketplace`, `Invoice`, `Ongkir`,
`Asuransi`, `Omzet Barang`, `Nominal Invoice`, `Tanggal Remit`, `Potongan
Pembayaran`, `Nominal Remit`, `Keuntungan Tambahan`, `Kerugian Tambahan`, `Cek
Remit`, `Untung Lainnya`, `Rugi Lainnya`, `Keterangan`.

- `Omzet Barang` = sum of `Jual` `Omzet` for that `Invoice`; `Nominal
  Invoice` = `Omzet Barang` + `Ongkir` + `Asuransi`.
- Every `Invoice` order is listed (left join). The remit is looked up across
  **all** of that marketplace's `Laporan` workbooks — so an order placed this
  month but remitted next month still finds its remit; if none is found the remit
  columns stay blank.
- `Cek Remit`, `Untung Lainnya`, `Rugi Lainnya`, and `Keterangan` are left blank
  for manual entry.

The sheet mirrors the `Invoice` styling (row-number gutter, bold centered
headers) and **color-codes the three column groups** by source: the order side
(`Tanggal Pesan` → `Nominal Invoice`) is blue, the remit side (`Tanggal Remit` →
`Cek Remit`) is green, and the manual columns (`Untung Lainnya` → `Keterangan`)
are amber. Money columns use a thousands separator.

Because the remit lookup spans periods, `Final` is built **after** all of a
marketplace's workbooks have been generated.

## Marketplace coverage

| Marketplace | Versions | Sheets generated |
| --- | --- | --- |
| Shopee | v2, v3 | Invoice, Jual, Remit, Bonus |
| Tokopedia | v1, v2 | v1: Invoice, Jual · v2: + Remit, Bonus |
| Tiktok | v1 | Invoice, Jual, Remit (from `Fee`) |
| Bukalapak | v2 | Invoice, Jual, Remit |

Every workbook also gets a `Final` sheet (the sheets listed above plus `Final`).

## Project layout

```
laporan/                     # this folder — the importable `laporan` package
  __init__.py  __main__.py   # package marker + `python -m laporan` entry point
  main.py                    # CLI (argparse) + orchestration: MARKETPLACE_PROCESSORS + run() + Final
  process/
    preprocess.py            # recursive discovery of data/ inputs
    <marketplace>/<vN>.py    # per-marketplace/version readers + filters
  invoice/ jual/ remit/ bonus/ fee/
    generic.py               # the shared *_to_excel writer for each sheet type
    <marketplace>/<vN>.py    # per-marketplace sheet builders
  final/
    generic.py               # builds the Final sheet (marketplace-agnostic)
  rekonsiliasi/
    generic.py               # --reconcile audit (Saldo/Fee vs captured)
  keywordchecker/            # validates marketplace status / saldo keywords
  utility/
    constant.py              # data/report dirs + marketplace->folder mapping
    generic.py               # warnings, create_directory, build_report_path
    sku.py                   # SKU standardization
  data/                      # raw marketplace exports (sample set committed)
  reports/                   # generated Laporan workbooks (gitignored)
```

See [`RUNBOOK.md`](RUNBOOK.md) for step-by-step operating instructions and
[`CLAUDE.md`](CLAUDE.md) for the engineering/contribution conventions.

## Spreadsheet helpers

VLOOKUP used to pull columns from a generated `Laporan` sheet into another sheet:

```
=IFERROR(VLOOKUP($I2,$B$2:$G$1600,J$1,FALSE),"")
```
