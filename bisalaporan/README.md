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

> **Python / pandas:** this tool targets **pandas 2.x** (`requirements.txt` pins
> `pandas>=2.0,<3.0`). pandas **3.0** is not yet supported (its new `str` dtype and
> Copy-on-Write default need a separate port). Use Python **3.8–3.11**.

## CLI (`python main.py`)

| Command | Effect |
| --- | --- |
| `python main.py` | Process **every** marketplace (Bukalapak, Tokopedia, Shopee, Tiktok). |
| `python main.py --shopee` | Process Shopee only. (`--tokopedia`, `--tiktok`, `--bukalapak` likewise; flags combine.) |
| `python main.py --shopee --tiktok` | Process the selected subset. |
| `python main.py --data-dir <dir>` | Read inputs from `<dir>` instead of `./data`. |
| `python main.py --output-dir <dir>` | Write reports under `<dir>` instead of `./reports`. |
| `python main.py --show-files` | Log every input file discovered, then run. |
| `python main.py --reconcile` | Write `Rekonsiliasi <Marketplace>.xlsx` (read-only audit; see below). Generates no reports. |
| `python main.py --reconcile --bisajual-dir <dir>` | As above, but read the itbisa-shop-report-bot `*BisaJual*.xlsx` ledger from `<dir>` for the `Cek Omzet vs Fee` sheet. |
| `python main.py --cek-bisajual --bisajual-dir <dir>` | Reconcile a list of invoices against the BisaJual ledger to find entry bugs (see below). |
| `python main.py --cek-bisajual --invoices <file>` | Same, for the invoices listed in `<file>` (one per line). |
| `python main.py -v` / `--verbose` | Enable debug logging. |

## Reconciliation (`--reconcile`)

`python main.py --reconcile` writes a **read-only** `reports/<marketplace>/Rekonsiliasi
<Marketplace>.xlsx` for each marketplace that has a `BisaSaldo` (Shopee, Tokopedia,
Bukalapak). It **changes no generated numbers** — it re-reads the raw `BisaSaldo` /
`BisaFee` inputs and audits what the generator captures, so you can spot money that
silently falls out of `BisaLaporan`. Each workbook has:

- **Ringkasan** — per period: net balance change split into Remit / Bonus /
  Penarikan-Transfer / **Tidak Tercatat** (uncaptured), with a red **Perlu Dicek**
  flag when anything is uncaptured.
- **Rincian per Deskripsi** — every distinct `BisaSaldo` description (rolled up to its
  matched keyword) with its category, bucket, row count, total, and an example raw
  description. The review list: it shows exactly how each description is currently
  treated; uncaptured groups are highlighted red.
- **Rincian Saldo** — the full row-level detail (period, date, description, amount,
  category, bucket) to drill into any period.
- **Saldo Tidak Tercatat** — every `BisaSaldo` row not captured into `BisaRemit`/
  `BisaBonus`, with the reason (e.g. *matched a remit keyword but was excluded by the
  invoice filter* — this is where a `Pencairan SPinjam untuk Penjual` loan row lands).
- **Cek Omzet vs Fee** (Shopee) — per invoice, the booked `BisaJual` Omzet vs the real
  money received (`BisaSaldo`) and `BisaFee` (`Total Penghasilan` + the refund/fee
  `Kerugian`). Flags orders whose Omzet is **not** real money: *Retur — rugi = omzet*
  (a return whose loss only lives in `BisaFee`, so net is 0, e.g. a returned item — and
  in particular an order **not voided** in the ledger but with no real money),
  *Omzet tidak settle*, and *Belum ada penghasilan* (booked but not yet remitted —
  not a loss). This is how you confirm `BisaJual` represents real money. Omzet is taken
  from the **itbisa-shop-report-bot** `*BisaJual*.xlsx` ledger (non-void `Omzet Barang`)
  when you point `--bisajual-dir` at it (or drop the file in `data/`); otherwise it's
  re-derived from raw `BisaTransaksi`.
- **Cek Remit Saldo vs Fee** (Shopee) — every invoice side by side: the remit amount
  from `BisaSaldo` vs the `Total Penghasilan` from `BisaFee`, with a **Cocok**/**Beda**
  status, so you can confirm `BisaRemit` is correct against both sources at a glance.
- **BisaFee Tidak Cocok** (Shopee) — the filtered problem list: fee rows with no
  matching remit, and remit rows whose amount doesn't match the fee (so the
  `(Invoice, Nominal Remit)` join drops the fee). Overlapping fee files are
  de-duplicated first.

Run it for one marketplace with the usual flags, e.g. `python main.py --reconcile --shopee`.

## Find BisaJual entry bugs (`--cek-bisajual`)

`python main.py --cek-bisajual` reconciles a list of invoices against the
**itbisa-shop-report-bot** `BisaJual` ledger to answer one question: *how was each
order entered, vs the real money it made?* For each invoice it pairs the booked Omzet
(non-void, across every `BisaJual*` sheet) with the real money in `BisaSaldo` (full
net) and `BisaFee`, then flags the entry bugs:

| Verdict | Meaning |
| --- | --- |
| `BUG: entry hilang` | money came in, but the order isn't in BisaJual |
| `BUG: … -> Void` | Omzet booked, but no money — a return left un-voided |
| `BUG: di-Void tapi ada uang` | voided in BisaJual, yet money was received |
| `BUG: omzet != uang diterima` | booked Omzet doesn't match the money received |
| `OK: cocok / retur / void` | the entry already matches reality |

It writes `reports/shopee/Cek BisaJual Shopee.xlsx` (BUG rows red, at the top). The
invoice list defaults to a built-in set; pass `--invoices <file>` (one invoice per
line, `#` comments allowed) for any other — the list is meant to grow over time. Point
`--bisajual-dir` at the bot repo's `data/` so it reads the real ledger. Read-only.

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
  bisarekonsiliasi/
    generic.py               # --reconcile audit (BisaSaldo/BisaFee vs captured)
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
