# itbisa-shop-report-bot

Generates the **BisaLaporan** finance reports (BisaInvoice, BisaJual,
BisaRemit, BisaFee, BisaBonus) for Shopee Indonesia and TikTok Shop
Indonesia — same shape as the legacy `itbisa-bisalaporan` Python
generator, but driven entirely from the Shopee Open API and TikTok
Shop Open API instead of the manually-downloaded seller-center Excel
exports.

> **Status: PR 1 — scaffolding + token bootstrap + reason-code
> discovery.** The xlsx generators come in PR 2; the Worker
> integration (`/report_shop_shopee`, `/report_shop_tiktokshop`,
> `/report_order_all`) comes in PR 3. See
> [Roadmap](#roadmap) below.

## What it does (when complete)

User types `/report_shop_shopee 2026-01` in Telegram →
`@ITBisaShopBot` Worker dispatches a workflow on this repo →
GitHub Actions calls the marketplace API for the period →
generates a single xlsx with all five sheets →
commits the xlsx to the `bot-state` branch under `output/shopee/` →
sends the file to Telegram via `sendDocument`.

You then `git pull` on your machine and copy the xlsx into the
matching folder under
`H:\My Drive\...\BisaLaporan - Marketplace\` — same place the
legacy generator wrote to.

Two output formats per platform:

```text
ITBisa_com_-_BisaLaporan_Shopee_2026-04.xlsx        # finalized month
ITBisa_com_-_BisaLaporan_Shopee_2026-04-15.xlsx     # in-progress month
```

The TikTok Shop workflow produces both `Tiktok` and `Tokopedia` xlsx
files, since the same TikTok Open API serves both purchase channels
(split by `purchase_channel`).

## Architecture

```text
User (Telegram)
  → @ITBisaShopBot (Cloudflare Worker)
      → workflow_dispatch (this repo)
          → Python: pull API data → build xlsx → commit + Telegram

Token chain: independent per repo
  Same Shopee partner key + TikTok app secret as the order bots,
  but each repo holds its own data/*_tokens.json on its own
  bot-state branch. You re-authorize Shopee + TikTok once during
  setup so the new repo gets a fresh token pair that does not
  collide with the order bots' rotation.

State on bot-state branch
  data/shopee_tokens.json
  data/tiktokshop_tokens.json
  state/shopee_periods.json        (PR 2)
  state/tiktokshop_periods.json    (PR 2)
  output/shopee/*.xlsx              (PR 2)
  output/tiktokshop/*.xlsx          (PR 2)
  output/tokopedia/*.xlsx           (PR 3)

Source of truth: API only
  Shopee:  payment.get_wallet_transaction_list   → BisaRemit, BisaBonus, KERUGIAN/KEUNTUNGAN TAMBAHAN
           payment.get_escrow_detail             → BisaFee per order
           order.get_order_list + get_order_detail → BisaInvoice, BisaJual

  TikTok:  finance.statements + statement_transactions  → BisaRemit, BisaBonus, KERUGIAN/KEUNTUNGAN
           finance.order_settlements                    → BisaFee per order
           order.search                                 → BisaInvoice, BisaJual
           split by purchase_channel: TIKTOK | TOKOPEDIA → 2 xlsx
```

## Project structure

```text
itbisa-shop-report-bot/
├── .github/workflows/
│   ├── shopee_report.yml           (PR 3)
│   └── tiktokshop_report.yml       (PR 3)
├── data/                           # bot-state: token files only
│   ├── shopee_tokens.json
│   └── tiktokshop_tokens.json
├── output/                         # bot-state: generated xlsx (PR 2+)
├── state/                          # bot-state: period tracking (PR 2+)
├── scripts/
│   ├── bootstrap_shopee_tokens.py        ✅ PR 1
│   ├── bootstrap_tiktokshop_tokens.py    ✅ PR 1
│   └── dump_reason_codes.py              ✅ PR 1
├── src/
│   ├── __init__.py                       ✅ PR 1
│   ├── main.py                           ✅ PR 1 (CLI orchestrator, discovery mode)
│   ├── config.py                         ✅ PR 1
│   ├── discovery.py                      ✅ PR 1 (reason-code aggregation)
│   ├── shopee_auth.py                    ✅ PR 1
│   ├── shopee_client.py                  ✅ PR 1 (auth + wallet endpoint)
│   ├── tiktokshop_auth.py                ✅ PR 1
│   ├── tiktokshop_client.py              ✅ PR 1 (auth + finance statements)
│   ├── telegram_sender.py                ✅ PR 1
│   ├── reason_mapping/                   (PR 2)
│   ├── generators/                       (PR 2)
│   ├── excel_writer.py                   (PR 2)
│   ├── sku_normalizer.py                 (PR 2)
│   └── state_manager.py                  (PR 2)
├── requirements.txt
├── .env.example
├── .gitignore
└── README.md
```

## Initial setup (PR 1)

### 1. Clone and install

```bash
conda create -n itbisa_shop_report_bot python=3.11
conda activate itbisa_shop_report_bot
cd C:\path\to\itbisa-shop-report-bot
python -m pip install -r requirements.txt
```

### 2. Configure secrets

Copy `.env.example` to `.env` and fill in. **Use the same Shopee
partner key and TikTok app secret as the order bots.**

### 3. Bootstrap tokens (independent from the order bots)

You authorize the shop fresh for this repo so it gets its own
token pair — no collision with the order bots' rotation.

```bash
# Shopee:
#   1. Log into Shopee Open Platform Console.
#   2. Open your app → Authorize → confirm shop.
#   3. Copy "code" from the redirect URL.
#   4. Run:
python scripts/bootstrap_shopee_tokens.py

# TikTok Shop:
#   1. Log into TikTok Shop Partner Center.
#   2. Open your app → authorize for shop.
#   3. Copy "auth_code" from the redirect URL.
#   4. Run:
python scripts/bootstrap_tiktokshop_tokens.py
```

This writes `data/shopee_tokens.json` and `data/tiktokshop_tokens.json`.

### 4. Push to GitHub

```bash
git checkout -b feature/initialize-app
git add .
git commit -m "PR 1: scaffold, auth, token bootstrap, reason-code discovery"
git push origin feature/initialize-app
```

Once merged into main, manually create the `bot-state` branch and
push the two token JSON files there. The PR 3 workflow will overlay
them at runtime.

```bash
git checkout --orphan bot-state
git rm -rf .
mkdir -p data
cp /path/to/shopee_tokens.json data/
cp /path/to/tiktokshop_tokens.json data/
git add data/
git commit -m "Bootstrap initial state files"
git push origin bot-state
```

### 5. Configure GitHub Secrets (when PR 3 lands)

Go to **Settings → Secrets and variables → Actions** and add:

- `SHOPEE_PARTNER_ID`
- `SHOPEE_PARTNER_KEY`
- `SHOPEE_SHOP_ID`
- `TIKTOKSHOP_APP_KEY`
- `TIKTOKSHOP_APP_SECRET`
- `TIKTOKSHOP_SHOP_ID`
- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`

## Discovery — `dump_reason_codes.py`

Before PR 2 can be written, we need to know **every distinct reason
code / adjustment type** that appears in your wallet transactions
and finance statements. The legacy generator handled this with
hand-maintained Bahasa keyword lists in
`generator/keywordchecker/{shopee,tiktok,tokopedia}.py`. The API
returns typed reason codes per row, so we can replace those lists
with a clean code → bucket mapping — but only after we see what
codes the platforms actually return for your shop.

Run once per month for January 2026 → today and paste the output
back to me. I will use it to build the reason-mapping table in
PR 2.

**Recommended — single command per month, both platforms:**

```bash
# Walks both Shopee + TikTok Shop sequentially, saves discovery_*.json
# for each, and pings Telegram with start/end heartbeats.
python -m src.main --period 2026-01
python -m src.main --period 2026-02
python -m src.main --period 2026-03
python -m src.main --period 2026-04
```

Useful flags:

```bash
python -m src.main --period 2026-01 --platform shopee     # one platform only
python -m src.main --period 2026-01 --platform tiktokshop
python -m src.main --period 2026-01 --no-save             # skip JSON dump
python -m src.main --period 2026-01 --no-telegram         # skip heartbeats
```

**Alternative — direct script for one platform at a time** (no Telegram, no auth health-check):

```bash
python scripts/dump_reason_codes.py --platform shopee     --period 2026-01
python scripts/dump_reason_codes.py --platform tiktokshop --period 2026-01
```

Both entry points share the same discovery logic in `src/discovery.py`
and produce identical `discovery_<platform>_<period>.json` files.

The script prints to stdout AND writes `discovery_<platform>_<period>.json`
for offline review. Both outputs are gitignored — they contain
shop-internal financial data.

## Roadmap

| PR  | Scope                                                                                                  | Status     |
| --- | ------------------------------------------------------------------------------------------------------ | ---------- |
| 1   | Repo scaffold, `config.py`, auth modules, minimal clients, `bootstrap_*` scripts, `dump_reason_codes`  | ✅ this PR |
| 2   | Reason-mapping tables, BisaInvoice + BisaJual + BisaRemit generators, xlsx writer, `main.py` for one period via CLI, state tracking | 🔜       |
| 3   | BisaFee, BisaBonus, Tokopedia split, Worker `/report_shop_*` commands, workflow_dispatch, range backfill | 🔜       |

## Cost

Free forever once running. GitHub Actions free-tier minutes only.
Each report run is well under one minute of compute even for a full
month. No server, no VM, no database.

## Cross-references

- **`itbisa-shopee-order-bot`** — source for `shopee_auth.py`,
  `shopee_client.py` signing logic. Independent token chain.
- **`itbisa-tiktokshop-order-bot`** — source for `tiktokshop_auth.py`,
  `tiktokshop_client.py` signing + shop_cipher caching. Independent
  token chain.
- **`itbisa-shop-stock-bot`** (a.k.a. `itbisa-inventory-bot`) — same
  setup pattern as this repo (independent tokens, two bootstrap
  scripts, bot-state branch contract).
- **`itbisa-shop-telegram-bot`** — Cloudflare Worker. PR 3 will add
  the three `/report_*` commands here.
- **`itbisa-bisalaporan`** (legacy) — manual Python generator.
  Reference for output sheet shapes and the SKU normalization
  rules that PR 2 will lift verbatim.
