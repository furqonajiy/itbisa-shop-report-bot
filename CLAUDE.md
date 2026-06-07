# CLAUDE.md — itbisa-shop-report-bot

Python bot for **biweekly finance report automation** across Shopee Indonesia and TikTok Shop Indonesia. **Early development / scaffolded** — structure is still being established; mirror the sibling repos' conventions and flag anything not yet decided rather than inventing new patterns.

## Stack & expected structure
- Python 3.11.
- Follow the sibling-bot layout: `src/main.py` (orchestration), `src/config.py`, per-platform clients + auth (`shopee_client.py`/`shopee_auth.py`, `tiktokshop_client.py`/`tiktokshop_auth.py`), `src/telegram_sender.py`.
- Workflow: `.github/workflows/` — `workflow_dispatch` (deliberate, like the stock bot; biweekly cadence). No order-style 5×/day cron unless explicitly decided.
- Runs once, builds the report, persists output + rotated tokens to `bot-state`, then exits.

## What it produces
- Five sheet types **per platform**: `BisaInvoice`, `BisaJual`, `BisaRemit`, `BisaFee`, `BisaBonus`.
- **`BisaRemit` is the financial source of truth.**
- `KERUGIAN TAMBAHAN` is tracked as a distinct bucket (do not fold it into another sheet).
- Output is committed to the `bot-state` branch. Transfer to Google Sheets is **manual copy-paste for now** — no Sheets API integration yet.

## API specifics already settled
- **Shopee wallet/transaction reads use a 15-day window** — chunk the date range into ≤15-day windows.
- **TikTok Shop statements endpoint requires `sort_field` and `sort_order`** — omitting them fails. Use the `202309`/`202502` versions consistent with the other bots.
- Auth reuses the established patterns: Shopee HMAC-SHA256 signed requests (shop-level signature where required); TikTok Shop OAuth with `shop_cipher` fetched per run. Each repo keeps its own bootstrap; no shared tokens.

## State / tokens
- Token files only (`data/shopee_tokens.json`, `data/tiktokshop_tokens.json`), committed to `bot-state`. Save rotated tokens immediately. No `processed_orders.json`.

## Secrets
Shopee credentials, TikTok Shop credentials, `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID`. (No `STOCK_DISPATCH_TOKEN` — this bot does not dispatch stock balance.)

## Global architecture & conventions (shared across all ITBisa repos)
- GitHub Actions only. No VM, server, database, queue, or long-running process.
- `main` = source code. `bot-state` = runtime state/token + report output files only. Never protect `bot-state`. Never commit live token files to `main`.
- Never hardcode secrets — use GitHub Secrets / env vars. Constants and API base URLs live in `src/config.py`, never as secrets.
- Self-contained repo, no shared library — duplicate platform-label constants and auth helpers rather than importing from another repo.
- Minimal, targeted changes only. No broad refactors or architecture rewrites; preserve existing behavior unless explicitly in scope.
- Telegram user-facing strings: Bahasa Indonesia. Never abbreviate "TikTok Shop" to "TikTok". Use "stock", not "inventory" (except real endpoint names).
- Platform labels (in `src/telegram_sender.py`): `SHOPEE_LABEL = "🟧Shopee"`, `TIKTOKSHOP_LABEL = "♪TikTok Shop"` (U+266A text glyph, renders in text colour — intentional). Single-space formatting; no multi-space alignment.
- GitHub Actions baseline: `actions/checkout@v5+`, `actions/setup-python@v6+` (Node 24; Node 20 actions deprecated June 2026, removed Sept 2026).
- Runtime dispatch/checkout ref is `main`. `feature/improve` must be merged to `main` before production uses it.

## Flag before changing
`BisaRemit` as source of truth, the `KERUGIAN TAMBAHAN` bucket, the five sheet-type definitions, the Shopee 15-day wallet-window chunking, the TikTok Shop statements `sort_field`/`sort_order` requirement, signing, token rotation, `bot-state`, schedule/cron, secret scope. Because the repo is early-stage, also flag any new top-level structure or naming choice before committing it.