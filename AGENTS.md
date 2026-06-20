# AGENTS.md — itbisa-bisalaporan

Instructions for **ChatGPT Codex** (and any other coding agent).

## Read this first
**The full project + agent instructions live in [`CLAUDE.md`](CLAUDE.md).** It is the single source of truth — read it and follow it. `AGENTS.md` and `CLAUDE.md` describe the same repo; `CLAUDE.md` is kept current to avoid drift, so defer to it for the codebase layout, CLI, inputs/outputs, core logic, and the "flag before changing" list. Operating steps are in [`RUNBOOK.md`](RUNBOOK.md).

## Non-negotiables (also in CLAUDE.md)
- **Author** every commit/PR as `C - Furqon Aji Yudhistira <furqonajiy@gmail.com>` (set `git config user.name`/`user.email` at session start).
- **No AI/assistant references anywhere** — branch names, commit messages, PR titles/bodies, code/comments. No "Claude"/"Anthropic"/"Codex"/"OpenAI", no `Co-Authored-By` trailer, no "Generated with …" line, no model ids, no session links. History must read as human-authored. After opening a PR, strip any auto-appended footer.
- **Workflow**: branch `feature/<short-desc>` off `main`; doc updates ride in the same PR as the code; PR → **merge commit (`--no-ff`)** with a representative title ending in the PR number.
- **CLI handed to the user**: PowerShell syntax (the user is on Windows).
- **Docs language**: all Markdown (`*.md`) is written in **English prose**, keeping special Indonesian/domain terms verbatim — input type tokens (`BisaTransaksi`, `BisaSaldo`, `BisaFee`), sheet names (`BisaInvoice`, `BisaJual`, `BisaRemit`, `BisaBonus`), domain terms (`omzet`, `Toko`), and any quoted console string. The console strings emitted by the code stay Bahasa Indonesia.
- **pandas `>=2.0,<3.0`** is a hard constraint (pandas 3.0's new `str` dtype + Copy-on-Write need a separate port).
- Keep changes minimal and targeted; respect the **flag-before-changing** items in `CLAUDE.md`.

## AI-instruction file map
- `CLAUDE.md` — Claude Code **and Claude Chat** (full detail; this `AGENTS.md` → Codex defers to it).
- `AGENTS.md` — ChatGPT Codex (this file).
- `README.md` — human-facing overview & usage. `RUNBOOK.md` — operating procedure.
