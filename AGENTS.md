# AGENTS.md — itbisa-shop-report-bot

Instructions for **ChatGPT Codex** (and any other coding agent).

## Read this first
**The full project + agent instructions live in [`CLAUDE.md`](CLAUDE.md).** It is the single source of truth — read it and follow it. `AGENTS.md` and `CLAUDE.md` describe the same repo; `CLAUDE.md` is kept current to avoid drift, so defer to it for the codebase layout, CLI, inputs/outputs, core analysis logic, and the "flag before changing" list.

## Non-negotiables (also in CLAUDE.md)
- **Author** every commit/PR as `C - Furqon Aji Yudhistira <furqonajiy@gmail.com>` (set `git config user.name`/`user.email` at session start).
- **No AI/assistant references anywhere** — branch names, commit messages, PR titles/bodies, code/comments. No "Claude"/"Anthropic"/"Codex"/"OpenAI", no `Co-Authored-By` trailer, no "Generated with …" line, no model ids, no session links. History must read as human-authored. After opening a PR, strip any auto-appended footer.
- **Workflow**: branch `feature/<short-desc>` off `main`; doc/marker updates ride in the same PR as the code; PR → **merge commit (`--no-ff`)** with a representative title ending in the PR number.
- **Sync marker**: rename the root `YYYY-MM-DD_HHMM.txt` (WIB) to the current timestamp on every update.
- **CLI handed to the user**: PowerShell syntax (the user is on Windows).
- Keep changes minimal and targeted; respect the **flag-before-changing** items in `CLAUDE.md`.

## AI-instruction file map
- `CLAUDE.md` — Claude Code **and Claude Chat** (full detail; this `AGENTS.md` → Codex defers to it).
- `AGENTS.md` — ChatGPT Codex (this file).
- `CHATGPT_CHAT.md` — **ChatGPT Chat only** (≤8000-char condensed copy; exists because of ChatGPT's 8K limit).
