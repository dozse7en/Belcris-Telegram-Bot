# Belcris Inventory Bot — CHANGELOG

This file tracks all changes, fixes, lessons learned, and architectural decisions.
**Always update this file when making changes so future sessions can retrace history.**

---

## Architecture Overview

| Component | Detail |
|---|---|
| Language | Python 3.11.9 |
| Framework | python-telegram-bot v20+ (async) |
| Web server | Flask (serves webhook endpoint) |
| Deployment | Railway (GitHub-connected, auto-deploy on push) |
| Data source | Google Drive (Excel files via direct download URL) |
| Bot token | Stored in Railway Variables as `BOT_TOKEN` |
| Webhook URL | `https://inventory-bot-production-af17.up.railway.app/webhook` |
| GitHub repo | `https://github.com/dozse7en/Belcris-Telegram-Bot` |

---

## Environment Variables (Railway)

| Variable | Purpose |
|---|---|
| `BOT_TOKEN` | Telegram bot token from BotFather |
| `INVENTORY_FILE_ID` | Google Drive file ID for inventory Excel |
| `AR_FILE_ID` | Google Drive file ID for AR Excel |
| `AP_FILE_ID` | Google Drive file ID for AP Excel |
| `ADMIN_IDS` | Comma-separated Telegram user IDs with admin rights (required for `/register`, `/unregister`, `/listusers`, `/seen`, `/accessmode`) |
| `ACCESS_MODE` | `off` (open, IDs logged only) / `soft` (warn unregistered, still allow) / `hard` (block unregistered) — default `off` |
| `ALLOWED_GROUP_CHAT_IDS` | Comma-separated negative group Chat IDs that are fully open for all members regardless of `ACCESS_MODE` (recommended: the internal team chat) |
| `ACCESS_DB_PATH` | Optional override for the SQLite access-control DB path (default: next to bot.py, `access_control.db`) |

**Never hardcode these in bot.py. Always use Railway Variables.**

---

## Version History

### v4.3 — 2026-09-03
**Segment Support & Accuracy Improvements:**
- **New Arrival Handling:** `/slowmoving` now automatically excludes new items that were first received within the analysis period (default 30 days). This prevents new stock from being incorrectly flagged as slow-moving before it has a chance to sell.
- **Expiring Report Segment Filtering:** `/expiring [days] [segment]` now supports filtering by area/segment (e.g., `/expiring 30 Manila`). This allows sales offices to monitor their specific upcoming expirations.
- **Slow-Moving Segment Filtering:** `/slowmoving [days] [segment]` now supports filtering by sales office. It filters both current inventory warehouses and Portal DB sales/transaction records by the segment keyword.
- **Improved Accuracy:**
    - Fixed header lookup logic to handle zero-index columns correctly.
    - Enhanced date parsing to support multiple formats (`YYYY/MM/DD`, `DD/MM/YYYY`, `MM/DD/YYYY`, and dash-separated variants).
    - Added segment-aware sales office filtering in Portal DB queries (`location` field).
- **Diagnostic Tool:** New admin-only command `/itemactivity <itemCode> [days]` provides a deep-dive report on stock levels, Portal DB sales, and transaction history for a specific item to verify slow-moving status.
- **Unified Logic:** The weekly automated slow-moving report now uses the same robust analysis logic as the manual command.
- **Show All Support:** Added segment persistence in "Show All" callbacks for both `/slowmoving` and `/expiring` reports.

### v4.2 — 2026-08-10
**Help visibility fix:**
- `/help` and `/start` now show the Admin Commands section **only to users in `ADMIN_IDS`**. Non-admins (incl. regular staff and the boss) see just the Inventory / AR / AP / Sales sections — no knowledge of admin commands.
- Admins still see the full help including `/register`, `/whitelistgroup`, etc.

---

### v4.1 — 2026-08-10
**Group whitelist (open chats — zero action needed from group members, incl. the boss):**
- Whitelisted group chats are fully open for ALL members regardless of `ACCESS_MODE` — no registration, no notices, no blocking. The internal team chat (where the boss is present) is handled this way.
- Groups are passively logged on every interaction (`seen_groups` table: chat ID, type, title, timestamps, count) — the same DB the user log uses.
- New admin commands: `/whitelistgroup [chat_id]` (no arg = current chat; arg = given group ID, useful from a private chat), `/unallowgroup <chat_id>`, `/allowedgroups`, `/seengroups` (groups used but not yet open).
- New env var `ALLOWED_GROUP_CHAT_IDS` — comma-separated negative group Chat IDs pre-opened at boot (alternatively manage via `/whitelistgroup`).
- Enforcement entry points updated: `_access_gate` (all commands, free text) and all 6 inline-button callbacks (`cb_search_item_check`, `cb_check_show_all`, `cb_expiring_show_all`, `cb_warehouse_show_all`, `cb_components_show_all`, `cb_sales_area_pick`) now log the chat and exempt whitelisted groups before the registered-user check.
- Whitelisting logic never blocks admin commands; private chats (positive IDs) can never be whitelisted.

---

### v4.0 — 2026-08-10
**Access control rollout (batch registration + `/listusers`):**
- Passive ID collection: every user who interacts with the bot (commands, free text, inline buttons) is now logged to a local SQLite DB (`access_control.db`): `seen_users` table (ID, username, name, first/last seen, message count). Users no longer need to manually send `/myid` — admins can register them from the collected log.
- `registered_users` table with full audit trail: who registered whom, when, and an optional note.
- `/register <id1>,<id2>,... [note]` — single or batch registration from collected IDs. Shows known name/username from the log next to each ID so admins can verify before registering. Duplicate IDs are reported as warnings, not errors.
- `/unregister <id>` — remove a registered user.
- `/listusers` — list all registered users with IDs, names, last seen, and registration notes.
- `/seen` — list users who have interacted with the bot but are NOT yet registered — the primary batch-registration view.
- `/accessmode [off|soft|hard]` — view or set the enforcement mode at runtime (must also set `ACCESS_MODE` in Railway variables to persist across restarts).
- `/myid` now shows an admin badge for users in `ADMIN_IDS` and passively logs the caller.
- Soft/hard enforcement applied at every entry point: all inventory/AR/AP/sales commands, free-text search, unknown-command replies, and all inline-button callbacks (search item buttons, Show All buttons, sales area picker).
- `/help` now includes the Admin Commands section.

**Architecture:**
- `ADMIN_IDS` (Railway variable, comma-separated) gates admin commands. Admin commands are never access-gated themselves.
- `ACCESS_MODE` (Railway variable): `off` (default) = fully open, log only; `soft` = warn unregistered users but allow; `hard` = block unregistered users from all data commands.
- SQLite DB kept next to the code; Railway disk is ephemeral, but the seen-user log rebuilds itself passively and registration state is re-created from `/seen` in under a minute.

---

### v3.2.1 — 2026-07-29
**Bugfix:**
- `/components` "Show All" button was silently truncating output at ~4000 characters when all 94 items exceeded Telegram's message limit. Root cause: the chunking logic was correct but the `if len(full_text) <= MAX_LEN` branch bypassed chunking entirely for the common case, and only the first chunk was ever sent via `edit_message_text` — subsequent chunks were sent as `reply_text` but only when the single-message path was NOT taken. Fixed by always running the chunking logic, editing the original message with chunk 1, and sending remaining chunks as follow-up reply messages.

---

### v3.2 — 2026-07-29
**Features added:**
- `/components [keyword]` — New command for PRD (Product Research & Development) to query Component Warehouse stock without it appearing in general inventory results. Component Warehouse rows are now captured separately during `load_inventory()` and stored in `store.component_inventory`. Supports optional keyword filtering, expiry colour-coding, "Show All" inline button when results exceed 20 items.
- Expiry dates now always shown for every component item in `/components` output (not just near-expiry). Format: `exp MM/DD/YY` for normal items, `🟠 exp MM/DD/YY (Nd)` for ≤30 days, `🟡` for ≤90 days, `🔴 EXPIRED (exp MM/DD/YY)` for expired. Extracted into shared `_exp_tag_full()` helper used by both `/components` and the Show All callback.
- `/components_expiring [days]` — Lists only Component Warehouse items expiring within N days (default 30), sorted by nearest expiry first. Sections: 🔴 Already Expired / 🟠 Expiring soon. Supports `/components_expiring 0` for expired-only view.
- Updated `HELP_TEXT`, `correct_command_typo` known-command list, README, and CHANGELOG quick reference to include both new commands.

---

### v3.1 — 2026-07-01
**Fixes:**
- Silent failure on "pork" and similar free-text searches — root cause was Markdown parse error when item descriptions contained special characters (`*`, `_`, `` ` ``, `[`, `]`). Fixed by removing `parse_mode=MARKDOWN` from plain text chunks and stripping special chars from descriptions.
- Search results exceeding Telegram's 4096-char limit — fixed by splitting results into multiple messages (chunks of ~3800 chars).
- `/refresh` showing stale "As of 6/23/2026" date — was using `store.last_refresh` (internal clock, stuck at last Railway restart) instead of `store.inventory_source_ts` (actual Google Drive file timestamp). Fixed to always show source file timestamp.
- Railway build failure after GitHub connection — `mise` was failing to install Python due to GitHub artifact attestation verification. Fixed by adding `mise.toml` with `python.github_attestations = false`.
- Webhook sync error (`last_synchronization_error_date` from 2026-06-29) — reset webhook via Telegram API to force clean reconnection.

**Features added:**
- Clickable inline keyboard buttons for top 10 search results — tap to instantly view `/check` detail without typing.
- Item codes formatted as monospace (`` `ITEMCODE` ``) in search results — desktop users can click to copy.
- Button labels show item code + brand name extracted from description (e.g. `TRPBSBN012  PORK BELLY BI/SO [SEARA]`).

**Infrastructure:**
- Bot source code recovered from Railway Docker container and pushed to GitHub (`dozse7en/Belcris-Telegram-Bot`).
- Railway service reconnected from Docker image source to GitHub repo — now auto-deploys on every push to `main`.
- CHANGELOG.md created (this file).

---

### v3.0 — 2026-06-08 (original Railway deployment)
**Features:**
- Full inventory search with free-text and `/search` command
- `/check <item_code>` — full item detail with batches, expiry, warehouse breakdown
- `/category` — browse by product category (HAM, BACON, TOCINO, etc.)
- `/expiring [days]` — items expiring within N days
- `/low` / `/lowstock` — items below stock threshold
- `/warehouse` — items by warehouse
- `/top` — top 20 items by quantity
- `/summary` — full inventory snapshot
- `/refresh` — manual data reload
- AR commands: `/arsummary`, `/ar`, `/aging`, `/overdue`, `/area`, `/agent`, `/arsearch`
- AP commands: `/apsummary`, `/ap`, `/apaging`, `/apoverdue`, `/aptop`, `/due_today`
- Auto-refresh every 30 minutes in background thread
- `//` prefix to silently ignore messages (private notes)
- Typo correction for mistyped commands
- "Show All" inline buttons for long batch/expiry/warehouse lists
- Data loaded from Google Drive Excel files on startup

---

## Known Issues & Limitations

| Issue | Status | Notes |
|---|---|---|
| `CATEGORY_KEYWORDS` has no PORK entry | Open | Pork items still searchable via free-text; `/category pork` works via fallback |
| `/refresh` summary "As of" date | Fixed v3.1 | Was showing internal clock instead of source file timestamp |
| Markdown parse errors on special chars | Fixed v3.1 | Item descriptions with `*_\`[]` broke Telegram Markdown parser |
| Long search results (78+ items) | Fixed v3.1 | Now split into multiple messages |
| Railway build attestation error | Fixed v3.1 | Added `mise.toml` with `python.github_attestations = false` |

---

## Recurring Bug Patterns — LESSONS LEARNED

### 1. Telegram Markdown Parse Errors (CRITICAL)
**Problem:** Using `parse_mode=ParseMode.MARKDOWN` when message content contains `*`, `_`, `` ` ``, `[`, `]` causes a `BadRequest` exception. The bot sends **nothing** — silent failure.

**Rule:** Never use Markdown mode for messages that include user data (item descriptions, client names, vendor names). Either:
- Strip special chars before sending: `.replace('*','').replace('_','').replace('`','')...`
- Use `parse_mode=ParseMode.MARKDOWN` only for hardcoded template text, not dynamic data
- Always wrap `reply_text()` in try/except with a plain-text fallback

### 2. Telegram Message Length Limit
**Problem:** Telegram rejects messages over 4096 characters. For searches returning many results (e.g. 78 pork items), the full list exceeds this limit and the send fails silently.

**Rule:** Always chunk long messages at ~3800 chars (leave buffer). Use the pattern:
```python
if len(chunk) + len(line) > 3800:
    await update.message.reply_text(chunk.rstrip())
    chunk = line
else:
    chunk += line
await update.message.reply_text(chunk.rstrip())  # send final chunk
```

### 3. Railway Build Failures After GitHub Connection
**Problem:** When switching from Docker image source to GitHub repo, Railway uses `mise` to install Python. If `mise.toml` is missing or attestation is enabled, the build fails.

**Rule:** Always include `mise.toml` in the repo:
```toml
[settings]
python.github_attestations = false

[tools]
python = "3.11.9"
```

### 4. Webhook Sync Errors
**Problem:** After Railway redeploys, Telegram's webhook may have a `last_synchronization_error_date` indicating delivery failures. Some messages may be dropped.

**Rule:** After any Railway redeploy or domain change, reset the webhook:
```bash
curl -s "https://api.telegram.org/bot<TOKEN>/deleteWebhook"
curl -s --data-urlencode "url=https://inventory-bot-production-af17.up.railway.app/webhook" \
  --data "allowed_updates=[\"message\",\"callback_query\"]" \
  "https://api.telegram.org/bot<TOKEN>/setWebhook"
```

### 5. Stale Timestamp Display
**Problem:** `store.last_refresh` tracks when the bot's background thread last ran, not when the source data was actually updated. After a Railway restart, `last_refresh` resets and may show an old date from the last successful refresh before restart.

**Rule:** Always display `store.inventory_source_ts` (parsed from Google Drive Last-Modified header) as the authoritative data timestamp. Use `store.last_refresh` only for internal monitoring.

### 6. Code Lost Between Sessions
**Problem:** Bot code was only in the Railway Docker container — not in GitHub. When the Manus session ended, the code was lost and had to be recovered from the running container.

**Rule:** Always keep the bot code in GitHub (`dozse7en/Belcris-Telegram-Bot`). Railway must be connected to GitHub (not Docker image). Every change must be committed and pushed before the session ends.

---

## Deployment Checklist

Before pushing any change:
- [ ] Run `python3 -c "import ast; ast.parse(open('bot.py').read()); print('Syntax OK')"` to check for syntax errors
- [ ] Commit with a descriptive message
- [ ] Push to `main` branch
- [ ] Watch Railway Deployments tab for green build
- [ ] Test the affected command in Telegram

After Railway redeploy:
- [ ] Check bot responds to `/refresh`
- [ ] Reset webhook if messages are being dropped (see Lesson #4 above)

---

## File Structure

```
bot.py              — Main bot code (all commands, handlers, data loading)
requirements.txt    — Python dependencies
Procfile            — Railway start command: web: python bot.py
runtime.txt         — Python version: python-3.11.9
mise.toml           — mise config: disables GitHub attestation for Python install
.gitignore          — Excludes .env, __pycache__, *.db
README.md           — Full command reference and setup guide
CHANGELOG.md        — This file
```

---

## Commands Quick Reference

### Inventory
`/summary` `/search <kw>` `/check <code>` `/category [name]` `/expiring [days]` `/low [threshold]` `/lowstock [kw] [threshold]` `/warehouse <name>` `/top` `/components [kw]` `/refresh`

### Accounts Receivable
`/arsummary` `/ar <client>` `/client <client>` `/aging` `/overdue [days]` `/top30overdue` `/od30` `/area <name>` `/agent <name>` `/arsearch <kw>` `/arrefresh`

### Accounts Payable
`/apsummary` `/ap_summary` `/ap <vendor>` `/vendor <vendor>` `/apaging` `/apoverdue` `/aptop [n]` `/due_today` `/aprefresh`

### Special Behaviour
- Free text → inventory search
- `//message` → silently ignored (private notes)
- Mistyped commands → bot suggests correction
- Admin only: `/register <id1>,<id2> [note]` (batch), `/unregister <id>`, `/listusers`, `/seen`, `/accessmode [off|soft|hard]`, `/whitelistgroup [chat_id]`, `/unallowgroup <chat_id>`, `/allowedgroups`, `/seengroups`
- Whitelisted group chats are fully open for all members — no registration needed (recommended for the internal team chat where the boss is present)
- Access control modes: `off` (open, log only) | `soft` (warn) | `hard` (block)
