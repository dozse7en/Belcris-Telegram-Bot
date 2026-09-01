# Belcris Inventory Bot v3.0

A Telegram bot for Belcris Foods that provides real-time inventory, accounts receivable (AR), and accounts payable (AP) data pulled from Google Drive spreadsheets.

## Features

### 📦 Inventory
- `/refresh` — Reload all data from Google Drive
- `/summary` — Inventory snapshot (total stock, top items, expiry alerts)
- `/search <keyword>` — Search items by name
- `/check <item_code>` — Look up item by SAP code
- `/category [name]` — Browse by category (HAM, BACON, TOCINO…)
- `/expiring [days]` — Items expiring within N days (default 30)
- `/lowstock [keyword] [threshold]` — Items below threshold
- `/warehouse <name or code>` — Stock in a specific warehouse
- `/top` — Top 20 items by quantity
- `/components [keyword]` — Component Warehouse stock with expiry dates (PRD use)
- `/components_expiring [days]` — Components expiring within N days (default 30)

### 📋 Accounts Receivable
- `/arsummary` — AR executive summary + top clients & areas
- `/client <name or code>` — Full AR detail for a client
- `/aging` — AR aging breakdown by bucket
- `/overdue [days]` — Top overdue clients
- `/area <name>` — Receivables grouped by area
- `/agent <name>` — All clients under a sales agent
- `/arsearch <keyword>` — Search clients by name or code
- `/arrefresh` — Refresh AR data only

### 💳 Accounts Payable
- `/apsummary` — AP summary with aging breakdown
- `/ap <vendor>` — Vendor AP balance & transaction detail
- `/apaging` — AP aging by bucket
- `/apoverdue` — Vendors overdue 61+ days
- `/aptop [n]` — Top N vendors by outstanding amount
- `/due_today` — AP payments due this week
- `/aprefresh` — Refresh AP data only

## Deployment (Railway)

### Environment Variables
| Variable | Description |
|---|---|
| `BOT_TOKEN` | Telegram Bot API token |
| `RAILWAY_STATIC_URL` | Railway public domain (auto-set by Railway) |
| `INVENTORY_FILE_ID` | Google Drive file ID for inventory Excel |
| `AR_FILE_ID` | Google Drive file ID for AR Google Sheet |
| `AP_FILE_ID` | Google Drive file ID for AP Excel |
| `PORT` | Port to run Flask on (default 8080, auto-set by Railway) |
| `ADMIN_IDS` | Comma-separated Telegram user IDs with admin rights (required for `/register`, `/unregister`, `/listusers`, `/seen`, `/accessmode`) |
| `ACCESS_MODE` | Access enforcement mode: `off` (open, IDs logged only — default) / `soft` (warn unregistered, still allow) / `hard` (block unregistered from data commands) |
| `ALLOWED_GROUP_CHAT_IDS` | Comma-separated negative group Chat IDs that are fully open for all members regardless of `ACCESS_MODE` (recommended for the internal team chat) |

### Access Control (v4.0)
Every user who interacts with the bot is passively logged to a local SQLite database (`access_control.db`). Admins can then register users in bulk:
1. Send `/seen` to list users who have used the bot but are not registered.
2. Batch register them: `/register <id1>,<id2>,<id3> Sales Team` — the bot shows each user's known name/username from the log so you can verify.
3. Manage with `/listusers`, `/unregister <id>`, and `/accessmode`.
Admin rights require your Telegram user ID (from `/myid`) to be in the `ADMIN_IDS` variable. Enforcement only activates when `ACCESS_MODE` is set to `soft` or `hard`.

**Group chats (v4.1):** whitelisted group chats are fully open for all members — no registration, no notices, no blocking. Ideal for the internal team chat (e.g. where the boss is present): once you send `/whitelistgroup` in that chat, everyone in it can use the bot freely. Use `/seengroups` to list all groups the bot has been used in, and `/allowedgroups` to see which are open. You can also pre-open groups via the `ALLOWED_GROUP_CHAT_IDS` Railway variable (negative Chat IDs).

### Setup
1. Push this repo to GitHub
2. Connect the GitHub repo to Railway
3. Set the environment variables above
4. Railway will auto-deploy on every push

## Architecture
- **Python 3.11** + Flask (webhook mode)
- **python-telegram-bot 21.6** for Telegram API
- **openpyxl** for reading Excel/Google Sheets files
- Data refreshed every 30 minutes automatically
- Webhook endpoint: `POST /webhook`
- Health check: `GET /`
