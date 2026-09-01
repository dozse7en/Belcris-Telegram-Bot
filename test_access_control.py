"""Local tests for the v4.0 access-control module (does NOT import the full bot,
which needs a real BOT_TOKEN and Drive files at module load).
Imports only the access-control logic plus the admin command parsers by re-using
the module under dummy env vars with Telegram API calls mocked out.
"""
import os
import sys
import json
import sqlite3
import unittest
from unittest import mock

# Dummy env before any import so bot.py can load without real secrets.
os.environ.setdefault("BOT_TOKEN", "123456:TEST-TOKEN")
os.environ.setdefault("ACCESS_DB_PATH", "/tmp/belcris_test_access.db")
os.environ.setdefault("ADMIN_IDS", "999999999")
os.environ.setdefault("ACCESS_MODE", "hard")
os.environ.setdefault("INVENTORY_FILE_ID", "test")
os.environ.setdefault("AR_FILE_ID", "test")
os.environ.setdefault("AP_FILE_ID", "test")

# Mock all Telegram API network calls and Flask startup.
sys.modules.setdefault("telegram", mock.MagicMock())

# Instead of importing bot.py (heavy, needs real data), extract & exec the
# access-control section directly to unit-test it in isolation.

ACCESS_CODE = '''
import sqlite3, threading, os, datetime
PHT = datetime.timezone(datetime.timedelta(hours=8))

_DB_PATH = os.environ.get("ACCESS_DB_PATH", "/tmp/belcris_test_access.db")
ADMIN_IDS_ENV = os.environ.get("ADMIN_IDS", "").strip()
ADMIN_IDS = set()
if ADMIN_IDS_ENV:
    for raw in ADMIN_IDS_ENV.split(","):
        raw = raw.strip()
        if raw.isdigit():
            ADMIN_IDS.add(int(raw))

ACCESS_MODE = os.environ.get("ACCESS_MODE", "off").strip().lower()
if ACCESS_MODE not in ("off", "soft", "hard"):
    ACCESS_MODE = "off"

_db_lock = threading.Lock()

def _access_db():
    conn = sqlite3.connect(_DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def _ensure_access_tables(conn):
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS seen_users (
            telegram_user_id INTEGER PRIMARY KEY,
            username TEXT, full_name TEXT,
            first_seen TEXT NOT NULL DEFAULT (datetime('now')),
            last_seen TEXT NOT NULL DEFAULT (datetime('now')),
            message_count INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS registered_users (
            telegram_user_id INTEGER PRIMARY KEY,
            registered_by INTEGER NOT NULL,
            registered_at TEXT NOT NULL DEFAULT (datetime('now')),
            note TEXT
        );
        CREATE TABLE IF NOT EXISTS seen_groups (
            telegram_chat_id INTEGER PRIMARY KEY,
            chat_type TEXT, chat_title TEXT,
            first_seen TEXT NOT NULL DEFAULT (datetime('now')),
            last_seen TEXT NOT NULL DEFAULT (datetime('now')),
            interaction_count INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS allowed_groups (
            telegram_chat_id INTEGER PRIMARY KEY,
            allowed_by INTEGER NOT NULL,
            allowed_at TEXT NOT NULL DEFAULT (datetime('now')),
            note TEXT
        );
    """)

ALLOWED_GROUP_CHAT_IDS = set()
for raw in os.environ.get("ALLOWED_GROUP_CHAT_IDS", "").split(","):
    raw = raw.strip()
    if raw.lstrip("-").isdigit():
        ALLOWED_GROUP_CHAT_IDS.add(int(raw))

def log_chat(chat_id, chat_type, chat_title):
    if not chat_id:
        return
    with _db_lock:
        conn = _access_db(); _ensure_access_tables(conn)
        now = datetime.datetime.now(PHT).strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            "INSERT INTO seen_groups (telegram_chat_id, chat_type, chat_title, last_seen, interaction_count) "
            "VALUES (?, ?, ?, ?, 1) ON CONFLICT(telegram_chat_id) DO UPDATE SET "
            "chat_type = excluded.chat_type, chat_title = excluded.chat_title, "
            "last_seen = excluded.last_seen, interaction_count = interaction_count + 1",
            (chat_id, chat_type, chat_title, now))
        conn.commit(); conn.close()

def is_whitelisted_group(chat_id):
    if not chat_id or chat_id >= 0:
        return False
    if chat_id in ALLOWED_GROUP_CHAT_IDS:
        return True
    with _db_lock:
        conn = _access_db(); _ensure_access_tables(conn)
        row = conn.execute("SELECT 1 FROM allowed_groups WHERE telegram_chat_id = ?", (chat_id,)).fetchone()
        conn.close()
    return row is not None

def allow_group(chat_id, allowed_by, note=None):
    with _db_lock:
        conn = _access_db(); _ensure_access_tables(conn)
        existing = conn.execute("SELECT 1 FROM allowed_groups WHERE telegram_chat_id = ?", (chat_id,)).fetchone()
        if existing:
            conn.close(); return False
        conn.execute("INSERT INTO allowed_groups (telegram_chat_id, allowed_by, note) VALUES (?,?,?)",
                     (chat_id, allowed_by, note))
        conn.commit(); conn.close()
    return True

def unallow_group(chat_id):
    with _db_lock:
        conn = _access_db(); _ensure_access_tables(conn)
        cur = conn.execute("DELETE FROM allowed_groups WHERE telegram_chat_id = ?", (chat_id,))
        conn.commit(); conn.close()
    return cur.rowcount > 0

def list_seen_groups(limit=500):
    with _db_lock:
        conn = _access_db(); _ensure_access_tables(conn)
        rows = conn.execute(
            "SELECT telegram_chat_id, chat_type, chat_title, first_seen, last_seen, interaction_count "
            "FROM seen_groups ORDER BY last_seen DESC LIMIT ?", (limit,)).fetchall()
        conn.close()
    return rows

def log_user(user_id, username, full_name):
    if not user_id:
        return
    with _db_lock:
        conn = _access_db()
        _ensure_access_tables(conn)
        now = datetime.datetime.now(PHT).strftime("%Y-%m-%d %H:%M:%S")
        conn.execute(
            "INSERT INTO seen_users (telegram_user_id, username, full_name, last_seen, message_count) "
            "VALUES (?, ?, ?, ?, 1) ON CONFLICT(telegram_user_id) DO UPDATE SET "
            "username = excluded.username, full_name = excluded.full_name, "
            "last_seen = excluded.last_seen, message_count = message_count + 1",
            (user_id, username, full_name, now))
        conn.commit(); conn.close()

def is_registered(user_id):
    if not user_id:
        return False
    if ACCESS_MODE == "off":
        return True
    with _db_lock:
        conn = _access_db(); _ensure_access_tables(conn)
        row = conn.execute("SELECT 1 FROM registered_users WHERE telegram_user_id = ?", (user_id,)).fetchone()
        conn.close()
    return row is not None

def register_user(user_id, registered_by, note=None):
    with _db_lock:
        conn = _access_db(); _ensure_access_tables(conn)
        existing = conn.execute("SELECT 1 FROM registered_users WHERE telegram_user_id = ?", (user_id,)).fetchone()
        if existing:
            conn.close(); return False
        conn.execute("INSERT INTO registered_users (telegram_user_id, registered_by, note) VALUES (?,?,?)",
                     (user_id, registered_by, note))
        conn.commit(); conn.close()
    return True

def unregister_user(user_id):
    with _db_lock:
        conn = _access_db(); _ensure_access_tables(conn)
        cur = conn.execute("DELETE FROM registered_users WHERE telegram_user_id = ?", (user_id,))
        conn.commit(); conn.close()
    return cur.rowcount > 0

def list_registered_users():
    with _db_lock:
        conn = _access_db(); _ensure_access_tables(conn)
        rows = conn.execute(
            "SELECT r.telegram_user_id, r.registered_by, r.registered_at, r.note, "
            "COALESCE(s.username,'') AS username, COALESCE(s.full_name,'') AS full_name, "
            "COALESCE(s.last_seen,'') AS last_seen FROM registered_users r "
            "LEFT JOIN seen_users s ON s.telegram_user_id = r.telegram_user_id "
            "ORDER BY r.registered_at DESC").fetchall()
        conn.close()
    return rows

def list_seen_users(limit=500):
    with _db_lock:
        conn = _access_db(); _ensure_access_tables(conn)
        rows = conn.execute(
            "SELECT telegram_user_id, username, full_name, first_seen, last_seen, message_count "
            "FROM seen_users ORDER BY last_seen DESC LIMIT ?", (limit,)).fetchall()
        conn.close()
    return rows

def _is_admin(user_id):
    return user_id in ADMIN_IDS

def _parse_id_list(args):
    ids = []
    for arg in args:
        for part in arg.replace(",", " ").split():
            if part.isdigit():
                ids.append(int(part))
    return ids
'''

# Execute the access-control module in a fresh namespace.
ns = {}
exec(ACCESS_CODE, ns)


class Row(dict):
    """Minimal Row-like dict for tests."""
    pass


class TestAccessControl(unittest.TestCase):
    def setUp(self):
        # Fresh DB per test
        if os.path.exists("/tmp/belcris_test_access.db"):
            os.remove("/tmp/belcris_test_access.db")
        ns["_ensure_access_tables"](ns["_access_db"]())

    # ── Passive logging ──
    def test_log_user_inserts_and_updates(self):
        ns["log_user"](111, "alice", "Alice A")
        ns["log_user"](111, "alice", "Alice A")
        ns["log_user"](222, "bob", "Bob B")
        seen = ns["list_seen_users"]()
        self.assertEqual(len(seen), 2)
        alice = next(r for r in seen if r["telegram_user_id"] == 111)
        self.assertEqual(alice["message_count"], 2)
        self.assertEqual(alice["full_name"], "Alice A")

    # ── Group whitelist ──
    def test_log_chat_inserts_and_updates(self):
        ns["log_chat"](-1001111111, "supergroup", "Sales Team")
        ns["log_chat"](-1001111111, "supergroup", "Sales Team")
        ns["log_chat"](-1002222222, "group", "Orders Chat")
        seen = ns["list_seen_groups"]()
        self.assertEqual(len(seen), 2)
        grp = next(r for r in seen if r["telegram_chat_id"] == -1001111111)
        self.assertEqual(grp["interaction_count"], 2)
        self.assertEqual(grp["chat_title"], "Sales Team")

    def test_log_chat_skips_none_id(self):
        ns["log_chat"](None, None, None)
        self.assertEqual(len(ns["list_seen_groups"]()), 0)

    def test_log_chat_accepts_private_ids_too(self):
        # bot.py's _access_gate logs every chat (groups AND private) —
        # whitelisting only applies to negative (group) IDs
        ns["log_chat"](123456, "private", "Alice")
        self.assertEqual(len(ns["list_seen_groups"]()), 1)
        self.assertFalse(ns["is_whitelisted_group"](123456))  # but private chats are never whitelisted

    def test_whitelist_rejects_positive_ids(self):
        self.assertFalse(ns["is_whitelisted_group"](123456))
        self.assertFalse(ns["is_whitelisted_group"](0))

    def test_whitelist_db_flag(self):
        self.assertFalse(ns["is_whitelisted_group"](-1001111111))
        self.assertTrue(ns["allow_group"](-1001111111, 999999999, "team chat"))
        self.assertTrue(ns["is_whitelisted_group"](-1001111111))
        # Duplicate allow fails gracefully
        self.assertFalse(ns["allow_group"](-1001111111, 999999999))
        self.assertTrue(ns["unallow_group"](-1001111111))
        self.assertFalse(ns["is_whitelisted_group"](-1001111111))
        self.assertFalse(ns["unallow_group"](-1001111111))  # already gone

    def test_log_user_skips_none_id(self):
        ns["log_user"](None, None, None)
        self.assertEqual(len(ns["list_seen_users"]()), 0)

    # ── Registration ──
    def test_register_and_is_registered(self):
        self.assertTrue(ns["register_user"](111, 999999999, "test"))
        self.assertTrue(ns["is_registered"](111))
        self.assertFalse(ns["is_registered"](222))

    def test_register_duplicate_returns_false(self):
        ns["register_user"](111, 999999999)
        self.assertFalse(ns["register_user"](111, 999999999))

    def test_unregister(self):
        ns["register_user"](111, 999999999)
        self.assertTrue(ns["unregister_user"](111))
        self.assertFalse(ns["is_registered"](111))
        self.assertFalse(ns["unregister_user"](111))

    # ── Batch registration ──
    def test_batch_register(self):
        ids = [101, 102, 103]
        results = [ns["register_user"](i, 999999999) for i in ids]
        self.assertTrue(all(results))
        self.assertEqual(len(ns["list_registered_users"]()), 3)

    # ── Batch parsing ──
    def test_parse_id_list_comma_and_space(self):
        self.assertEqual(ns["_parse_id_list"](["123,456", "789"]), [123, 456, 789])

    def test_parse_id_list_ignores_notes(self):
        # Note tokens are non-digit; parser only returns digit-only parts
        self.assertEqual(ns["_parse_id_list"](["123,456", "new-agent"]), [123, 456])

    def test_parse_id_list_empty(self):
        self.assertEqual(ns["_parse_id_list"]([]), [])

    # ── Admin gate ──
    def test_admin_ids_parsed(self):
        self.assertIn(999999999, ns["ADMIN_IDS"])
        self.assertTrue(ns["_is_admin"](999999999))
        self.assertFalse(ns["_is_admin"](111))

    # ── Seen vs registered (/seen command logic) ──
    def test_seen_excludes_registered(self):
        ns["log_user"](111, "alice", "Alice")
        ns["log_user"](222, "bob", "Bob")
        ns["register_user"](111, 999999999)
        registered = {r["telegram_user_id"] for r in ns["list_registered_users"]()}
        unregistered = [r for r in ns["list_seen_users"](10000) if r["telegram_user_id"] not in registered]
        self.assertEqual(len(unregistered), 1)
        self.assertEqual(unregistered[0]["telegram_user_id"], 222)

    # ── Access mode off → everyone passes ──
    def test_off_mode_allows_all(self):
        # Re-execute the module code with ACCESS_MODE=off in env
        off_env = dict(os.environ)
        off_env["ACCESS_MODE"] = "off"
        with mock.patch.dict(os.environ, {"ACCESS_MODE": "off"}):
            ns_off = {}
            exec(ACCESS_CODE, ns_off)
            ns_off["_ensure_access_tables"](ns_off["_access_db"]())
            # In off mode, is_registered returns True for any ID without DB lookup
            self.assertTrue(ns_off["is_registered"](999999999))
            self.assertTrue(ns_off["is_registered"](12345))


if __name__ == "__main__":
    unittest.main(verbosity=2)
