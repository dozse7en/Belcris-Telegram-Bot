"""Full-module smoke test for bot.py v4.0.

Imports the real bot.py with:
- BOT_TOKEN set to a dummy (the bot registers the webhook and Flask starts
  serving, so we patch app.run after import to avoid binding a port)
- Google Drive HTTP mocked

Asserts the access-control helpers, admin commands, and handlers exist and work.
"""
import os
import sys
import unittest
from unittest import mock

os.environ["BOT_TOKEN"] = "123456:SMOKE-TEST-TOKEN"
os.environ["ACCESS_DB_PATH"] = "/tmp/belcris_smoke_access.db"
os.environ["ADMIN_IDS"] = "999999999,888888888"
os.environ["ACCESS_MODE"] = "soft"

if os.path.exists("/tmp/belcris_smoke_access.db"):
    os.remove("/tmp/belcris_smoke_access.db")


class SmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with mock.patch("requests.get"), mock.patch("requests.post"):
            # Prevent Flask from blocking in app.run()
            with mock.patch("flask.Flask.run"):
                import importlib.util
                spec = importlib.util.spec_from_file_location(
                    "bot", "/home/ubuntu/Belcris-Telegram-Bot/bot.py")
                cls.bot = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(cls.bot)

    def test_admin_ids_parsed(self):
        self.assertEqual(self.bot.ADMIN_IDS, {999999999, 888888888})

    def test_access_mode(self):
        self.assertEqual(self.bot.ACCESS_MODE, "soft")

    def test_admin_commands_registered(self):
        app = self.bot.tg_app
        handler_names = []
        for handler_list in app.handlers.values():
            for h in handler_list:
                handler_names.append(h.callback.__name__ if h.callback else "unknown")
        for cmd in ("register", "unregister", "listusers", "seen", "accessmode"):
            self.assertIn(f"cmd_{cmd}", handler_names, f"{cmd} handler not registered")

    def test_helpers_exist(self):
        for name in ("log_user", "is_registered", "register_user", "unregister_user",
                     "list_registered_users", "list_seen_users", "_is_admin",
                     "_access_gate", "_send_blocked_notice", "_parse_id_list"):
            self.assertTrue(callable(getattr(self.bot, name)), name)

    def test_log_and_register_flow(self):
        b = self.bot
        b.log_user(111, "alice", "Alice A")
        b.log_user(222, "bob", "Bob B")
        seen = b.list_seen_users(10000)
        self.assertEqual(len(seen), 2)
        self.assertTrue(b.register_user(111, 999999999, "smoke test"))
        self.assertFalse(b.register_user(111, 999999999))
        self.assertFalse(b.is_registered(222))
        self.assertTrue(b.is_registered(111))
        users = b.list_registered_users()
        self.assertEqual(len(users), 1)
        self.assertEqual(users[0]["telegram_user_id"], 111)
        registered = {r["telegram_user_id"] for r in b.list_registered_users()}
        unreg = [r for r in b.list_seen_users(10000) if r["telegram_user_id"] not in registered]
        self.assertEqual(len(unreg), 1)
        self.assertEqual(unreg[0]["telegram_user_id"], 222)
        self.assertTrue(b.unregister_user(111))
        self.assertFalse(b.is_registered(111))


if __name__ == "__main__":
    unittest.main(verbosity=2)
