import os
import unittest

from core.command_router import _rewrite_followup_command
from core.session_memory import session_memory
from os_control.path_resolver import KNOWN_FOLDERS


class ReferenceResolverTests(unittest.TestCase):
    """Phase 3's follow-up resolver ("close it" -> last_app, etc.) — RAM-only,
    bilingual, freshness-gated. See core.command_router._rewrite_followup_command.
    """

    def setUp(self):
        self._saved_last_app = session_memory.get_last_app()
        self._saved_last_file = session_memory.get_last_file()
        self.desktop = KNOWN_FOLDERS.get("Desktop")
        self.fixture_path = os.path.join(str(self.desktop), "jarvis_ref_resolver_fixture.txt")
        with open(self.fixture_path, "w", encoding="utf-8") as f:
            f.write("test")

    def tearDown(self):
        if os.path.exists(self.fixture_path):
            os.remove(self.fixture_path)
        if self._saved_last_app:
            session_memory.set_last_app(self._saved_last_app)
        if self._saved_last_file:
            session_memory.set_last_file(self._saved_last_file)

    def test_open_chrome_then_close_it_resolves_to_last_app(self):
        session_memory.set_last_app("Chrome")
        rewritten, meta = _rewrite_followup_command("close it", language="en")
        self.assertEqual(rewritten, "close app Chrome")
        self.assertEqual(meta.get("followup_rewrite"), "close_last_app")

    def test_find_file_then_open_it_resolves_to_last_file(self):
        session_memory.set_last_file(self.fixture_path)
        rewritten, meta = _rewrite_followup_command("open it", language="en")
        self.assertIn(self.fixture_path, rewritten)
        self.assertIn(meta.get("followup_rewrite"), {"open_last_file", "file_info_last_file"})

    def test_delete_that_file_with_fresh_referent_resolves(self):
        # "delete that file" is an explicit reference (mentions "file"), unlike
        # "delete it" — see _DELETE_LAST_FILE_FOLLOWUP_TEXTS vs the vague set below.
        session_memory.set_last_file(self.fixture_path)
        rewritten, meta = _rewrite_followup_command("delete that file", language="en")
        self.assertEqual(rewritten, f"delete {self.fixture_path}")
        self.assertEqual(meta.get("followup_rewrite"), "delete_last_file")
        # The resolver only rewrites text to a delete command — the PIN gate
        # itself lives downstream in risk_policy/file_ops, not in the resolver.

    def test_delete_it_vague_pronoun_is_blocked_even_with_fresh_referent(self):
        # By default FOLLOWUP_DESTRUCTIVE_REQUIRE_EXPLICIT_REFERENCE=True blocks
        # bare pronoun references ("delete it"/"delete this") for destructive
        # actions even when a fresh last_file exists — safety over convenience.
        session_memory.set_last_file(self.fixture_path)
        rewritten, meta = _rewrite_followup_command("delete it", language="en")
        self.assertEqual(rewritten, "delete it")
        self.assertTrue(meta.get("followup_blocked"))

    def test_close_it_with_no_last_app_is_unchanged(self):
        # set_last_app("") is a no-op by design (no public "clear a single
        # slot" API), so reach into the private slot directly for this test.
        session_memory._context_slots["last_app"] = ""
        session_memory._context_slots["last_app_updated_at"] = 0.0
        rewritten, meta = _rewrite_followup_command("close it", language="en")
        self.assertEqual(rewritten, "close it")
        self.assertTrue(meta.get("followup_blocked"))

    def test_arabic_eftah_resolves_to_last_file(self):
        session_memory.set_last_file(self.fixture_path)
        rewritten, meta = _rewrite_followup_command("افتحه", language="ar")
        self.assertIn(self.fixture_path, rewritten)


if __name__ == "__main__":
    unittest.main()
