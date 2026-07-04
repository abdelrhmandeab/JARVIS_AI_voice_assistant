import os
import unittest

from nlp.code_switch_router import try_codeswitch
from os_control.path_resolver import KNOWN_FOLDERS


class CodeswitchRoutingTests(unittest.TestCase):
    """Asserts Phase 4's code-switch shortcut resolves mixed EN/AR
    verb+entity utterances to the right intent+slots without the semantic
    embedding tier running.
    """

    def test_arabic_verb_english_app_resolves_open(self):
        parsed = try_codeswitch("افتح Chrome", "ar")
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.intent, "OS_APP_OPEN")
        self.assertEqual(parsed.args.get("app_name"), "chrome")

    def test_arabic_verb_english_entity_resolves_volume_up(self):
        parsed = try_codeswitch("زود volume", "ar")
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.intent, "OS_SYSTEM_COMMAND")
        self.assertEqual(parsed.args.get("action_key"), "volume_up")

    def test_close_verb_resolves_app_close(self):
        parsed = try_codeswitch("close Spotify", "en")
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.intent, "OS_APP_CLOSE")
        self.assertEqual(parsed.args.get("app_name"), "spotify")

    def test_open_folder_entity_resolves_file_navigation(self):
        parsed = try_codeswitch("open downloads", "en")
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.intent, "OS_FILE_NAVIGATION")
        self.assertEqual(parsed.action, "cd")

    def test_ambiguous_text_falls_through(self):
        parsed = try_codeswitch("random ambiguous text here", "en")
        self.assertIsNone(parsed)


class CodeswitchFileOpsTests(unittest.TestCase):
    """Delete/move/rename need a real file to resolve against (via
    os_control.file_ops.resolve_name_in_location across the default search
    roots) — unlike open/close, which only resolve app names/folder aliases.
    Creates and removes a throwaway fixture file on Desktop for each test.
    """

    def setUp(self):
        self.desktop = KNOWN_FOLDERS.get("Desktop")
        self.fixture_name = "jarvis_codeswitch_test_fixture"
        self.fixture_path = os.path.join(str(self.desktop), f"{self.fixture_name}.pdf")
        with open(self.fixture_path, "w", encoding="utf-8") as f:
            f.write("test")

    def tearDown(self):
        for suffix in (".pdf",):
            path = os.path.join(str(self.desktop), f"{self.fixture_name}{suffix}")
            if os.path.exists(path):
                os.remove(path)

    def test_arabic_delete_verb_not_in_parser_regex_still_resolves(self):
        # "احذف" is in core.intent_confidence._DELETE_VERBS but not in
        # command_parser's delete regex (only امسح/شيل/delete/remove) — the
        # code-switch shortcut is what catches this gap.
        parsed = try_codeswitch(f"احذف {self.fixture_name}", "ar")
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.intent, "OS_FILE_NAVIGATION")
        self.assertEqual(parsed.action, "delete_item")
        self.assertEqual(parsed.args.get("path"), self.fixture_path)

    def test_english_delete_no_extension_resolves(self):
        parsed = try_codeswitch(f"delete {self.fixture_name}", "en")
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.intent, "OS_FILE_NAVIGATION")
        self.assertEqual(parsed.action, "delete_item")
        self.assertEqual(parsed.args.get("path"), self.fixture_path)

    def test_mixed_rename_with_arabic_separator(self):
        parsed = try_codeswitch(f"rename {self.fixture_name} الى renamed.pdf", "ar")
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.intent, "OS_FILE_NAVIGATION")
        self.assertEqual(parsed.action, "rename_item")
        self.assertEqual(parsed.args.get("source"), self.fixture_path)
        self.assertEqual(parsed.args.get("new_name"), "renamed.pdf")

    def test_arabic_move_verb_english_destination(self):
        parsed = try_codeswitch(f"انقل {self.fixture_name} to Downloads", "ar")
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.intent, "OS_FILE_NAVIGATION")
        self.assertEqual(parsed.action, "move_item")
        self.assertEqual(parsed.args.get("source"), self.fixture_path)
        self.assertEqual(parsed.args.get("destination").lower(), "downloads")

    def test_delete_nonexistent_file_falls_through(self):
        parsed = try_codeswitch("احذف nonexistent_jarvis_test_xyz123", "ar")
        self.assertIsNone(parsed)


if __name__ == "__main__":
    unittest.main()
