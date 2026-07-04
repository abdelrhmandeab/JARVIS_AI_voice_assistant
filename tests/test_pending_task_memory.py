import time
import unittest
from unittest.mock import patch

from core.command_router import _advance_pending_task, _start_pending_task
from core.session_memory import session_memory


class PendingTaskMemoryTests(unittest.TestCase):
    """Phase 4's multi-turn slot filling (email compose: to -> body -> subject),
    RAM-only via session_memory.set_pending_task/get_pending_task, no LLM used
    to run the flow. See core.command_router._start_pending_task/_advance_pending_task.
    """

    def setUp(self):
        session_memory.clear_pending_task()

    def tearDown(self):
        session_memory.clear_pending_task()

    def test_email_multi_turn_slot_fill_dispatches_after_all_slots_filled(self):
        missing_slot, question = _start_pending_task(
            "OS_EMAIL", "draft", {"to": "ahmed", "subject": "", "body": ""}, "en"
        )
        self.assertEqual(missing_slot, "body")
        self.assertIn("say", question.lower())

        handled, response = _advance_pending_task("the report is ready", "en")
        self.assertTrue(handled)
        self.assertIn("subject", response.lower())
        self.assertIsNotNone(session_memory.get_pending_task())

        with patch("core.command_router.draft_email", return_value="Opening Outlook to ahmed.") as mock_draft:
            handled2, response2 = _advance_pending_task("graduation update", "en")
        self.assertTrue(handled2)
        mock_draft.assert_called_once_with(to="ahmed", subject="graduation update", body="the report is ready", language="en")
        self.assertIsNone(session_memory.get_pending_task())

    def test_pending_task_ttl_expiry_clears_task(self):
        session_memory.set_pending_task(
            {"intent": "OS_EMAIL", "action": "draft", "args": {"to": "ahmed"}},
            ttl_seconds=1,
        )
        task = session_memory.get_pending_task()
        self.assertIsNotNone(task)
        time.sleep(task["ttl_seconds"] + 1)
        self.assertIsNone(session_memory.get_pending_task())

    def test_advance_with_no_pending_task_returns_not_handled(self):
        handled, response = _advance_pending_task("anything", "en")
        self.assertFalse(handled)


if __name__ == "__main__":
    unittest.main()
