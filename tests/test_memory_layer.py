import unittest
from unittest.mock import patch

from core.memory_manager import memory_manager
from core.session_memory import session_memory


class MemoryFastContextTests(unittest.TestCase):
    """Asserts the router's hot path (get_fast_context) reflects RAM slot
    state and never touches the vector store — Phase 2's core guarantee.
    """

    def setUp(self):
        self._saved_last_app = session_memory.get_last_app()

    def tearDown(self):
        if self._saved_last_app:
            session_memory.set_last_app(self._saved_last_app)

    def test_fast_context_reflects_last_app(self):
        session_memory.set_last_app("Notepad")
        ctx = memory_manager.get_fast_context()
        self.assertEqual(ctx.working_slots.get("last_app"), "Notepad")

    def test_fast_context_never_touches_vector_store(self):
        with patch("core.memory_store.vector_memory.recall") as mock_recall, \
             patch("core.memory_store.vector_memory.remember") as mock_remember:
            memory_manager.get_fast_context()
            mock_recall.assert_not_called()
            mock_remember.assert_not_called()

    def test_fast_context_under_latency_ceiling(self):
        """Warn (don't fail) between 5-15ms to tolerate slow CI; hard-fail above 15ms."""
        import time

        durations = []
        for _ in range(5):
            start = time.perf_counter()
            memory_manager.get_fast_context()
            durations.append((time.perf_counter() - start) * 1000.0)
        best = min(durations)
        if best > 15.0:
            self.fail(f"get_fast_context() took {best:.2f}ms (ceiling 15ms): {durations}")
        elif best > 5.0:
            print(f"WARNING: get_fast_context() best-of-5 was {best:.2f}ms (target <5ms)")


class MemoryLlmContextTests(unittest.TestCase):
    def test_llm_context_includes_recent_turns(self):
        session_memory.add_turn(
            "what is the capital of Egypt", "Cairo is the capital of Egypt.",
            language="en", intent="LLM_QUERY",
        )
        ctx = memory_manager.get_llm_context("tell me more about that", language="en")
        self.assertTrue(any(t.get("user") for t in ctx.recent_turns))

    def test_llm_context_short_query_skips_semantic_recall(self):
        with patch("core.session_memory.vector_memory.recall") as mock_recall:
            memory_manager.get_llm_context("hi", language="en")
            mock_recall.assert_not_called()


if __name__ == "__main__":
    unittest.main()
