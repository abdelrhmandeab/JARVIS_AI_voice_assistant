"""Phase 2 — MemoryManager: fast (RAM) vs LLM memory context.

Wraps the existing ``session_memory`` singleton (SQLite-backed slots + turns)
and its vector store. Does not introduce a second store — every accessor here
maps onto real ``SessionMemory`` methods.

``get_fast_context()`` is the router's hot path: RAM slots only, no vector
lookup, no LLM. ``get_llm_context()`` additionally pulls recent turns and
semantic recall for LLM/uncertain routes.
"""

from __future__ import annotations

from core.config import (
    MEMORY_FAST_CONTEXT_ENABLED,
    MEMORY_LLM_CONTEXT_ENABLED,
    MEMORY_PREFERENCES_ENABLED,
    MEMORY_SHORT_TERM_TURNS,
    MEMORY_VECTOR_RECALL_ENABLED,
    MEMORY_VECTOR_RECALL_MAX_RESULTS,
    MEMORY_VECTOR_RECALL_MIN_QUERY_WORDS,
)
from core.logger import get_logger
from core.memory_types import MemoryContext
from core.metrics import stage_timer
from core.session_memory import session_memory

logger = get_logger("memory")


class MemoryManager:
    def get_fast_context(self) -> MemoryContext:
        """RAM-only working slots. No vector store access, no LLM. Target < 5ms."""
        with stage_timer("memory_fast"):
            if not MEMORY_FAST_CONTEXT_ENABLED:
                return MemoryContext()
            snapshot = session_memory.context_snapshot()
            working_slots = {
                "last_app": snapshot.get("last_app", ""),
                "previous_app": snapshot.get("previous_app", ""),
                "last_file": snapshot.get("last_file", ""),
                "pending_confirmation_token": snapshot.get("pending_confirmation_token", ""),
                "language_history": snapshot.get("language_history", []),
                "response_mode": snapshot.get("response_mode", "default"),
            }
            preferences = {}
            if MEMORY_PREFERENCES_ENABLED:
                try:
                    preferences = session_memory.get_user_preferences()
                except Exception as exc:
                    logger.debug("get_fast_context: preferences unavailable: %s", exc)
            return MemoryContext(working_slots=working_slots, preferences=preferences)

    def get_llm_context(self, query: str, language: str = "en") -> MemoryContext:
        """Fast context + recent turns + bounded semantic recall.

        Only meant for LLM/uncertain routes — never called on the direct
        command hot path.
        """
        with stage_timer("memory_llm"):
            if not MEMORY_LLM_CONTEXT_ENABLED:
                return self.get_fast_context()

            context = self.get_fast_context()

            try:
                context.recent_turns = session_memory.recent(limit=MEMORY_SHORT_TERM_TURNS)
            except Exception as exc:
                logger.debug("get_llm_context: recent turns unavailable: %s", exc)
                context.recent_turns = []

            query_words = len(str(query or "").split())
            if MEMORY_VECTOR_RECALL_ENABLED and query_words >= MEMORY_VECTOR_RECALL_MIN_QUERY_WORDS:
                with stage_timer("vector_recall"):
                    try:
                        context.semantic_hits = session_memory.recall_semantic(
                            query, n=MEMORY_VECTOR_RECALL_MAX_RESULTS, language=language
                        )
                    except Exception as exc:
                        logger.debug("get_llm_context: semantic recall unavailable: %s", exc)
                        context.semantic_hits = []
            else:
                context.semantic_hits = []

            return context


memory_manager = MemoryManager()
