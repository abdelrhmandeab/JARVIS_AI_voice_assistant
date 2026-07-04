"""Phase 2 — memory context shapes shared by the fast (RAM) and LLM paths.

``MemoryContext`` is the single value object ``MemoryManager`` returns. The
router's fast path only ever touches ``working_slots``; the LLM path also
gets ``recent_turns``, ``preferences``, ``pending_task``, and
``semantic_hits``.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class MemoryContext:
    working_slots: dict = field(default_factory=dict)
    recent_turns: list = field(default_factory=list)
    preferences: dict = field(default_factory=dict)
    pending_task: dict = field(default_factory=dict)
    command_hints: dict = field(default_factory=dict)
    semantic_hits: list = field(default_factory=list)

    def compact_for_router(self) -> dict:
        """Minimal dict for fast-path consumers (no turns/vector data)."""
        return {
            "last_app": self.working_slots.get("last_app", ""),
            "previous_app": self.working_slots.get("previous_app", ""),
            "last_file": self.working_slots.get("last_file", ""),
            "pending_confirmation_token": self.working_slots.get("pending_confirmation_token", ""),
            "language_history": self.working_slots.get("language_history", []),
            "response_mode": self.working_slots.get("response_mode", "default"),
            "preferences": self.preferences,
            "pending_task": self.pending_task,
        }

    def to_prompt_block(self, max_chars: int = 900) -> str:
        """Render a single bounded ``MEMORY CONTEXT:`` block for the LLM prompt."""
        lines = []

        last_app = self.working_slots.get("last_app")
        if last_app:
            lines.append(f"- last app opened: {last_app}")
        last_file = self.working_slots.get("last_file")
        if last_file:
            lines.append(f"- last file referenced: {last_file}")
        pending_confirmation = self.working_slots.get("pending_confirmation_token")
        if pending_confirmation:
            lines.append(f"- pending confirmation: {pending_confirmation}")

        if self.preferences:
            pref_line = ", ".join(f"{k}={v}" for k, v in self.preferences.items())
            lines.append(f"- user preferences: {pref_line}")

        if self.pending_task:
            lines.append(f"- pending task: {self.pending_task}")

        if self.recent_turns:
            lines.append("- recent conversation:")
            for turn in self.recent_turns:
                user_text = str(turn.get("user") or "").strip()
                assistant_text = str(turn.get("assistant") or "").strip()
                if user_text:
                    lines.append(f"  user: {user_text}")
                if assistant_text:
                    lines.append(f"  assistant: {assistant_text}")

        if self.semantic_hits:
            lines.append("- relevant past exchanges:")
            for hit in self.semantic_hits:
                user_text = str(hit.get("user") or "").strip()
                assistant_text = str(hit.get("assistant") or "").strip()
                if user_text:
                    lines.append(f"  user: {user_text}")
                if assistant_text:
                    lines.append(f"  assistant: {assistant_text}")

        if not lines:
            return ""

        block = "MEMORY CONTEXT:\n" + "\n".join(lines)
        if len(block) > max_chars:
            block = block[:max_chars].rstrip()
        return block
