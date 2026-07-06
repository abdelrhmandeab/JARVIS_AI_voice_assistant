# Cleanup Report

Record of dead code, unused dependencies, and orphaned files removed from the
repository, so deletions stay auditable. Entries are dated; each records what
was removed, why it was confirmed dead (not just suspected), and what was
deliberately kept.

## 2026-07-06 - Final pre-defense sweep (Phase 7)

### Confirmed already clean (no action needed)

- `.swarm/`, `jarvis.log.1`/`jarvis.log.3`, `jarvis_doc_updated.html`, and
  duplicate `.docx` books were already absent from the working tree and from
  git history — a prior session had already removed them. Added `.swarm/`
  and `jarvis.log*` to `.gitignore` as a guard against recurrence.
- `vulture --min-confidence 60` (and 80) found zero dead code repo-wide.
- `scripts/` and `tools/` contents are all referenced by the doc generator,
  CHANGELOG, tests, or `core/command_router.py` — no orphaned one-off scripts.
- `desktop/src-tauri/target/` has empty directories (Cargo build cache) but
  they are already gitignored and untracked — not a repo cleanliness issue.

### Removed: dead hex-token confirmation system
The spoken-PIN system (`SENSITIVE_CONFIRM_MODE=pin`) fully replaced an older
hex-token confirmation flow, but the old flow's code was never deleted.
Verified dead by tracing the call graph: `os_control/persistence.py`'s
`store_confirmation()` — the only function that ever wrote a token — had zero
callers anywhere in the codebase, so every reader of that table
(`confirm()`, `confirm_with_second_factor()`, token-based `cancel()`) was
unreachable in practice. Confirmed empirically by testing
`_rewrite_followup_command()` under a live pending-PIN state: every input
pattern that would have built a `confirm <token> ...` string is intercepted
earlier by the PIN-required branch, so those branches never ran.

Removed:
- `os_control/confirmation.py`: `confirm()`, `confirm_with_second_factor()`,
  `_check_confirmation_rate_limit()`, `cancel()`, `pending_count()`.
- `os_control/persistence.py`: `store_confirmation`, `get_confirmation`,
  `delete_confirmation`, `pop_confirmation`, `consume_confirmation`,
  `cleanup_expired_confirmations`, `count_pending_confirmations`.
- `core/command_parser.py`: the `confirm <hex-token>` regex/`OS_CONFIRMATION`
  intent emission, `_CONFIRMATION_TOKEN_MAX_HEX_LEN`.
- `core/command_router.py`: the `OS_CONFIRMATION` dispatch branch, its
  `_is_sensitive_command`/`_PARSER_FASTPATH_INTENTS`/`_PERMISSION_MAP`
  entries, and the dead `if pending_token:` sub-branches of
  `_rewrite_followup_command` (their live "no pending token" siblings and the
  PIN-required early-return were kept).
- `core/config.py`: `CONFIRMATION_TOKEN_BYTES`, `CONFIRMATION_TOKEN_MIN_HEX_LEN`.

Kept (still load-bearing for the live PIN flow): `CONFIRMATION_TIMEOUT_SECONDS`
(used by `format_confirmation_prompt`'s legacy-shaped signature, still called
from `app_ops.py`/`system_ops.py`/`file_ops.py`), `CONFIRMATION_MAX_ATTEMPTS_PER_TOKEN`
and `CONFIRMATION_LOCKOUT_SECONDS` (used by `second_factor.py`'s rate limiter,
which the PIN flow also uses), the `confirmations` SQLite table schema
(harmless unused table, not worth a migration during a code cleanup pass).

Bug found and fixed in passing: `_update_short_term_context`'s "record last
app/file after a confirmed destructive action" bookkeeping checked
`parsed.intent == "OS_CONFIRMATION"`, which could never fire — so it silently
never ran under the PIN system. Changed to `OS_PIN_CONFIRM`, its live
equivalent (same `_execute_confirmed_payload()` call, same `exec_meta` shape).
Verified with a direct test: a PIN-confirmed app close now correctly records
the closed app as the last-used app for follow-up references.

### Removed: broken/unimplemented Claude LLM backend
`JARVIS_LLM_BACKEND=claude` was documented as a switch to the Anthropic API,
but `llm/claude_client.py` — imported by both of its call sites
(`core/command_router.py`'s `ask_claude_streaming` and
`llm/tool_caller.py`'s `call_tool_tier_claude`) — does not exist anywhere in
the repo. Selecting this backend would crash the first LLM query with
`ModuleNotFoundError`. (`scripts/rebuild_graduation_doc.py` already documented
this exact gap in its own generated-doc commentary.) Since Ollama is the only
backend that has ever actually run, removed the broken path entirely rather
than leave known-broken code in front of graders:

- `core/command_router.py`: both `LLM_BACKEND == "claude"` branches (now
  always use the Ollama path they already fell back to).
- `llm/tool_caller.py`: `call_tool_tier_claude()`, `_ollama_tools_to_claude_format()`.
- `llm/prompt_builder.py`: `build_claude_messages()`.
- `core/session_memory.py`: `get_messages_for_claude()` (only caller was
  `build_claude_messages`).
- `core/config.py`: `LLM_BACKEND`, `ANTHROPIC_API_KEY`, `CLAUDE_DEFAULT_MODEL`,
  `CLAUDE_QUALITY_MODEL`, `CLAUDE_MAX_TOKENS_COMMAND`, `CLAUDE_MAX_TOKENS_QUESTION`.
- `requirements.txt`: `anthropic>=0.40.0`.
- `.env.example` / `.env`: the corresponding `JARVIS_LLM_BACKEND`,
  `ANTHROPIC_API_KEY`, `JARVIS_CLAUDE_*` keys and the misleading "switch from
  Ollama" comment.

### Reviewed, left in place (out of scope)

- `core/intent_schema.py`'s `OS_CONFIRMATION` registration and
  `core/route_verifier.py`: both modules are explicit, documented
  ahead-of-schedule infrastructure for a separate, not-yet-active migration
  ("Do NOT delete those literals yet" / "the router keeps calling
  assess_intent_confidence directly for now" — their own docstrings). Not
  part of the dead hex-token flow's live call graph; touching them risks an
  unrelated in-progress refactor.
- Two remaining `"OS_CONFIRMATION"` entries in defensive disallow-lists
  (`core/command_router.py`'s job-queue guard, `os_control/batch_ops.py`'s
  `DISALLOWED_BATCH_INTENTS`) are inert (the intent can never be produced)
  but harmless — left as-is rather than editing security blocklists during
  a cleanup pass.
- `format_confirmation_prompt()` in `core/response_templates.py` always takes
  its `token == "PIN_REQUIRED"` early-return in current usage, making its
  `timeout_seconds`/`risk_tier` parameters effectively dead — but simplifying
  a shared, multi-caller utility's body is a larger refactor than this pass's
  scope; its callers were left untouched.

### Dependency review (7.3)
Checked each candidate for actual imports:
- `openwakeword`: used (`audio/wake_word.py`, `core/adaptive_wake.py`) — kept.
- `ddgs` / `duckduckgo_search`: both imported in `tools/web_search.py`
  (primary + fallback) — kept both.
- `wmi`, `watchdog`, `screen_brightness_control`, `pyperclip`: all used — kept.
- `pystray`: zero references anywhere, and not even listed in
  `requirements.txt` — nothing to remove.
- `anthropic`: zero references in application code (see above) — removed.

### Verification
`python -c "import main, audio.stt, audio.tts, core.orchestrator, core.command_router, ui.bridge"`
and `python -m pytest tests/ -q` (69 passed) after every file group;
`vulture --min-confidence 60` re-run clean at the end.
