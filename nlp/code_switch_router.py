"""Code-switch router — fast dictionary/token shortcut before the semantic tier.

Resolves mixed-language "verb + entity" utterances like "افتح Chrome",
"زود volume", "open الملفات" by dictionary + token match in well under the
embedding model's latency, so the semantic tier only runs when this shortcut
can't confidently resolve the command.

Reuses the verb sets already declared in core.intent_confidence and the
system-action / path-resolution helpers already used by the regex parser —
this module holds no new vocabulary of its own beyond the small close-verb
set that the rest of the codebase doesn't already centralise.
"""

from __future__ import annotations

import re
from typing import Optional

from core.command_parser import ParsedCommand
from core.intent_confidence import _OPEN_VERBS, _DELETE_VERBS, _MOVE_VERBS, _RENAME_VERBS
from core.config import FILE_DEFAULT_SEARCH_ROOTS
from os_control.file_ops import resolve_name_in_location
from os_control.path_resolver import resolve_location, KNOWN_FOLDERS
from os_control.system_ops import normalize_system_action

# Close verbs aren't centralised anywhere else in the repo (only open/delete/
# move/rename/system live in intent_confidence) — declared here, not
# duplicated, and reusable by any future caller that needs them.
_CLOSE_VERBS = {
    "close",
    "quit",
    "exit",
    "kill",
    "stop",
    "اقفل",
    "سكر",
    "اطفي",
    "قفل",
    "سكّر",
}

_WORD_SPLIT_RE = re.compile(r"\s+")

# Separator that splits "<source> <sep> <new_name/destination>" for rename
# and move — mirrors the separators command_parser's own rename/move regexes
# use, so code-switched phrasing ("rename report الى final_report") resolves
# the same way a same-language phrasing would.
_RENAME_MOVE_SEPARATOR_RE = re.compile(
    r"\s+(?:to|as|into|اسمه|باسم|الى|الي|على|علي|ل(?:ـ)?|للـ?)\s+",
    re.IGNORECASE,
)


def _split_verb_entity(normalized_text: str) -> Optional[tuple[str, str]]:
    """Split "<verb> <entity...>" into (verb, entity). Returns None if too short."""
    tokens = _WORD_SPLIT_RE.split(normalized_text.strip())
    tokens = [t for t in tokens if t]
    if len(tokens) < 2:
        return None
    return tokens[0].lower(), " ".join(tokens[1:]).strip()


def _split_verb_entity_target(normalized_text: str) -> Optional[tuple[str, str, str]]:
    """Split "<verb> <source> <sep> <target>" into (verb, source, target).

    Used for rename/move, which need two parts (source + new name / dest)
    unlike open/close/delete's single-entity shortcut. Returns None if the
    separator isn't found or either side is empty.
    """
    split = _split_verb_entity(normalized_text)
    if split is None:
        return None
    verb, rest = split
    parts = _RENAME_MOVE_SEPARATOR_RE.split(rest, maxsplit=1)
    if len(parts) != 2:
        return None
    source, target = parts[0].strip(), parts[1].strip()
    if not source or not target:
        return None
    return verb, source, target


def _resolve_in_default_roots(name: str) -> Optional[str]:
    """Search FILE_DEFAULT_SEARCH_ROOTS for a single unambiguous match for
    *name* (extension-less or not). Returns the resolved path, or None if
    not found / ambiguous — callers fall through rather than guess."""
    for root_name in FILE_DEFAULT_SEARCH_ROOTS:
        folder = KNOWN_FOLDERS.get(root_name)
        if not folder:
            continue
        status, value = resolve_name_in_location(name, str(folder))
        if status == "single":
            return value
    return None


def try_codeswitch(normalized_text: str, language: str = "") -> Optional[ParsedCommand]:
    """Attempt to resolve a mixed-language verb+entity command without the
    semantic embedding model.

    Returns a ParsedCommand on a confident resolution, else None (caller
    falls through to the semantic tier).
    """
    text = str(normalized_text or "").strip()
    if not text:
        return None

    normalized = " ".join(text.lower().split())

    # System actions (volume/brightness/wifi/bluetooth/media/...) already have
    # a robust bilingual + mixed-script normalizer — try it on the full phrase
    # first since it doesn't require a clean verb/entity split.
    action_key = normalize_system_action(text)
    if action_key:
        return ParsedCommand(
            intent="OS_SYSTEM_COMMAND",
            raw=text,
            normalized=normalized,
            args={"action_key": action_key},
        )

    split = _split_verb_entity(normalized)
    if split is None:
        return None
    verb, entity = split
    if not entity:
        return None

    if verb in _OPEN_VERBS:
        # Folder/drive entity → file navigation, not app open.
        location = resolve_location(entity)
        if location is not None:
            return ParsedCommand(
                intent="OS_FILE_NAVIGATION",
                raw=text,
                normalized=normalized,
                action="cd",
                args={"path": entity},
            )
        return ParsedCommand(
            intent="OS_APP_OPEN",
            raw=text,
            normalized=normalized,
            args={"app_name": entity},
        )

    if verb in _CLOSE_VERBS:
        return ParsedCommand(
            intent="OS_APP_CLOSE",
            raw=text,
            normalized=normalized,
            args={"app_name": entity},
        )

    if verb in _DELETE_VERBS:
        # Delete only needs one target — resolve it (extension-less or not)
        # across the default search roots rather than requiring an exact path.
        resolved = _resolve_in_default_roots(entity)
        if resolved is not None:
            return ParsedCommand(
                intent="OS_FILE_NAVIGATION",
                raw=text,
                normalized=normalized,
                action="delete_item",
                args={"path": resolved},
            )
        return None

    if verb in _MOVE_VERBS or verb in _RENAME_VERBS:
        # Both need a "<source> <sep> <target>" split — re-split from the
        # normalized text since the simple verb+entity split doesn't capture
        # the separator (to/as/الى/ل...).
        target_split = _split_verb_entity_target(normalized)
        if target_split is None:
            return None
        _verb, source, target = target_split
        resolved = _resolve_in_default_roots(source)
        if resolved is None:
            return None
        if verb in _RENAME_VERBS:
            return ParsedCommand(
                intent="OS_FILE_NAVIGATION",
                raw=text,
                normalized=normalized,
                action="rename_item",
                args={"source": resolved, "new_name": target},
            )
        return ParsedCommand(
            intent="OS_FILE_NAVIGATION",
            raw=text,
            normalized=normalized,
            action="move_item",
            args={"source": resolved, "destination": target},
        )

    return None
