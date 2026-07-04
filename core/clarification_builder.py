"""Clarification builder — template-based, slot-specific clarification prompts.

Centralizes the bilingual "which app?" / "for how long?" style questions that
were previously a small ad-hoc table in command_router.py (_SLOT_QUESTIONS),
and adds the semantic near-tie ("did you mean X or Y?") prompt consumed from
Phase 3's margin-scoring signal (meta["semantic_ambiguous"]/["semantic_candidates"]).

No LLM involved — these are plain string templates. Resolution still goes
through the existing pending-clarification store (session_memory) and
resolve_clarification_reply / the missing-slot slot-fill handler in
command_router.py; this module only builds the prompt text.
"""

from __future__ import annotations

# Per-slot bilingual templates, keyed by the generic slot name (not per-intent
# — the same "which app?" question works for OS_APP_OPEN and OS_APP_CLOSE).
_SLOT_TEMPLATES: dict[str, dict[str, str]] = {
    "app_name": {
        "en": "Which app?",
        "ar": "أي برنامج؟",
    },
    "seconds": {
        "en": "For how long?",
        "ar": "لمدة قد إيه؟",
    },
    "time_str": {
        "en": "For how long?",
        "ar": "لمدة قد إيه؟",
    },
    "path": {
        "en": "Which folder?",
        "ar": "أنهي فولدر؟",
    },
    "search_query": {
        "en": "Search for what?",
        "ar": "أدور على إيه؟",
    },
    "filename": {
        "en": "Search for what?",
        "ar": "أدور على إيه؟",
    },
    "query": {
        "en": "Search for what?",
        "ar": "أدور على إيه؟",
    },
    "to": {
        "en": "Who is this for?",
        "ar": "لمين؟",
    },
    "subject": {
        "en": "What's the subject?",
        "ar": "الموضوع إيه؟",
    },
    "body": {
        "en": "What should it say?",
        "ar": "عايز أكتب فيها إيه؟",
    },
}

# Per-intent overrides for a slot, when the generic template above doesn't
# fit (e.g. OS_APP_OPEN vs OS_APP_CLOSE want different verbs). Falls back to
# _SLOT_TEMPLATES when no override exists for (intent, slot).
_INTENT_SLOT_OVERRIDES: dict[tuple[str, str], dict[str, str]] = {
    ("OS_APP_OPEN", "app_name"): {
        "en": "Which app would you like to open?",
        "ar": "أي تطبيق تريد تفتحه؟",
    },
    ("OS_APP_CLOSE", "app_name"): {
        "en": "Which app would you like to close?",
        "ar": "أي تطبيق تريد تقفله؟",
    },
    ("OS_TIMER", "seconds"): {
        "en": "How long should the timer run?",
        "ar": "التايمر على كام؟",
    },
    ("OS_FILE_SEARCH", "filename"): {
        "en": "What file are you looking for?",
        "ar": "بتدور على أي ملف؟",
    },
}

_DEFAULT_SLOT_TEMPLATE = {
    "en": "Could you clarify?",
    "ar": "ممكن توضح أكتر؟",
}

_AMBIGUITY_TEMPLATE = {
    "en": "Did you mean {first} or {second}?",
    "ar": "تقصد {first} ولا {second}؟",
}

# Human-friendly labels for intent names used in the ambiguity prompt — the
# raw SCHEMA/route names (OS_APP_OPEN) aren't something to read aloud.
_INTENT_LABELS: dict[str, dict[str, str]] = {
    "OS_APP_OPEN": {"en": "opening an app", "ar": "فتح برنامج"},
    "OS_APP_CLOSE": {"en": "closing an app", "ar": "قفل برنامج"},
    "OS_FILE_SEARCH": {"en": "searching for a file", "ar": "الدور على ملف"},
    "OS_FILE_NAVIGATION": {"en": "browsing files", "ar": "تصفح الملفات"},
    "OS_SYSTEM_COMMAND": {"en": "a system action", "ar": "أمر نظام"},
    "OS_SETTINGS": {"en": "opening settings", "ar": "فتح الإعدادات"},
    "OS_EMAIL": {"en": "email", "ar": "الإيميل"},
    "OS_CALENDAR": {"en": "the calendar", "ar": "الكالندر"},
    "OS_TIMER": {"en": "a timer", "ar": "تايمر"},
    "OS_NOTE": {"en": "a note", "ar": "نوتة"},
    "OS_SCREEN_DESCRIBE": {"en": "describing the screen", "ar": "وصف الشاشة"},
    "IDENTITY": {"en": "asking who I am", "ar": "تعرف عليا"},
    "LLM_QUERY": {"en": "a question", "ar": "سؤال"},
}


def _lang_key(language: str) -> str:
    return "ar" if str(language or "en").strip().lower().startswith("ar") else "en"


def build_slot_clarification(intent: str, missing_slot: str, language: str = "en") -> str:
    """Return a bilingual, slot-specific clarification question.

    Looks up an (intent, slot) override first, then a generic per-slot
    template, then falls back to a generic "could you clarify?" prompt.
    """
    intent_key = str(intent or "").strip().upper()
    slot_key = str(missing_slot or "").strip().lower()
    lang = _lang_key(language)

    override = _INTENT_SLOT_OVERRIDES.get((intent_key, slot_key))
    if override:
        return override.get(lang, override.get("en", ""))

    generic = _SLOT_TEMPLATES.get(slot_key)
    if generic:
        return generic.get(lang, generic.get("en", ""))

    return _DEFAULT_SLOT_TEMPLATE.get(lang, _DEFAULT_SLOT_TEMPLATE["en"])


def intent_label(intent: str, language: str = "en") -> str:
    """Human-friendly label for an intent name (e.g. OS_APP_OPEN -> "opening an app")."""
    key = str(intent or "").strip().upper()
    lang = _lang_key(language)
    labels = _INTENT_LABELS.get(key)
    if labels:
        return labels.get(lang, labels.get("en", key))
    return key


_intent_label = intent_label


def build_ambiguity_clarification(candidates_top2, language: str = "en") -> str:
    """Return a "did you mean X or Y?" prompt from the top-2 semantic candidates.

    candidates_top2: [(intent_name, score), (intent_name, score)] — as produced
    by nlp.semantic_router.classify_semantic_topk / meta["semantic_candidates"].
    """
    lang = _lang_key(language)
    candidates = list(candidates_top2 or [])
    if not candidates:
        return _DEFAULT_SLOT_TEMPLATE.get(lang, _DEFAULT_SLOT_TEMPLATE["en"])

    first = _intent_label(candidates[0][0] if candidates[0] else "", lang)
    if len(candidates) < 2:
        # Only one candidate — ask for confirmation rather than a false choice.
        confirm_template = {
            "en": "Did you mean {first}?",
            "ar": "تقصد {first}؟",
        }
        return confirm_template[lang].format(first=first)

    second = _intent_label(candidates[1][0] if candidates[1] else "", lang)
    return _AMBIGUITY_TEMPLATE[lang].format(first=first, second=second)
