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


# Radio/device toggles whose training utterances overlap between OS_SETTINGS
# ("open X settings") and OS_SYSTEM_COMMAND ("turn X on/off") in the semantic
# router — a bare "Bluetooth"/"البلوتوث" utterance sits ambiguously between
# them since both intents' training data share the device name as the
# strongest token. Detecting the device lets us ask a concrete on/off/settings
# question instead of the generic, unhelpful "did you mean X or Y?" phrasing.
_RADIO_DEVICE_MARKERS: dict[str, tuple[str, ...]] = {
    "bluetooth": ("bluetooth", "بلوتوث", "البلوتوث"),
    "wifi": ("wifi", "wi-fi", "wi fi", "واي فاي", "الواي فاي", "وايفاي"),
    "airplane_mode": ("airplane", "flight mode", "الطيران", "وضع الطيران"),
}

_RADIO_DEVICE_LABELS: dict[str, dict[str, str]] = {
    "bluetooth": {"en": "Bluetooth", "ar": "البلوتوث"},
    "wifi": {"en": "Wi-Fi", "ar": "الواي فاي"},
    "airplane_mode": {"en": "airplane mode", "ar": "وضع الطيران"},
}

_RADIO_AMBIGUITY_TEMPLATE = {
    "en": "Do you want to turn {device} on or off, or open its settings?",
    "ar": "تقصد تشغّل أو تطفي {device}، ولا تفتح إعداداته؟",
}

# action_key suffix used by os_control.system_ops for each device's on/off
# toggle (e.g. "bluetooth" -> "bluetooth_on"/"bluetooth_off"). airplane_mode's
# action_key is "airplane_on"/"airplane_off" (no "_mode" suffix).
_RADIO_ACTION_KEY_PREFIX = {
    "bluetooth": "bluetooth",
    "wifi": "wifi",
    "airplane_mode": "airplane",
}

_RADIO_OPTION_LABELS = {
    "on": {"en": "turn on", "ar": "تشغيل"},
    "off": {"en": "turn off", "ar": "تطفية"},
    "settings": {"en": "open settings", "ar": "فتح الإعدادات"},
}

_RADIO_OPTION_REPLY_TOKENS = {
    "on": ("on", "turn on", "enable", "شغل", "شغله", "شغّل", "تشغيل"),
    "off": ("off", "turn off", "disable", "اطفي", "اطفيه", "أطفي", "قفل", "تطفية", "ايقاف", "إيقاف"),
    "settings": ("settings", "الاعدادات", "الإعدادات", "اعدادات", "إعدادات"),
}


def detect_radio_device(source_text: str) -> str:
    """Return the radio device key mentioned in source_text, or "" if none."""
    normalized = str(source_text or "").strip().lower()
    if not normalized:
        return ""
    for device_key, markers in _RADIO_DEVICE_MARKERS.items():
        if any(marker in normalized for marker in markers):
            return device_key
    return ""


# Backward-compatible alias for the original internal name.
_detect_radio_device = detect_radio_device


def build_radio_device_options(device_key: str, language: str = "en") -> list[dict]:
    """Return 3 dispatchable options (on/off/settings) for a radio device.

    Each option carries a real action_key so the resolver can dispatch
    directly — no re-parsing of the original ambiguous text needed.
    """
    lang = _lang_key(language)
    prefix = _RADIO_ACTION_KEY_PREFIX.get(device_key)
    if not prefix:
        return []
    return [
        {
            "id": "on",
            "label": _RADIO_OPTION_LABELS["on"][lang],
            "intent": "OS_SYSTEM_COMMAND",
            "action": "",
            "args": {"action_key": f"{prefix}_on"},
            "reply_tokens": list(_RADIO_OPTION_REPLY_TOKENS["on"]),
        },
        {
            "id": "off",
            "label": _RADIO_OPTION_LABELS["off"][lang],
            "intent": "OS_SYSTEM_COMMAND",
            "action": "",
            "args": {"action_key": f"{prefix}_off"},
            "reply_tokens": list(_RADIO_OPTION_REPLY_TOKENS["off"]),
        },
        {
            "id": "settings",
            "label": _RADIO_OPTION_LABELS["settings"][lang],
            "intent": "OS_SETTINGS",
            "action": "",
            "args": {"page": device_key.replace("_", " ")},
            "reply_tokens": list(_RADIO_OPTION_REPLY_TOKENS["settings"]),
        },
    ]


def build_ambiguity_clarification(candidates_top2, language: str = "en", source_text: str = "") -> str:
    """Return a "did you mean X or Y?" prompt from the top-2 semantic candidates.

    candidates_top2: [(intent_name, score), (intent_name, score)] — as produced
    by nlp.semantic_router.classify_semantic_topk / meta["semantic_candidates"].
    """
    lang = _lang_key(language)
    candidates = list(candidates_top2 or [])
    if not candidates:
        return _DEFAULT_SLOT_TEMPLATE.get(lang, _DEFAULT_SLOT_TEMPLATE["en"])

    candidate_names = {str(c[0]).strip().upper() for c in candidates if c}
    if candidate_names == {"OS_SETTINGS", "OS_SYSTEM_COMMAND"}:
        device_key = _detect_radio_device(source_text)
        if device_key:
            device_label = _RADIO_DEVICE_LABELS[device_key][lang]
            return _RADIO_AMBIGUITY_TEMPLATE[lang].format(device=device_label)

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
