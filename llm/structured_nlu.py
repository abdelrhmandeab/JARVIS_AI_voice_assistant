"""Structured LLM NLU — hard-gated, off-by-default fallback for genuinely
ambiguous/complex commands that every earlier tier (parser, code-switch,
semantic, keyword-NLP, Tier-4 tool-calling) failed to resolve.

Reuses the same Ollama chat endpoint as llm.tool_caller (non-streaming,
strict JSON) rather than adding a new HTTP path. The LLM is asked to return
ONE JSON object describing intent+slots — never executed directly, always
passed through core.route_verifier.verify() by the caller.

This module has no side effects and no config gating of its own — the
STRUCTURED_LLM_NLU_* flags live in core.config and are checked by the
router before this is ever called.
"""

from __future__ import annotations

import json
from typing import Optional

import httpx

from core.config import LLM_MODEL, LLM_OLLAMA_BASE_URL
from core.intent_schema import SCHEMA
from core.logger import logger

_OLLAMA_BASE_URL = str(LLM_OLLAMA_BASE_URL or "http://localhost:11434").rstrip("/")
_CHAT_ENDPOINT = f"{_OLLAMA_BASE_URL}/api/chat"


def _intent_catalog() -> str:
    """One line per schema intent: name, domain, required slots — keeps the
    system prompt grounded in the real schema instead of letting the model
    invent intent names."""
    lines = []
    for name, spec in sorted(SCHEMA.items()):
        slots = ", ".join(spec.required_slots) or "none"
        lines.append(f"- {name} (domain={spec.domain}, required_slots={slots})")
    return "\n".join(lines)


def _build_system_prompt() -> str:
    return (
        "You are a strict command-understanding classifier for a voice assistant. "
        "Given the user's utterance, respond with ONLY a single JSON object matching "
        "this exact schema — no prose, no markdown fences, no explanation:\n"
        "{\n"
        '  "intent": "<one of the known intents below, or LLM_QUERY if none fit>",\n'
        '  "action": "<short action string, or empty>",\n'
        '  "slots": {"<slot_name>": "<value>", ...},\n'
        '  "confidence": <float 0.0-1.0>,\n'
        '  "missing_slots": ["<slot_name>", ...],\n'
        '  "requires_confirmation": <true|false>,\n'
        '  "reason": "<one short phrase explaining the classification>"\n'
        "}\n\n"
        "Known intents:\n" + _intent_catalog() + "\n\n"
        "If the utterance is a question, opinion request, or anything that isn't a "
        "clear device/OS command, use intent=LLM_QUERY. Never invent an intent name "
        "that isn't in the list above."
    )


def _coerce_result(raw: dict) -> Optional[dict]:
    if not isinstance(raw, dict):
        return None
    intent = str(raw.get("intent") or "").strip().upper()
    if not intent:
        return None
    slots = raw.get("slots")
    if not isinstance(slots, dict):
        slots = {}
    missing_slots = raw.get("missing_slots")
    if not isinstance(missing_slots, list):
        missing_slots = []
    try:
        confidence = float(raw.get("confidence"))
    except (TypeError, ValueError):
        confidence = 0.0
    confidence = max(0.0, min(1.0, confidence))
    return {
        "intent": intent,
        "action": str(raw.get("action") or "").strip(),
        "slots": {str(k): v for k, v in slots.items()},
        "confidence": confidence,
        "missing_slots": [str(s) for s in missing_slots],
        "requires_confirmation": bool(raw.get("requires_confirmation")),
        "reason": str(raw.get("reason") or "").strip(),
    }


def understand_structured(
    normalized_text: str,
    language: str = "en",
    timeout: float = 4.0,
    model_name: Optional[str] = None,
) -> Optional[dict]:
    """Ask the LLM for a strict-JSON intent+slots classification.

    Returns a dict with keys _RESPONSE_KEYS on success, or None on any
    failure (network error, non-200, unparseable JSON, missing intent).
    Never raises — callers should treat None the same as "couldn't classify".
    """
    text = str(normalized_text or "").strip()
    if not text:
        return None

    payload = {
        "model": str(model_name or LLM_MODEL or "qwen3:4b"),
        "messages": [
            {"role": "system", "content": _build_system_prompt()},
            {"role": "user", "content": text},
        ],
        "stream": False,
        "format": "json",
        "options": {"temperature": 0.0},
    }

    try:
        response = httpx.post(_CHAT_ENDPOINT, json=payload, timeout=timeout)
    except Exception as exc:
        logger.debug("structured_nlu: request failed: %s", exc)
        return None

    if response.status_code != 200:
        logger.debug("structured_nlu: non-200 response %s: %s", response.status_code, response.text[:200])
        return None

    try:
        payload_json = response.json()
    except Exception as exc:
        logger.debug("structured_nlu: response was not JSON: %s", exc)
        return None

    message = payload_json.get("message") or {}
    content = str(message.get("content") or "").strip()
    if not content:
        return None

    try:
        raw = json.loads(content)
    except Exception as exc:
        logger.debug("structured_nlu: model content was not valid JSON: %s", exc)
        return None

    result = _coerce_result(raw)
    if result is None:
        logger.debug("structured_nlu: coerced result missing required intent field")
        return None

    return result
