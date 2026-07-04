"""Decision-only routing for the eval harness (tools/evaluate_nlu.py).

route_command() is a large, deeply side-effecting function (dispatches real
OS actions, mutates session_memory, makes network/LLM calls). Rather than
thread a dry_run flag through every branch of it — high-risk for a function
that size — this module composes the same already-isolated tier functions
route_command uses, in the same order, with zero side effects:

    parse_command -> codeswitch -> semantic -> keyword-NLP
    -> NLU entity enrichment -> assess_intent_confidence -> route_verifier.verify

Deliberately excluded: Tier-4 tool-calling and the Phase 7 structured-LLM
NLU fallback, since both make real network LLM calls — running those on
every eval case would be slow, non-deterministic, and a departure from
"decision-only". Cases that only those tiers could resolve will show up as
LLM_QUERY here, which is a known, documented limitation of this harness.
"""

from __future__ import annotations

import time
from dataclasses import dataclass

from core import route_verifier
from core.command_parser import parse_command
from core.intent_confidence import assess_intent_confidence
from core.language_gate import detect_supported_language
from core.config import FAST_COMMAND_MIN_CONFIDENCE

try:
    from nlp.nlu import understand as _nlu_understand
except Exception:  # pragma: no cover - matches command_router's own guard
    _nlu_understand = None


@dataclass
class EvalDecision:
    text: str
    language: str
    source_tier: str  # parser | codeswitch | semantic | keyword_nlp | fallback
    intent: str
    action: str
    slots: dict
    missing_slots: list
    decision: str  # execute | clarify | confirm | llm
    reason: str
    confidence: float
    latency_ms: float
    semantic_top_3: list


def _try_codeswitch_tier(source_text, parser_candidate, language):
    """Delegates to core.command_router's real Tier-1.5 function rather than
    reimplementing its verb/entity/path-resolution rules here."""
    try:
        from core.command_router import _try_codeswitch_routing
        mapped, _meta = _try_codeswitch_routing(source_text, parser_candidate, language)
        return mapped
    except Exception:
        return None


def _try_semantic_tier(source_text, parser_candidate):
    """Delegates to core.command_router's real Tier-2 function (top-k +
    margin scoring) rather than reimplementing the threshold/margin logic
    here, which would drift out of sync with Phase 3's tuning."""
    try:
        from core.command_router import _try_semantic_routing
        mapped, meta = _try_semantic_routing(source_text, parser_candidate)
        return mapped, list(meta.get("semantic_top_3") or [])
    except Exception:
        return None, []


def _try_keyword_nlp_tier(source_text, parser_candidate):
    """Delegates to core.command_router's real keyword-NLP tier (classifier +
    _map_keyword_nlp_intent_to_command) instead of reimplementing its
    informational-query guards / confidence floors / intent mapping here —
    those rules are non-trivial and drift-prone to duplicate."""
    try:
        from core.command_router import _try_keyword_nlp_routing
        mapped, _meta = _try_keyword_nlp_routing(source_text, parser_candidate)
        return mapped
    except Exception:
        return None


def decision_only(text: str, language_hint: str = "") -> EvalDecision:
    """Run the routing cascade (no dispatch, no session_memory, no LLM
    network calls) and return the RouteDecision the verifier would make."""
    start = time.perf_counter()
    source_text = str(text or "")

    language_result = detect_supported_language(source_text, previous_language=language_hint or "en")
    language = language_result.language if language_result.reason != "blocked" else (language_hint or "en")

    parser_candidate = parse_command(source_text)
    parsed = None
    source_tier = "fallback"
    semantic_top_3: list = []

    if str(parser_candidate.intent or "").strip().upper() != "LLM_QUERY":
        parsed = parser_candidate
        source_tier = "parser"

    if parsed is None:
        codeswitch_parsed = _try_codeswitch_tier(source_text, parser_candidate, language)
        if codeswitch_parsed is not None:
            parsed = codeswitch_parsed
            source_tier = "codeswitch"

    if parsed is None:
        semantic_parsed, semantic_top_3 = _try_semantic_tier(source_text, parser_candidate)
        if semantic_parsed is not None:
            parsed = semantic_parsed
            source_tier = "semantic"

    if parsed is None:
        keyword_parsed = _try_keyword_nlp_tier(source_text, parser_candidate)
        if keyword_parsed is not None:
            parsed = keyword_parsed
            source_tier = "keyword_nlp"

    if parsed is None:
        parsed = parser_candidate
        source_tier = "fallback"

    missing_slots: list = []
    intent_upper = str(parsed.intent or "").strip().upper()
    if _nlu_understand is not None and intent_upper not in ("LLM_QUERY", ""):
        try:
            nlu_result = _nlu_understand(
                source_text, language, intent=intent_upper, existing_args=dict(parsed.args or {}),
            )
            for slot_key, slot_val in nlu_result.entities.items():
                if slot_key not in parsed.args or not parsed.args[slot_key]:
                    parsed.args[slot_key] = slot_val
            missing_slots = list(nlu_result.missing_slots)
        except Exception:
            pass

    assessment = assess_intent_confidence(source_text, parsed, language=language)

    decision = route_verifier.verify(
        parsed,
        source_text,
        assessment.confidence,
        entity_scores=assessment.entity_scores,
        language=language,
        fast_command_min_confidence=FAST_COMMAND_MIN_CONFIDENCE,
        should_clarify=assessment.should_clarify,
    )

    latency_ms = (time.perf_counter() - start) * 1000.0
    return EvalDecision(
        text=source_text,
        language=language,
        source_tier=source_tier,
        intent=decision.intent,
        action=str(getattr(parsed, "action", "") or ""),
        slots=dict(decision.slots or {}),
        missing_slots=missing_slots,
        decision=decision.action,
        reason=decision.reason,
        confidence=float(decision.confidence),
        latency_ms=latency_ms,
        semantic_top_3=semantic_top_3,
    )
