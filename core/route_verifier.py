"""Route verifier — one gate that turns a routed candidate into a decision.

Consolidates checks that were previously spread across intent_confidence.py's
IntentAssessment, the router's inline permission check, and the schema's risk
levels into a single RouteDecision: execute / clarify / confirm / llm.

This does not replace the existing dispatch/confirmation machinery (PIN gate,
policy_engine, confirmation_manager) — those remain the actual enforcement
points. verify() is the single place that *explains* what should happen and
why, for logging/telemetry and for Phases 6/7 to consume. Back-compat: the
router keeps calling assess_intent_confidence directly for now; this module
reads its output rather than re-deriving entity/question scoring.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from core import intent_schema
from core.intent_confidence import _QUESTION_OPENER_RE, _QUESTION_PENALTY_INTENTS
from core.logger import get_logger

try:
    from os_control.risk_policy import (
        risk_tier_for_system,
        risk_tier_for_file_operation,
        risk_tier_for_app_operation,
    )
except Exception:  # pragma: no cover - risk_policy always present in this repo
    risk_tier_for_system = None
    risk_tier_for_file_operation = None
    risk_tier_for_app_operation = None

logger = get_logger("router")

_RISK_ORDER = {"none": 0, "low": 1, "medium": 2, "high": 3}


@dataclass
class RouteDecision:
    action: str  # "execute" | "clarify" | "confirm" | "llm"
    intent: str
    slots: dict = field(default_factory=dict)
    reason: str = ""
    confidence: float = 0.0


def _effective_risk(intent: str, action: str, slots: dict) -> str:
    """Return the finer-grained risk tier when a per-action override exists,
    else fall back to the coarse per-intent schema risk.

    `action` is ParsedCommand.action (e.g. "delete_item", "move_item") — not
    an args/slots entry. OS_SYSTEM_COMMAND is the exception: its per-action
    key lives in slots["action_key"] instead of the .action field.
    """
    schema_risk = intent_schema.risk(intent)
    intent = str(intent or "").strip().upper()
    slots = slots or {}
    action = str(action or "").strip().lower()

    if intent == "OS_SYSTEM_COMMAND" and risk_tier_for_system is not None:
        action_key = str(slots.get("action_key") or "").strip().lower()
        if action_key:
            from os_control.system_ops import SYSTEM_COMMANDS

            cfg = SYSTEM_COMMANDS.get(action_key) or {}
            destructive = bool(cfg.get("destructive"))
            requires_confirmation = bool(cfg.get("requires_confirmation", destructive))
            override = risk_tier_for_system(
                action_key, destructive=destructive, requires_confirmation=requires_confirmation
            )
            # Use whichever is higher — never let a per-action override silently
            # downgrade what the schema already considers risky.
            if _RISK_ORDER.get(override, 0) > _RISK_ORDER.get(schema_risk, 0):
                return override
            return schema_risk

    if intent == "OS_FILE_NAVIGATION" and risk_tier_for_file_operation is not None and action:
        override = risk_tier_for_file_operation(action)
        if _RISK_ORDER.get(override, 0) > _RISK_ORDER.get(schema_risk, 0):
            return override

    if intent == "OS_APP_CLOSE" and risk_tier_for_app_operation is not None:
        override = risk_tier_for_app_operation("close_app")
        if _RISK_ORDER.get(override, 0) > _RISK_ORDER.get(schema_risk, 0):
            return override

    return schema_risk


def verify(candidate_parsed, normalized, confidence, entity_scores=None,
           language="en", fast_command_min_confidence=0.88,
           should_clarify=False) -> RouteDecision:
    """Turn a routed candidate into an execute/clarify/confirm/llm decision.

    candidate_parsed: ParsedCommand-like object with .intent, .action, .args
    normalized: normalized utterance text (for the question-opener check)
    confidence: numeric confidence for this candidate (0..1)
    entity_scores: optional dict of per-slot entity confidence scores
    should_clarify: the caller's own assess_intent_confidence verdict — this
        verifier defers to it rather than re-deriving a confidence floor.
    """
    intent = str(getattr(candidate_parsed, "intent", "") or "").strip().upper()
    action = str(getattr(candidate_parsed, "action", "") or "")
    slots = dict(getattr(candidate_parsed, "args", None) or {})
    entity_scores = entity_scores or {}

    # 1. Schema existence — no schema entry means we don't know its shape/risk.
    spec = intent_schema.get_spec(intent)
    if spec is None:
        return RouteDecision(
            action="llm", intent=intent, slots=slots,
            reason="no_schema_entry", confidence=confidence,
        )

    # 2. Required slots present.
    missing = [slot for slot in spec.required_slots if not str(slots.get(slot) or "").strip()]
    if missing:
        return RouteDecision(
            action="clarify", intent=intent, slots=slots,
            reason=f"missing_slot:{missing[0]}", confidence=confidence,
        )

    # 3. Question-masquerading-as-command penalty.
    if intent in _QUESTION_PENALTY_INTENTS and _QUESTION_OPENER_RE.match(
        " ".join(str(normalized or "").split())
    ):
        return RouteDecision(
            action="llm", intent=intent, slots=slots,
            reason="question_opener_detected", confidence=confidence,
        )

    # 4. Entity confidence.
    if entity_scores:
        min_score = min(float(v) for v in entity_scores.values())
        if min_score < 0.45:
            return RouteDecision(
                action="clarify", intent=intent, slots=slots,
                reason="low_entity_confidence", confidence=confidence,
            )

    # 5. Risk gate.
    risk = _effective_risk(intent, action, slots)
    if risk == "high":
        return RouteDecision(
            action="confirm", intent=intent, slots=slots,
            reason="high_risk", confidence=confidence,
        )
    if risk == "medium":
        return RouteDecision(
            action="confirm", intent=intent, slots=slots,
            reason="medium_risk", confidence=confidence,
        )

    # Low/none risk: defer to the caller's own assess_intent_confidence
    # verdict. fast_command_min_confidence (0.88) is accepted here for the
    # eval harness (Phase 8/9) to tune against, but isn't applied as a gate
    # yet — assess_intent_confidence's current per-intent confidence
    # ceilings for common commands (e.g. OS_APP_OPEN/CLOSE=0.80,
    # OS_FILE_NAVIGATION list_directory=0.78) sit below 0.88 by design, so a
    # flat floor here would force clarification on clean, common commands
    # that assess_intent_confidence already judged confident enough.
    if should_clarify:
        return RouteDecision(
            action="clarify", intent=intent, slots=slots,
            reason="low_confidence", confidence=confidence,
        )

    # 6. Policy permission.
    try:
        from core.command_router import _required_permission
        from os_control.policy import policy_engine

        permission_key = _required_permission(candidate_parsed)
        if permission_key and not policy_engine.is_command_allowed(permission_key):
            return RouteDecision(
                action="llm", intent=intent, slots=slots,
                reason=f"policy_blocked:{permission_key}", confidence=confidence,
            )
    except Exception as exc:
        logger.debug("route_verifier: policy check skipped (%s)", exc)

    return RouteDecision(
        action="execute", intent=intent, slots=slots,
        reason="ok", confidence=confidence,
    )
