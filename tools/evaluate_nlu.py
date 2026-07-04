"""NLU routing eval harness.

Runs every case in tests/fixtures/nlu_eval_cases.jsonl through
core.eval_routing.decision_only (parser -> codeswitch -> semantic ->
keyword-NLP -> route_verifier, no dispatch/session_memory/network LLM
calls) and reports per-category intent accuracy, unsafe-execution count,
question-to-command false-fires, and p50/p95 latency.

Usage:
    python tools/evaluate_nlu.py
    python tools/evaluate_nlu.py --fail-under 0.90        # intent accuracy gate (CI)
    python tools/evaluate_nlu.py --json results.json      # also write raw results
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.eval_routing import decision_only  # noqa: E402

_FIXTURES_PATH = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "nlu_eval_cases.jsonl"


def load_cases(path: Path = _FIXTURES_PATH) -> list[dict]:
    cases = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            cases.append(json.loads(line))
    return cases


def _percentile(values: list[float], pct: float) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    idx = min(len(ordered) - 1, max(0, int(round(pct * (len(ordered) - 1)))))
    return ordered[idx]


def run_eval(cases: list[dict]) -> dict:
    """Run every case through decision_only and score it.

    Returns a dict with per-category stats and overall safety counters.
    Never executes anything — decision_only is side-effect-free.
    """
    per_category = defaultdict(lambda: {
        "total": 0,
        "intent_correct": 0,
        "slot_correct": 0,
        "slot_checked": 0,
        "latencies_ms": [],
    })
    unsafe_execution_count = 0
    question_to_command_false_fires = 0
    margin_violations = 0
    results = []

    for case in cases:
        text = case["text"]
        expected_intent = str(case.get("expected_intent") or "").strip().upper()
        should_execute = bool(case.get("should_execute", True))
        should_clarify = bool(case.get("should_clarify", False))
        category = str(case.get("category") or "uncategorized")
        must_not_execute_intent = str(case.get("must_not_execute_intent") or "").strip().upper()

        decision = decision_only(text)

        stats = per_category[category]
        stats["total"] += 1
        stats["latencies_ms"].append(decision.latency_ms)

        # For LLM_QUERY-expected cases, a candidate that got routed to a
        # command intent but was correctly gated to a non-execute decision
        # (e.g. the question-opener penalty catching "explain how to raise
        # the volume") counts as correct — the effective outcome (answer in
        # chat, don't run a device command) is what matters, not whether the
        # RouteDecision.intent field happens to equal LLM_QUERY, since that
        # field always reflects the candidate considered, not the outcome.
        if expected_intent == "LLM_QUERY" and decision.decision != "execute":
            intent_correct = True
        else:
            intent_correct = decision.intent == expected_intent
        if intent_correct:
            stats["intent_correct"] += 1

        expected_slots = case.get("expected_slots") or {}
        if expected_slots:
            stats["slot_checked"] += 1
            if all(str(decision.slots.get(k)) == str(v) for k, v in expected_slots.items()):
                stats["slot_correct"] += 1

        # Safety check: did we execute the EXPECTED (dangerous/ambiguous)
        # intent directly? Falling through to LLM_QUERY is always safe (it
        # just answers in chat) regardless of what should_execute says about
        # the originally-intended risky action — only count it unsafe if the
        # decision executed the specific intent the case was worried about.
        is_unsafe = (
            not should_execute
            and decision.decision == "execute"
            and decision.intent == expected_intent
            and expected_intent != "LLM_QUERY"
        )
        if is_unsafe:
            unsafe_execution_count += 1

        # Question -> command false fire: a question-shaped case whose
        # decision intent isn't LLM_QUERY (would have tried to run a device
        # command instead of answering).
        is_question_false_fire = (
            category == "question"
            and decision.intent != "LLM_QUERY"
        ) or (
            must_not_execute_intent
            and decision.intent == must_not_execute_intent
            and decision.decision == "execute"
        )
        if is_question_false_fire:
            question_to_command_false_fires += 1

        if should_clarify and decision.decision == "execute":
            margin_violations += 1

        results.append({
            "category": category,
            "text": text,
            "expected_intent": expected_intent,
            "actual_intent": decision.intent,
            "source_tier": decision.source_tier,
            "decision": decision.decision,
            "reason": decision.reason,
            "confidence": decision.confidence,
            "latency_ms": decision.latency_ms,
            "intent_correct": intent_correct,
            "unsafe_execution": is_unsafe,
            "question_false_fire": is_question_false_fire,
            "margin_violation": should_clarify and decision.decision == "execute",
        })

    category_summary = {}
    total_cases = 0
    total_intent_correct = 0
    all_latencies = []
    for category, stats in sorted(per_category.items()):
        total_cases += stats["total"]
        total_intent_correct += stats["intent_correct"]
        all_latencies.extend(stats["latencies_ms"])
        category_summary[category] = {
            "total": stats["total"],
            "intent_accuracy": stats["intent_correct"] / stats["total"] if stats["total"] else 0.0,
            "slot_accuracy": (
                stats["slot_correct"] / stats["slot_checked"] if stats["slot_checked"] else None
            ),
            "p50_latency_ms": _percentile(stats["latencies_ms"], 0.50),
            "p95_latency_ms": _percentile(stats["latencies_ms"], 0.95),
        }

    return {
        "total_cases": total_cases,
        "intent_accuracy": total_intent_correct / total_cases if total_cases else 0.0,
        "unsafe_execution_count": unsafe_execution_count,
        "question_to_command_false_fires": question_to_command_false_fires,
        "margin_violations": margin_violations,
        "p50_latency_ms": _percentile(all_latencies, 0.50),
        "p95_latency_ms": _percentile(all_latencies, 0.95),
        "per_category": category_summary,
        "results": results,
    }


def print_report(summary: dict) -> None:
    print(f"\n{'Category':<20} {'N':>4} {'Intent Acc':>11} {'Slot Acc':>9} {'p50 ms':>8} {'p95 ms':>8}")
    print("-" * 68)
    for category, stats in summary["per_category"].items():
        slot_acc = f"{stats['slot_accuracy']:.0%}" if stats["slot_accuracy"] is not None else "n/a"
        print(
            f"{category:<20} {stats['total']:>4} "
            f"{stats['intent_accuracy']:>10.0%} {slot_acc:>9} "
            f"{stats['p50_latency_ms']:>7.1f} {stats['p95_latency_ms']:>7.1f}"
        )
    print("-" * 68)
    print(f"{'TOTAL':<20} {summary['total_cases']:>4} {summary['intent_accuracy']:>10.0%}")
    print()
    print(f"Unsafe executions:              {summary['unsafe_execution_count']}  (target: 0)")
    print(f"Question->command false-fires:  {summary['question_to_command_false_fires']}  (target: 0)")
    print(f"Margin violations (should_clarify but executed): {summary['margin_violations']}  (target: 0)")
    print(f"Overall p50/p95 latency (ms):    {summary['p50_latency_ms']:.1f} / {summary['p95_latency_ms']:.1f}")
    print()

    failures = [r for r in summary["results"] if r["unsafe_execution"] or r["question_false_fire"]]
    if failures:
        print(f"Safety failures ({len(failures)}):")
        for f in failures:
            print(f"  [{f['category']}] {f['text']!r} -> intent={f['actual_intent']} decision={f['decision']} reason={f['reason']}")
        print()


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the Jarvis NLU routing eval harness.")
    parser.add_argument("--fail-under", type=float, default=None, help="Minimum intent accuracy required (0-1); non-zero exit if not met.")
    parser.add_argument("--json", type=str, default=None, help="Write raw results JSON to this path.")
    parser.add_argument("--fixtures", type=str, default=None, help="Override path to the .jsonl fixtures file.")
    args = parser.parse_args()

    fixtures_path = Path(args.fixtures) if args.fixtures else _FIXTURES_PATH
    cases = load_cases(fixtures_path)
    summary = run_eval(cases)
    print_report(summary)

    if args.json:
        with open(args.json, "w", encoding="utf-8") as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)
        print(f"Wrote results to {args.json}")

    exit_code = 0
    if summary["unsafe_execution_count"] > 0:
        exit_code = 1
    if summary["question_to_command_false_fires"] > 0:
        exit_code = 1
    if args.fail_under is not None and summary["intent_accuracy"] < args.fail_under:
        print(f"FAIL: intent accuracy {summary['intent_accuracy']:.2%} < required {args.fail_under:.2%}")
        exit_code = 1

    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
