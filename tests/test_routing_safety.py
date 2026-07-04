import unittest

from tools.evaluate_nlu import load_cases, run_eval


class RoutingSafetyTests(unittest.TestCase):
    """Runs the full labeled eval set and asserts the hard safety invariants:
    zero unsafe auto-executions and zero question-to-command false-fires.
    Uses core.eval_routing.decision_only (no dispatch, no session_memory, no
    network LLM calls) — see tools/evaluate_nlu.py for the scoring rules.
    """

    @classmethod
    def setUpClass(cls):
        cls.cases = load_cases()
        cls.summary = run_eval(cls.cases)

    def test_zero_unsafe_executions(self):
        failures = [r for r in self.summary["results"] if r["unsafe_execution"]]
        self.assertEqual(
            self.summary["unsafe_execution_count"], 0,
            msg=f"Unsafe executions found: {failures}",
        )

    def test_zero_question_to_command_false_fires(self):
        failures = [r for r in self.summary["results"] if r["question_false_fire"]]
        self.assertEqual(
            self.summary["question_to_command_false_fires"], 0,
            msg=f"Question->command false-fires found: {failures}",
        )


if __name__ == "__main__":
    unittest.main()
