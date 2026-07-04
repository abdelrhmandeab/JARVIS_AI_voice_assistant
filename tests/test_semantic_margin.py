import unittest

from core.eval_routing import decision_only


class SemanticMarginTests(unittest.TestCase):
    """Asserts Phase 3's top-k + margin scoring: a near-tie semantic match
    must not silently auto-execute — it should defer (clarify) rather than
    guess between two similarly-scored intents.
    """

    def test_bare_open_defers_instead_of_guessing(self):
        # "open" alone is a real near-tie in this repo's semantic router
        # (OS_EMAIL vs OS_APP_OPEN scored ~0.82/0.79, margin ~0.03 < the
        # 0.08 minimum) — see jarvis_phase4_nlp_routing_plan.md Phase 3.
        decision = decision_only("open")
        self.assertNotEqual(
            decision.decision, "execute",
            msg=f"Near-tie utterance auto-executed: {decision}",
        )

    def test_clean_match_still_executes(self):
        decision = decision_only("open chrome")
        self.assertEqual(decision.decision, "execute")
        self.assertEqual(decision.intent, "OS_APP_OPEN")


if __name__ == "__main__":
    unittest.main()
