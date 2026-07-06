import threading
import time

from core.config import SENSITIVE_PIN_PENDING_TIMEOUT_SECONDS
from core.logger import logger
from os_control.action_log import log_action
from os_control.second_factor import (
    clear_confirmation_attempts,
    clear_second_factor_attempts,
    is_confirmation_allowed,
    normalize_spoken_pin,
    record_confirmation_attempt,
    verify_second_factor,
)

# Single-slot in-memory store for the pending PIN-confirmed action. There is
# only ever one outstanding sensitive command at a time (the user must speak
# the PIN before issuing another), so a fixed, non-spoken key is enough —
# unlike the old hex-token system, nothing here is ever read aloud.
_PIN_PENDING_KEY = "pin_pending"
_PIN_SENTINEL = "PIN_REQUIRED"
_pin_lock = threading.Lock()
_pin_pending_action = None


class ConfirmationManager:
    def __init__(self):
        self.pin_timeout_seconds = SENSITIVE_PIN_PENDING_TIMEOUT_SECONDS

    def create(self, action_name, description, payload):
        """Store the pending sensitive action and return the PIN_REQUIRED sentinel.

        Replaces the old hex-token flow: nothing here is spoken back to the
        user, so there is no token to leak or mistype.
        """
        global _pin_pending_action
        now_ts = time.time()
        expires_at = now_ts + max(5, int(self.pin_timeout_seconds or 30))

        with _pin_lock:
            _pin_pending_action = {
                "action_name": action_name,
                "description": description,
                "payload": payload,
                "created_at": now_ts,
                "expires_at": expires_at,
            }

        logger.info("PIN confirmation requested for %s", action_name)
        return _PIN_SENTINEL

    def has_pending_pin_action(self):
        with _pin_lock:
            pending = _pin_pending_action
        if not pending:
            return False
        if time.time() > pending["expires_at"]:
            return False
        return True

    def pending_pin_description(self):
        with _pin_lock:
            pending = _pin_pending_action
        return str((pending or {}).get("description") or "")

    def discard_pending_pin_action(self):
        global _pin_pending_action
        with _pin_lock:
            had_pending = _pin_pending_action is not None
            _pin_pending_action = None
        return had_pending

    def verify_pin_and_execute(self, spoken_pin):
        """Verify a spoken PIN against the pending action.

        Returns (status, message, payload) where status is one of:
          "executed"  -> payload is the stored action payload to run
          "wrong"     -> message explains the failure, retry allowed
          "locked"    -> too many wrong attempts; pending action discarded
          "no_pending"-> nothing was waiting on a PIN
        """
        global _pin_pending_action

        with _pin_lock:
            pending = _pin_pending_action

        if not pending:
            return "no_pending", "No pending action requires a PIN.", None

        if time.time() > pending["expires_at"]:
            self.discard_pending_pin_action()
            clear_second_factor_attempts(_PIN_PENDING_KEY)
            log_action(
                "confirmation_rejected",
                "failed",
                details={"action_name": pending.get("action_name"), "reason": "pin_expired"},
            )
            return "no_pending", "The PIN request expired.", None

        allowed, rate_message = is_confirmation_allowed(_PIN_PENDING_KEY)
        if not allowed:
            self.discard_pending_pin_action()
            clear_confirmation_attempts(_PIN_PENDING_KEY)
            clear_second_factor_attempts(_PIN_PENDING_KEY)
            log_action(
                "confirmation_rejected",
                "failed",
                details={"action_name": pending.get("action_name"), "reason": "pin_rate_limited"},
            )
            return "locked", rate_message, None

        normalized_pin = normalize_spoken_pin(spoken_pin)
        verified, factor_message = verify_second_factor(normalized_pin, token=_PIN_PENDING_KEY)

        if verified:
            self.discard_pending_pin_action()
            clear_confirmation_attempts(_PIN_PENDING_KEY)
            record_confirmation_attempt(_PIN_PENDING_KEY, success=True)
            log_action(
                "confirmation_accepted",
                "success",
                details={"action_name": pending.get("action_name")},
            )
            return "executed", "PIN accepted.", pending["payload"] or {}

        record_confirmation_attempt(_PIN_PENDING_KEY, success=False)
        still_allowed, _ = is_confirmation_allowed(_PIN_PENDING_KEY)
        if not still_allowed or "Too many failed" in (factor_message or ""):
            self.discard_pending_pin_action()
            clear_confirmation_attempts(_PIN_PENDING_KEY)
            log_action(
                "confirmation_rejected",
                "failed",
                details={"action_name": pending.get("action_name"), "reason": "pin_lockout"},
            )
            return "locked", factor_message or "Too many failed PIN attempts.", None

        log_action(
            "confirmation_rejected",
            "failed",
            details={"action_name": pending.get("action_name"), "reason": "pin_incorrect"},
        )
        return "wrong", "Wrong PIN.", None

    def cancel_pending_pin(self):
        pending_name = self.pending_pin_description()
        had_pending = self.discard_pending_pin_action()
        if not had_pending:
            return False, "No pending PIN request to cancel."
        clear_confirmation_attempts(_PIN_PENDING_KEY)
        clear_second_factor_attempts(_PIN_PENDING_KEY)
        log_action(
            "confirmation_cancelled",
            "success",
            details={"action_name": pending_name},
        )
        return True, "Pending PIN request cancelled."


confirmation_manager = ConfirmationManager()
