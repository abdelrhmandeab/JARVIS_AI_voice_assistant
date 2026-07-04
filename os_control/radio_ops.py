"""Windows radio (Wi-Fi / Bluetooth / Airplane mode) control.


Primary path: Windows.Devices.Radios WinRT API via the `winsdk` package.
No admin rights required.

Fallback path: PowerShell scripts — still no admin for Wi-Fi adapter toggling
via netsh, but BT may need elevation on some machines.
"""
from __future__ import annotations

import asyncio
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    import winsdk.windows.devices.radios as _wdr_type  # noqa: F401

import time

from core.config import CONTROLS_VERIFY_STATE, RADIO_BACKEND, AIRPLANE_RESTORE_RADIOS
from core.logger import get_logger

logger = get_logger("oscontrol")

# Radio state changes are not instantaneous — Wi-Fi adapter re-enable in
# particular goes through Disabled -> Enabling -> Up and can take several
# seconds (driver reinit, DHCP), far longer than a Bluetooth radio flip or
# a WinRT state change. Poll instead of a single fixed-delay check so a
# real success isn't misreported as a failure just because we looked too
# soon.
_VERIFY_POLL_INTERVAL_SECONDS = 0.4
_VERIFY_POLL_TIMEOUT_SECONDS = 6.0


def _poll_until(check_fn, expected: bool, *, timeout: float = _VERIFY_POLL_TIMEOUT_SECONDS) -> Optional[bool]:
    """Call check_fn() repeatedly until it returns `expected`, or timeout.

    Returns the last observed value (True/False), or None if check_fn never
    produced a determinate reading (e.g. device not found at all).
    """
    deadline = time.monotonic() + timeout
    last: Optional[bool] = None
    while True:
        last = check_fn()
        if last == expected:
            return last
        if time.monotonic() >= deadline:
            return last
        time.sleep(_VERIFY_POLL_INTERVAL_SECONDS)

# Module-level snapshot of radio states before airplane mode was engaged,
# so we can restore them on "airplane off".
_pre_airplane_states: dict[str, bool] = {}

# ──────────────────────────────────────────────
# WinRT helpers (winsdk)
# ──────────────────────────────────────────────

def _winrt_available() -> bool:
    if sys.platform != "win32":
        return False
    if RADIO_BACKEND == "powershell":
        return False
    try:
        import winsdk.windows.devices.radios as _wdr  # noqa: F401, type: ignore[import-not-found]
        return True
    except Exception:
        return False


def _run_async(coro):
    """Run an async coroutine synchronously, creating a loop if needed."""
    try:
        loop = asyncio.get_event_loop()
        if loop.is_closed():
            raise RuntimeError("closed")
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
    return loop.run_until_complete(coro)


async def _get_radios_async():
    import winsdk.windows.devices.radios as wdr  # type: ignore[import-not-found]
    return await wdr.Radio.get_radios_async()


def _winrt_get_radios() -> list:
    try:
        return list(_run_async(_get_radios_async()))
    except Exception as exc:
        logger.debug("WinRT get_radios failed: %s", exc)
        return []


async def _set_radio_state_async(radio, on: bool):
    import winsdk.windows.devices.radios as wdr  # type: ignore[import-not-found]
    state = wdr.RadioState.ON if on else wdr.RadioState.OFF
    await radio.set_state_async(state)


def _normalize_radio_kind(name: str) -> str:
    """Normalize a radio kind name for comparison. winsdk's RadioKind enum
    member name is e.g. 'WI_FI' (underscore, all-caps) while our own code
    uses 'Wi-Fi' (hyphen) — strip separators entirely so both compare equal
    regardless of which convention either side happens to use."""
    return "".join(ch for ch in str(name or "").lower() if ch.isalnum())


def _winrt_radio_is_on(kind_name: str) -> Optional[bool]:
    """Read back the current state of radios matching kind_name, or None if
    no matching radio could be found/queried."""
    try:
        import winsdk.windows.devices.radios as wdr  # type: ignore[import-not-found]
        target = _normalize_radio_kind(kind_name)
        radios = _winrt_get_radios()
        matched = [r for r in radios if _normalize_radio_kind(r.kind.name) == target]
        if not matched:
            return None
        return all(r.state == wdr.RadioState.ON for r in matched)
    except Exception as exc:
        logger.debug("WinRT radio state read failed for %s: %s", kind_name, exc)
        return None


def _winrt_set_radio(kind_name: str, on: bool) -> bool:
    """Toggle a radio by kind name ('Wi-Fi' or 'Bluetooth').

    Returns True only after the requested state is read back and confirmed —
    a set_state_async() call that doesn't throw is not itself success.
    """
    try:
        target = _normalize_radio_kind(kind_name)
        radios = _winrt_get_radios()
        matched = [r for r in radios if _normalize_radio_kind(r.kind.name) == target]
        if not matched:
            logger.debug("No WinRT radio found for kind=%s", kind_name)
            return False
        for radio in matched:
            _run_async(_set_radio_state_async(radio, on))
    except Exception as exc:
        logger.debug("WinRT set_radio(%s, %s) failed: %s", kind_name, on, exc)
        return False

    if not bool(CONTROLS_VERIFY_STATE):
        return True

    actual = _poll_until(lambda: _winrt_radio_is_on(kind_name), on)
    if actual is None:
        logger.warning("WinRT radio %s: could not read back state to verify.", kind_name)
        return False
    if actual != on:
        logger.warning(
            "WinRT radio %s set to %s but readback shows %s after polling — reporting failure.",
            kind_name,
            "on" if on else "off",
            "on" if actual else "off",
        )
        return False
    return True


def _winrt_snapshot_radios() -> dict[str, bool]:
    """Return {kind_name: is_on} for all current radios."""
    try:
        import winsdk.windows.devices.radios as wdr  # type: ignore[import-not-found]
        radios = _winrt_get_radios()
        return {r.kind.name: (r.state == wdr.RadioState.ON) for r in radios}
    except Exception as exc:
        logger.debug("WinRT snapshot_radios failed: %s", exc)
        return {}


# ──────────────────────────────────────────────
# PowerShell fallback helpers
# ──────────────────────────────────────────────

def _ps_run(script: str, timeout: int = 15) -> tuple[bool, str]:
    """Run a PowerShell snippet, return (success, output/error).

    PnP/NetAdapter cmdlets cold-start slowly (driver-store enumeration) on
    some machines — 15s default gives real hardware room without letting a
    genuinely hung call block forever.
    """
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        combined = (result.stdout + result.stderr).strip()
        return result.returncode == 0, combined
    except Exception as exc:
        return False, str(exc)


def _ps_wifi_adapter_state() -> Optional[bool]:
    """Return True if any Wi-Fi-named interface is enabled, False if all are
    disabled, or None if no matching interface/state could be determined.
    Requires admin to change but not to read.
    """
    script = (
        "Get-NetAdapter | Where-Object { $_.Name -match 'Wi-?Fi|Wireless' } | "
        "Select-Object -ExpandProperty Status"
    )
    ok, out = _ps_run(script)
    if not ok or not out.strip():
        return None
    statuses = [line.strip().lower() for line in out.splitlines() if line.strip()]
    if not statuses:
        return None
    return any(status == "up" for status in statuses)


def _ps_wifi(on: bool) -> bool:
    # Adapter enable/disable requires admin; netsh interface swallows errors
    # for non-admin callers, so the only honest signal is a readback.
    script = (
        "netsh interface set interface name="
        + '"Wi-Fi"'
        + f" admin={'enabled' if on else 'disabled'} 2>$null; "
        + "netsh interface set interface name="
        + '"Wireless Network Connection"'
        + f" admin={'enabled' if on else 'disabled'} 2>$null"
    )
    _ps_run(script)

    if not bool(CONTROLS_VERIFY_STATE):
        return True

    # Wi-Fi adapter re-enable is the slowest radio transition (driver
    # reinit + DHCP) — give it a longer poll window than Bluetooth/WinRT.
    actual = _poll_until(_ps_wifi_adapter_state, on, timeout=10.0)
    if actual is None:
        logger.warning("PS wifi: no Wi-Fi adapter found to verify state.")
        return False
    if actual != on:
        logger.warning(
            "PS wifi set to %s but adapter readback shows %s after polling — reporting "
            "failure (likely needs admin).",
            "on" if on else "off",
            "on" if actual else "off",
        )
        return False
    return True


# Get-PnpDevice -Class Bluetooth returns the radio adapter itself AND every
# paired peripheral (headsets, speakers, AVRCP/RFCOMM transport endpoints).
# Only the adapter actually controls whether Bluetooth is on — its InstanceId
# is USB/PCI-enumerated (e.g. "USB\VID_..."), while paired peripherals are
# bus-enumerated as "BTHENUM\..." or "BTH\...". Filtering to the adapter
# avoids both misreading state from a disconnected earbud ("Unknown" status)
# and accidentally disabling a paired peripheral instead of the radio.
_ps_bluetooth_adapter_filter = (
    "Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue | "
    "Where-Object { $_.InstanceId -match '^USB' }"
)


def _ps_bluetooth_state() -> Optional[bool]:
    """Return True if the Bluetooth radio adapter is enabled (Status=OK),
    False if it's present but disabled/errored, or None if no adapter was
    found at all (e.g. no Bluetooth hardware)."""
    script = f"{_ps_bluetooth_adapter_filter} | Select-Object -ExpandProperty Status"
    ok, out = _ps_run(script)
    if not ok or not out.strip():
        return None
    statuses = [line.strip().lower() for line in out.splitlines() if line.strip()]
    if not statuses:
        return None
    return any(status == "ok" for status in statuses)


def _ps_bluetooth(on: bool) -> bool:
    action = "enable" if on else "disable"
    script = (
        f"{_ps_bluetooth_adapter_filter} | "
        f"ForEach-Object {{ {action.capitalize()}-PnpDevice -InstanceId $_.InstanceId -Confirm:$false -ErrorAction SilentlyContinue }}"
    )
    _ps_run(script)

    if not bool(CONTROLS_VERIFY_STATE):
        return True

    actual = _poll_until(_ps_bluetooth_state, on)
    if actual is None:
        logger.warning("PS bluetooth: no Bluetooth radio adapter found to verify state.")
        return False
    if actual != on:
        logger.warning(
            "PS bluetooth set to %s but adapter readback shows %s after polling — reporting "
            "failure (likely needs admin).",
            "on" if on else "off",
            "on" if actual else "off",
        )
        return False
    return True


# ──────────────────────────────────────────────
# Public API
# ──────────────────────────────────────────────

def set_radio(kind: str, on: bool) -> bool:
    """Toggle Wi-Fi or Bluetooth.

    kind: 'wifi' | 'bluetooth'
    Returns True on success, False on failure (never raises).
    """
    kind_lower = kind.lower()
    winrt_name = "Wi-Fi" if kind_lower == "wifi" else "Bluetooth"

    if _winrt_available():
        ok = _winrt_set_radio(winrt_name, on)
        if ok:
            logger.info("WinRT radio %s -> %s", kind, "on" if on else "off")
            return True
        logger.debug("WinRT radio toggle failed, falling back to PowerShell")

    # PowerShell fallback
    if kind_lower == "wifi":
        ok = _ps_wifi(on)
    else:
        ok = _ps_bluetooth(on)

    if ok:
        logger.info("PS fallback radio %s -> %s", kind, "on" if on else "off")
    else:
        logger.warning("All radio backends failed for %s -> %s", kind, "on" if on else "off")
    return ok


def set_airplane(on: bool, restore: bool | None = None) -> bool:
    """Toggle airplane mode (all radios off/on).

    When on=True, snapshots current radio states then turns all off.
    When on=False and restore=True (default from config), re-applies snapshot.

    Returns True only if EVERY radio that was supposed to change actually
    verified the change — a partial result (e.g. Bluetooth restored but
    Wi-Fi silently stayed off) must never be reported as "Radios restored."
    """
    global _pre_airplane_states

    if restore is None:
        restore = AIRPLANE_RESTORE_RADIOS

    if on:
        # Snapshot current states
        if _winrt_available():
            _pre_airplane_states = _winrt_snapshot_radios()
            logger.debug("Airplane ON: snapshot=%s", _pre_airplane_states)

        # Turn both radios off in parallel to avoid sequential ~9s PS calls
        with ThreadPoolExecutor(max_workers=2) as pool:
            f_wifi = pool.submit(set_radio, "wifi", False)
            f_bt = pool.submit(set_radio, "bluetooth", False)
            wifi_ok = f_wifi.result()
            bt_ok = f_bt.result()
        if not (wifi_ok and bt_ok):
            logger.warning(
                "Airplane mode ON partial: wifi_ok=%s bluetooth_ok=%s — reporting failure.",
                wifi_ok, bt_ok,
            )
        return wifi_ok and bt_ok

    else:
        # Airplane OFF
        if restore and _pre_airplane_states:
            items = list(_pre_airplane_states.items())
            _pre_airplane_states.clear()
            with ThreadPoolExecutor(max_workers=2) as pool:
                futures = {
                    pool.submit(set_radio, "wifi" if "wifi" in _normalize_radio_kind(k) else "bluetooth", v): k
                    for k, v in items
                }
                results = {k: f.result() for f, k in futures.items()}
            if not all(results.values()):
                logger.warning(
                    "Airplane mode OFF partial restore: %s — reporting failure.",
                    results,
                )
            return all(results.values())
        else:
            # No snapshot — just turn both on in parallel
            with ThreadPoolExecutor(max_workers=2) as pool:
                f_wifi = pool.submit(set_radio, "wifi", True)
                f_bt = pool.submit(set_radio, "bluetooth", True)
                wifi_ok = f_wifi.result()
                bt_ok = f_bt.result()
            _pre_airplane_states.clear()
            if not (wifi_ok and bt_ok):
                logger.warning(
                    "Airplane mode OFF partial: wifi_ok=%s bluetooth_ok=%s — reporting failure.",
                    wifi_ok, bt_ok,
                )
            return wifi_ok and bt_ok


def get_radio_states() -> dict[str, bool]:
    """Return current radio states for diagnostics. Returns {} if unavailable."""
    if _winrt_available():
        snapshot = _winrt_snapshot_radios()
        if snapshot:
            return {_normalize_radio_kind(k): v for k, v in snapshot.items()}
    return {}
