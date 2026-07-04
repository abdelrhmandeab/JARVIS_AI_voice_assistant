"""Windows admin-elevation awareness.

Several controls (Disable-NetAdapter, Disable-PnpDevice for Bluetooth) only
work when Jarvis itself is running elevated. This module answers "are we
elevated right now" so honest-failure messages can tell the user precisely
what to do instead of guessing.
"""
from __future__ import annotations

import ctypes
import platform

_IS_WINDOWS = platform.system().lower() == "windows"

_is_admin_cache: bool | None = None


def is_admin() -> bool:
    """Return True if the current process has Administrator privileges."""
    global _is_admin_cache
    if _is_admin_cache is not None:
        return _is_admin_cache
    if not _IS_WINDOWS:
        _is_admin_cache = False
        return _is_admin_cache
    try:
        _is_admin_cache = bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        _is_admin_cache = False
    return _is_admin_cache
