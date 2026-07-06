import json
import logging
import os
import shutil
import sys
import time
from logging.handlers import RotatingFileHandler

from core.config import (
    LOG_CONSOLE_LEVEL,
    LOG_FILE,
    LOG_FILE_LEVEL,
    LOG_PRETTY,
    LOG_ROTATE_BACKUPS,
    LOG_ROTATE_MAX_BYTES,
)


def _enable_windows_vt_mode() -> bool:
    """Best-effort: turn on ANSI escape processing in the Windows console.

    Windows 10+ consoles support VT sequences once ENABLE_VIRTUAL_TERMINAL_PROCESSING
    is set on the output handle; older consoles (or non-tty output, e.g. redirected
    to a file) just keep printing the raw escape codes, so callers should only rely
    on this after confirming isatty().
    """
    if os.name != "nt":
        return True
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE (StreamHandler -> sys.stderr
        # shares the same console, but the console mode is a single per-console
        # setting, so enabling it via either standard handle applies to both).
        mode = ctypes.c_uint32()
        if not kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            return False
        new_mode = mode.value | 0x0004  # ENABLE_VIRTUAL_TERMINAL_PROCESSING
        return bool(kernel32.SetConsoleMode(handle, new_mode))
    except Exception:
        return False


class _ColorConsoleFormatter(logging.Formatter):
    """Color-coded, aligned console formatter. File logs stay plain (no ANSI)."""

    COLORS = {
        "DEBUG": "\033[2;37m",
        "INFO": "\033[36m",
        "WARNING": "\033[33m",
        "ERROR": "\033[31m",
        "CRITICAL": "\033[1;41m",
    }
    RESET = "\033[0m"
    DIM = "\033[2m"

    def format(self, record: logging.LogRecord) -> str:
        color = self.COLORS.get(record.levelname, "")
        ts = f"{self.DIM}{self.formatTime(record, '%H:%M:%S')}{self.RESET}"
        name = f"{record.name:<14.14}"
        return f"{ts} {color}{record.levelname:<7}{self.RESET} {self.DIM}{name}{self.RESET} {record.getMessage()}"


def _level(value, default):
    resolved = logging.getLevelName(str(value or "").strip().upper())
    return resolved if isinstance(resolved, int) else default


class _WindowsSafeRotatingFileHandler(RotatingFileHandler):
    """RotatingFileHandler that avoids WinError 32 on Windows.

    The standard handler uses os.rename() which fails when another process
    holds a handle on the log file.  This subclass copies the current log to
    the backup slot and then truncates the original in-place, so the file
    descriptor stays valid for all open handles.
    """

    def rotate(self, source: str, dest: str) -> None:
        if os.path.exists(dest):
            try:
                os.remove(dest)
            except OSError:
                pass
        try:
            shutil.copy2(source, dest)
            # Truncate the source file in-place so existing handles still work.
            with open(source, "w", encoding="utf-8"):
                pass
        except OSError:
            # Fall back to the default rename behaviour on non-Windows.
            super().rotate(source, dest)


logger = logging.getLogger("jarvis")

if not logger.handlers:
    console_level = _level(LOG_CONSOLE_LEVEL, logging.INFO)
    file_level = _level(LOG_FILE_LEVEL, logging.DEBUG)
    logger.setLevel(min(console_level, file_level))
    logger.propagate = False

    formatter = logging.Formatter("%(asctime)s | %(levelname)-7s | %(name)s | %(message)s")

    file_handler = _WindowsSafeRotatingFileHandler(
        LOG_FILE,
        maxBytes=max(1, int(LOG_ROTATE_MAX_BYTES)),
        backupCount=max(0, int(LOG_ROTATE_BACKUPS)),
        encoding="utf-8",
        delay=True,
    )
    file_handler.setLevel(file_level)
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(console_level)
    _use_color = bool(LOG_PRETTY) and hasattr(sys.stderr, "isatty") and sys.stderr.isatty()
    if _use_color:
        _use_color = _enable_windows_vt_mode()
    console_handler.setFormatter(_ColorConsoleFormatter() if _use_color else formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)


def get_logger(component: str) -> logging.Logger:
    """Return a component-tagged child of the shared Jarvis logger."""
    name = str(component or "general").strip().replace(" ", "_") or "general"
    return logger.getChild(name)


def section(title: str) -> None:
    logger.info("──────── %s ────────", str(title or "").strip())


def turn_separator(turn: int, **fields) -> None:
    """Log a compact rule at the start of each voice turn so the terminal
    reads as conversation blocks. Gated behind JARVIS_LOG_PRETTY (default on)
    so CI/file logs stay uncluttered when it's turned off.
    """
    if not LOG_PRETTY:
        return
    suffix = " ".join(f"{key}={value}" for key, value in fields.items() if value not in (None, ""))
    logger.info("────── turn ⸸ %d%s ──────", int(turn), f" {suffix}" if suffix else "")


def kv(component: str, **pairs) -> None:
    width = max((len(str(key)) for key in pairs), default=0)
    message = "  ".join(f"{str(key):<{width}}={value}" for key, value in pairs.items())
    get_logger(component).info(message)


def summary_table(title: str, rows: list[tuple]) -> None:
    section(title)
    normalized = [tuple(str(value) for value in row) for row in rows]
    if not normalized:
        return
    column_count = max(len(row) for row in normalized)
    widths = [
        max((len(row[index]) if index < len(row) else 0) for row in normalized)
        for index in range(column_count)
    ]
    table_logger = get_logger(str(title or "summary").strip().lower().replace(" ", "_"))
    for row in normalized:
        table_logger.info(
            "  ".join(
                (row[index] if index < len(row) else "").ljust(widths[index])
                for index in range(column_count)
            ).rstrip()
        )


def log_structured(event, level="debug", **fields):
    payload = {
        "event": str(event or "unknown"),
        "timestamp": float(fields.pop("timestamp", time.time())),
    }
    for key, value in fields.items():
        payload[str(key)] = value

    message = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    # Structured payloads are intentionally file-detail only; the ``level``
    # argument remains accepted for compatibility with existing callers.
    logger.debug(message)
    return payload
