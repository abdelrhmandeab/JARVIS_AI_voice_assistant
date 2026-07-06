"""Shared VAD timing helpers used by both audio/mic.py and audio/streaming_stt.py.

Pure functions only — no module-level state. Runtime-mutable VAD settings
(get_runtime_vad_settings / set_runtime_vad_settings) remain owned by
audio/mic.py, since that is the module callers already import for tuning
(see core/handlers/voice.py's mic_capture.set_runtime_vad_settings).
"""

from __future__ import annotations

import math
from typing import Any, Dict, Mapping, Optional

import numpy as np


def seconds_to_chunks(seconds: float, *, sample_rate: int, chunk_size: int) -> int:
    if seconds <= 0:
        return 1
    samples = int(seconds * sample_rate)
    return max(1, int(math.ceil(samples / float(chunk_size))))


def chunk_rms(chunk: np.ndarray) -> float:
    normalized = chunk.astype(np.float32) / 32768.0
    return float(np.sqrt(np.mean(np.square(normalized))))


def resolve_vad_mode(vad_mode: Optional[str]) -> str:
    mode = str(vad_mode or "command").strip().lower()
    if mode in {"chat", "conversation", "dialog", "turn"}:
        return "chat"
    return "command"


def resolve_silence_seconds(
    vad_mode: Optional[str],
    explicit_silence_seconds: Optional[float] = None,
    *,
    runtime: Mapping[str, Any],
    command_default: float,
    chat_default: float,
) -> float:
    """Resolve the base silence-cutoff seconds for a recording turn.

    `runtime` should be the live settings dict (audio.mic.get_runtime_vad_settings())
    so callers pick up any profile applied via set_runtime_vad_settings() — a
    static config fallback here previously let streaming_stt.py silently ignore
    runtime VAD profile changes that mic.py honored.
    """
    if explicit_silence_seconds is not None:
        return max(0.05, float(explicit_silence_seconds))

    mode = resolve_vad_mode(vad_mode)
    if mode == "chat":
        return float(runtime.get("chat_silence_seconds") or runtime.get("silence_seconds") or chat_default)
    return float(runtime.get("command_silence_seconds") or runtime.get("silence_seconds") or command_default)


def adaptive_silence_seconds(base_seconds: float, speech_seconds: float, max_seconds: float) -> float:
    """Scale silence threshold up with accumulated speech so long utterances get more grace.

    Denominator of 6.0 means only genuinely long utterances (6s+) reach max
    grace — a 3.0 denominator stretched a short command's cutoff toward the
    chat value after just 3s of speech, which is most short OS commands.
    """
    fraction = min(1.0, speech_seconds / 6.0)
    return base_seconds + (max_seconds - base_seconds) * fraction
