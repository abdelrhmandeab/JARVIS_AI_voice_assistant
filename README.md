# Jarvis

A local-first Windows voice assistant. Bilingual — English and Egyptian Arabic —
runs on any Windows 10/11 PC with 8GB+ RAM. No GPU required. Most controls work
without Administrator rights.

## Architecture

```text
Wake word (unified ONNX) → Streaming capture + STT → Intent routing cascade → Action / LLM → TTS
```

- **Wake word**: a single custom openWakeWord-compatible ONNX model detects both
  "Jarvis" (English) and "جارفيس" (Egyptian Arabic) — one model, not per-language
  variants. Detection uses an EMA-smoothed score plus a two-level threshold
  (trigger + peak) with debounce and cooldown, so a single flat threshold can't
  cause the classic "fires on noise, misses real speech" failure mode. An
  adaptive-retraining daemon (`core/adaptive_wake.py`) accumulates confirmed
  detections and false positives in the background and periodically retrains
  and hot-swaps the model once enough samples are collected and the new model
  clears a validation-accuracy floor.
- **STT**: hybrid ElevenLabs (cloud) + local Faster-Whisper fallback, hardware-sized
  automatically (`medium`/`small`/`base` depending on VRAM/RAM). Mixed English +
  Egyptian Arabic in one sentence is treated as the *normal* case, not an edge
  case — the primary language is locked per utterance (never decoded with
  `language=None`), but code-switched words in the other script are accepted
  rather than rejected. A confidence floor discards short, low-confidence
  transcripts as noise instead of passing along a hallucinated sentence, and a
  separate hallucination guard (avg log-probability + compression ratio) drops
  segments that look confabulated.
- **Intent routing** (cascade — earliest tier that resolves the command wins):
  1. **Parser fast-path** (~0 ms) — regex/keyword parser, high-confidence bilingual patterns.
  2. **Code-switch router** (~2 ms) — dictionary/token match for mixed-language commands.
  3. **Semantic router** (~10 ms) — multilingual MiniLM embeddings, handles paraphrases.
  4. **Keyword NLP** — fuzzy match (rapidfuzz) for noisy STT output.
  5. **Tool-calling LLM tier** — local Ollama (or Claude API, optional) with structured tool calls.
  6. **Structured LLM NLU fallback** (opt-in, off by default) — schema-constrained
     LLM output, verified by `core/route_verifier.py` before execution.
  7. **General LLM chat** — anything that doesn't resolve to a command.
- **Controls verify real OS state before reporting success** — see
  [Controls](#controls) below.
- **File commands execute, don't narrate** — search/find/list open File Explorer
  directly; responses never speak a raw filesystem path.
- **TTS**: ElevenLabs (natural) with edge-tts fallback (`ar-EG-SalmaNeural` /
  `en-US-AriaNeural`), sentence-level streaming synthesis so playback starts
  before the whole response is ready, a deterministic (non-LLM) voice
  normalizer that converts numbers/dates/times/place names to spoken form, and
  a persona layer (`core/persona.py`) that shapes tone without LLM calls.

## Controls

Every control re-reads the actual OS state after acting and only reports
success if it verifiably changed — never "the API call didn't throw, so it
must have worked."

| Control | Verified how | Needs admin? |
| --- | --- | --- |
| System volume / mute | pycaw `IAudioEndpointVolume`, read back within 2% | No |
| Wi-Fi / Bluetooth / Airplane mode | WinRT Radios (`winsdk`) primary, state polled after set; PowerShell/`netsh`/`Get-PnpDevice` fallback, also polled | No (WinRT path) |
| Night light | Registry blob write, read back | No |
| Do Not Disturb / Focus Assist | Registry write, read back | No |
| Energy saver | `powercfg /setactive`, verified via `powercfg /getactivescheme` | No |
| Live captions | `Win+Ctrl+L` hotkey / `WM_CLOSE`, verified by enumerating the actual "Live Captions" window | No |
| Brightness, lock, sleep, screenshot | Best-effort (no readback path exists for these) | No |
| Wi-Fi/Bluetooth **adapter** disable (PowerShell fallback path) | `Enable/Disable-NetAdapter`, `Enable/Disable-PnpDevice` | **Yes** |

If a control genuinely can't be verified as changed, Jarvis says so honestly
("I couldn't change that — it may need Administrator rights") instead of
claiming success.

### Running elevated

Jarvis doesn't ship a UAC manifest or an installer — elevation is manual:

1. **Run as Administrator per launch**: right-click the terminal/shortcut you
   start Jarvis from → **Run as administrator**. A startup log line records
   whether Jarvis is currently elevated.
2. **Always-elevated without a UAC prompt**: create a Windows **Scheduled
   Task** set to "Run with highest privileges," launched at logon, that starts
   `python main.py` in the project directory.

Most controls (volume, radios via WinRT, night light, DND, energy saver, live
captions) work fine without either — elevation is only needed for the
PowerShell adapter-disable fallback path.

## PIN confirmation

Destructive or sensitive actions (permanent delete, moving/renaming files,
system shutdown, etc.) ask for a spoken PIN before executing — nothing is
read back that could be overheard and replayed, unlike a spoken confirmation
token. Default PIN is `1234` (change it via `JARVIS_SECOND_FACTOR_PIN`);
repeated wrong attempts trigger a lockout.

## Hardware-aware model selection

At startup, [core/hardware_detect.py](core/hardware_detect.py) reads RAM + GPU and
picks the best Qwen3 model that fits. Override via `JARVIS_LLM_MODEL` in `.env`
(for example `qwen2.5:3b`), or disable with `JARVIS_LLM_AUTO_SELECT=false`.

| RAM    | GPU | Tier      | Model        | num_ctx | lightweight_ctx |
|--------|-----|-----------|--------------|---------|-----------------|
| 16 GB+ | Yes | `high`    | qwen3:8b     | 8192    | 4096            |
| 12 GB+ | Any | `medium`  | qwen3:4b     | 4096    | 2048            |
| 8 GB   | Yes | `medium`  | qwen3:4b     | 4096    | 2048            |
| 8 GB   | No  | `low`     | qwen3:1.7b   | 2048    | 1024            |
| < 8 GB | Any | `minimal` | qwen3:0.6b   | 1024    | 512             |

Missing models are auto-pulled on first run with streaming progress logs —
no manual `ollama pull` needed.

## Setup

### 1. Install dependencies

```powershell
python -m pip install -r requirements.txt
```

For wake-word model retraining/development only (not needed to run Jarvis):

```powershell
python -m pip install -r requirements-training.txt
```

### 2. Copy and edit the environment template

```powershell
copy .env.example .env
notepad .env
```

Every key has a working default. The ones worth reviewing:

- `ELEVENLABS_API_KEY` — only if you want cloud-quality Arabic STT/TTS.
  Without it, Jarvis falls back to local Whisper + edge-tts automatically.
- `JARVIS_SECOND_FACTOR_PIN` — change the default PIN used for destructive-action confirmation.
- `JARVIS_WAKE_WORD_UNIFIED_ONNX_PATH` — point at a retrained wake-word model if you've trained your own.

### 3. Start Ollama

```powershell
ollama serve
```

You don't need to manually `ollama pull` anything — the orchestrator auto-pulls
the configured / hardware-recommended model on first launch.

## Run

```powershell
python main.py
```

Say **"Jarvis"** or **"جارفيس"** — both are detected by the same model. The
greeting plays fully before the wake-word listener opens the microphone, so
there's no race between the two.

## Health check

Run the doctor to verify all dependencies, audio devices, Ollama, and feature tiers:

```powershell
python core/doctor.py
```

The report flags each optional feature as `available` or `degraded (...)` so you
can see exactly which extras are active without grepping logs.

## Optional dependencies

All of these gracefully degrade if missing — Jarvis still starts:

| Package | Feature |
| --- | --- |
| `pyperclip` | Clipboard read/write |
| `screen-brightness-control` | DDC/CI brightness (no admin) |
| `sentence-transformers` | Semantic intent routing, knowledge-base embeddings |
| `duckduckgo-search` / `ddgs` | Web search for live data |
| `pywin32` | Outlook email/calendar drafts + Windows Search Index |
| `winsdk` | WinRT radio control (Wi-Fi/Bluetooth/Airplane, no admin) |
| `chromadb` | Vector semantic memory |
| `faiss` / `sentence-transformers` | Offline knowledge base retrieval |

## Live data

Weather queries hit Open-Meteo (no API key). Search/news/price queries hit
DuckDuckGo (no API key). Results are injected into the LLM prompt as a `LIVE DATA`
section so the model can answer with real, current information instead of an
"I don't have live data" apology.

## Memory

SQLite (`data/memory/jarvis_memory.db`) is the primary store — an append-only
turn log plus a key/value slot table. `jarvis_memory.json` only exists for
one-time legacy migration and manual debug export; it is not the active store.
A separate vector store (ChromaDB + `all-MiniLM-L6-v2` embeddings, under
`data/vectors/`) provides semantic recall for the LLM context path, written
asynchronously so it never blocks the fast path.

## Data directory layout

```text
data/
  logs/       — structured + action logs
  memory/     — SQLite conversation/slot store
  index/      — Windows Search Index cache
  state/      — misc runtime state
  kb/         — offline knowledge-base FAISS index
  vectors/    — vector semantic memory (ChromaDB)
  wake_samples/ — wake-word enrollment + adaptive-retraining samples (local only, git-ignored)
```

## Project layout

```text
core/         — orchestrator (main loop), command router (routing cascade),
                config, memory (SQLite + vector), persona/voice-normalizer/
                prosody, knowledge base, elevation check, adaptive wake
                retraining, route verifier
core/handlers/— per-domain command handlers (file navigation, batch ops,
                job queue, knowledge base, memory, persona, voice)
audio/        — wake word (unified ONNX), STT (hybrid ElevenLabs/Whisper),
                TTS (SpeechEngine, sentence streaming), VAD, mic capture
llm/          — Ollama client, prompt builder, tool-calling tier, structured
                NLU fallback
nlp/          — semantic router, fuzzy matcher, keyword engine, code-switch
                router, entity types
os_control/   — Windows integration: volume/brightness/lock (native_ops),
                Wi-Fi/Bluetooth/Airplane (radio_ops), night light/DND/energy
                saver (windows_toggles), file ops + Explorer integration,
                PIN confirmation, app open/close, email/calendar, clipboard,
                screen capture, policy/risk gating
tools/        — weather (Open-Meteo), web_search (DuckDuckGo), calculator
ui/           — system tray + WebSocket bridge to the desktop UI
desktop/      — Tauri-based desktop UI (optional, separate build)
scripts/      — wake-word training/maintenance scripts
models/       — the shipped wake-word ONNX model
```

## Training data

Wake-word training clips, enrollment samples, and intermediate model
artifacts live under `data/wake_samples/` and are git-ignored — they stay on
your machine for retraining but are never committed. Only the finished
production model under `models/` is tracked.
