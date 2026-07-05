"""Optional WebSocket bridge between Jarvis engine events and desktop UI clients."""

from __future__ import annotations

import asyncio
import json
import threading
from typing import Any

from core import config
from core.dialogue_manager import dialogue_manager
from core.logger import get_logger
from ui.events import (
    COMMAND_CONFIG_REQUEST,
    COMMAND_FEATURE_FLAG,
    COMMAND_HEALTH_REQUEST,
    COMMAND_MUTE_TOGGLE,
    COMMAND_PIN_ATTEMPT,
    COMMAND_SETTING_UPDATE,
    COMMAND_TEXT,
    EVENT_CONFIG,
    EVENT_ERROR,
    EVENT_HEALTH,
    EVENT_NOTIFY,
    EVENT_PIN_RESULT,
    EVENT_RESPONSE,
    EVENT_STATE_CHANGED,
    make_event,
    to_json,
)

logger = get_logger("bridge")

try:
    from fastapi import FastAPI, WebSocket, WebSocketDisconnect
    import uvicorn
except ImportError as exc:  # pragma: no cover - exercised when optional deps are absent.
    FastAPI = None
    WebSocket = Any
    WebSocketDisconnect = Exception
    uvicorn = None
    _IMPORT_ERROR = exc
else:
    _IMPORT_ERROR = None


class JarvisBridge:
    def __init__(self) -> None:
        self.clients = set()
        self.lock = threading.Lock()
        self.loop = None
        self.muted = False
        self.enabled = bool(getattr(config, "UI_BRIDGE_ENABLED", True))
        self.host = str(getattr(config, "UI_BRIDGE_HOST", "127.0.0.1") or "127.0.0.1")
        self.port = int(getattr(config, "UI_BRIDGE_PORT", 9720) or 9720)
        self._running = False
        self._state_listener = None
        self._server_thread = None
        self._server = None
        self._start_lock = threading.Lock()
        self._missing_deps_logged = False

    @property
    def available(self) -> bool:
        return FastAPI is not None and uvicorn is not None

    @property
    def running(self) -> bool:
        return bool(self._running and self.loop is not None)

    def start(self) -> None:
        if not self.enabled:
            logger.info("UI bridge disabled.")
            return
        if not self.available:
            if not self._missing_deps_logged:
                logger.info("UI bridge optional dependencies unavailable; bridge disabled: %s", _IMPORT_ERROR)
                self._missing_deps_logged = True
            return

        with self._start_lock:
            if self._running:
                return
            app = self._create_app()
            self._state_listener = self._on_state_changed
            try:
                dialogue_manager.register_state_listener(self._state_listener)
            except Exception:
                logger.debug("Failed to register bridge state listener", exc_info=True)

            self._server_thread = threading.Thread(
                target=self._run_server,
                args=(app,),
                name="jarvis-ui-bridge",
                daemon=True,
            )
            self._running = True
            self._server_thread.start()
            logger.info("UI bridge starting on ws://%s:%s/ws", self.host, self.port)

    def stop(self, timeout: float = 3.0) -> None:
        """Signal the uvicorn server to exit and wait for its thread to stop.
        Safe to call even if the bridge was never started."""
        server = self._server
        loop = self.loop
        thread = self._server_thread
        if server is not None and loop is not None and not loop.is_closed():
            try:
                loop.call_soon_threadsafe(setattr, server, "should_exit", True)
            except Exception:
                logger.debug("Failed to signal UI bridge server to exit", exc_info=True)
        if thread is not None and thread.is_alive():
            thread.join(timeout=timeout)
        self._running = False

    def _create_app(self):
        app = FastAPI()

        @app.on_event("startup")
        async def _on_startup() -> None:
            self.loop = asyncio.get_running_loop()

        @app.websocket("/ws")
        async def _websocket_endpoint(websocket: WebSocket) -> None:
            await self.handle_client(websocket)

        return app

    def _run_server(self, app) -> None:
        try:
            # Build our own Server instead of uvicorn.run() so stop() has a
            # real handle to signal shutdown with — uvicorn.run() blocks with
            # no way to ask it to exit from another thread.
            server_config = uvicorn.Config(app, host=self.host, port=self.port, log_level="info")
            self._server = uvicorn.Server(server_config)
            self._server.run()
        except Exception:
            logger.exception("UI bridge server failed")
        finally:
            self._running = False
            self.loop = None
            self._server = None
            try:
                if self._state_listener is not None:
                    dialogue_manager.unregister_state_listener(self._state_listener)
            except Exception:
                logger.debug("Failed to unregister bridge state listener", exc_info=True)

    async def handle_client(self, websocket: WebSocket) -> None:
        try:
            await websocket.accept()
            self._register_client(websocket)
            await websocket.send_text(to_json(self._config_event()))
            while True:
                try:
                    raw_message = await websocket.receive_text()
                    await self._handle_message(websocket, raw_message)
                except WebSocketDisconnect:
                    break
                except Exception:
                    logger.exception("UI bridge websocket message failed")
                    try:
                        await websocket.send_text(
                            to_json(
                                make_event(
                                    EVENT_ERROR,
                                    message="Bridge failed to process websocket message.",
                                    recoverable=True,
                                )
                            )
                        )
                    except Exception:
                        logger.debug("Failed to send bridge error event", exc_info=True)
        except WebSocketDisconnect:
            pass
        except Exception:
            logger.exception("UI bridge websocket client failed")
        finally:
            self._unregister_client(websocket)

    async def _handle_message(self, websocket: WebSocket, raw_message: str) -> None:
        try:
            message = json.loads(raw_message)
        except json.JSONDecodeError:
            await websocket.send_text(
                to_json(make_event(EVENT_ERROR, message="Invalid JSON command.", recoverable=True))
            )
            return

        command_type = str(message.get("type") or "")
        if command_type == COMMAND_TEXT:
            text = str(message.get("text") or "").strip()
            language = message.get("language")
            if not text:
                return
            response = await asyncio.to_thread(self._route_text_command, text, language)
            self.broadcast(make_event(EVENT_RESPONSE, text=str(response or ""), language=language))
            return

        if command_type == COMMAND_PIN_ATTEMPT:
            pin = str(message.get("pin") or "")
            result = await asyncio.to_thread(self._apply_pin_attempt, pin)
            self.broadcast(make_event(EVENT_RESPONSE, text=result["message"], language=result["language"]))
            self.broadcast(
                make_event(
                    EVENT_PIN_RESULT,
                    status=result["status"],
                    message=result["message"],
                    attempts_remaining=result["attempts_remaining"],
                )
            )
            return

        if command_type == COMMAND_MUTE_TOGGLE:
            self.muted = bool(message.get("muted"))
            logger.info("UI bridge mute toggled: %s", self.muted)
            try:
                from audio.tts import speech_engine

                speech_engine.set_enabled(not self.muted)
                if self.muted:
                    speech_engine.interrupt()
            except Exception:
                logger.debug("Failed to apply mute to speech engine", exc_info=True)
            return

        if command_type == COMMAND_CONFIG_REQUEST:
            await websocket.send_text(to_json(self._config_event()))
            return

        if command_type == COMMAND_HEALTH_REQUEST:
            self.broadcast(self._health_event())
            return

        if command_type == COMMAND_SETTING_UPDATE:
            key = str(message.get("key") or "")
            value = message.get("value")
            logger.info("UI bridge setting_update received: key=%s", key)
            await asyncio.to_thread(self._apply_setting_update, key, value)
            await websocket.send_text(to_json(self._config_event()))
            return

        if command_type == COMMAND_FEATURE_FLAG:
            flag = str(message.get("flag") or "")
            enabled = bool(message.get("enabled"))
            if flag in config.FEATURE_FLAGS:
                config.FEATURE_FLAGS[flag] = enabled
                logger.info("UI bridge feature_flag applied: flag=%s enabled=%s", flag, enabled)
                # Enabling app discovery kicks off an immediate catalog rescan so
                # newly installed apps become usable right away; otherwise the flag
                # only gates the *next* scan and toggling it on looks like a no-op.
                if flag == "AUTO_APP_DISCOVERY_ENABLED" and enabled:
                    asyncio.create_task(self._rescan_app_catalog())
            else:
                logger.info("UI bridge feature_flag ignored (unknown flag): %s", flag)
            await websocket.send_text(to_json(self._config_event()))
            return

        logger.info("UI bridge ignored unknown command type: %s", command_type)

    def _apply_setting_update(self, key: str, value) -> None:
        try:
            if key == "JARVIS_STT_LANGUAGE_HINT":
                from audio.stt import set_runtime_stt_settings

                set_runtime_stt_settings(language_hint=value)
            elif key == "JARVIS_LLM_MODEL":
                from llm.ollama_client import set_runtime_model
                from core.hardware_detect import recommend_model_tier
                from core.config import LLM_OLLAMA_BASE_URL, LLM_OLLAMA_NUM_CTX, LLM_LIGHTWEIGHT_NUM_CTX

                model_name = str(value or "").strip()
                if not model_name or model_name.lower() == "auto":
                    tier = recommend_model_tier(str(LLM_OLLAMA_BASE_URL or "http://localhost:11434").rstrip("/"))
                    set_runtime_model(
                        tier.get("model"),
                        num_ctx=tier.get("num_ctx"),
                        lightweight_num_ctx=tier.get("lightweight_num_ctx"),
                        tier=tier.get("tier"),
                    )
                else:
                    set_runtime_model(
                        model_name,
                        num_ctx=LLM_OLLAMA_NUM_CTX,
                        lightweight_num_ctx=LLM_LIGHTWEIGHT_NUM_CTX,
                        tier="medium",
                    )
            elif key == "JARVIS_PERSONA":
                from core.persona import persona_manager

                persona_manager.set_profile(str(value or ""))
            else:
                logger.info("UI bridge setting_update ignored (unknown key): %s", key)
        except Exception:
            logger.exception("UI bridge setting_update failed for key=%s", key)

    async def _rescan_app_catalog(self) -> None:
        """Rescan installed apps in the background when Auto app discovery is
        switched on, so newly installed apps are usable immediately. Runs off the
        event loop; only a failure is surfaced to the UI (as a notification)."""
        try:
            from os_control.app_ops import refresh_app_catalog_result

            result = await asyncio.to_thread(refresh_app_catalog_result, True)
            if isinstance(result, dict) and not result.get("success", False):
                self.broadcast(
                    make_event(
                        EVENT_ERROR,
                        message=result.get("user_message") or "Auto app discovery rescan failed.",
                        recoverable=True,
                    )
                )
            else:
                logger.info("UI bridge app-catalog rescan complete.")
                message = "App list refreshed."
                if isinstance(result, dict) and result.get("user_message"):
                    message = str(result["user_message"])
                self.broadcast(make_event(EVENT_NOTIFY, message=message, tone="success"))
        except Exception:
            logger.exception("UI bridge app-catalog rescan failed")
            self.broadcast(
                make_event(EVENT_ERROR, message="Auto app discovery rescan failed.", recoverable=True)
            )

    def _route_text_command(self, text: str, language) -> str:
        try:
            from core.command_router import route_command

            return route_command(text, detected_language=language)
        except Exception:
            logger.exception("UI bridge text_command routing failed")
            return "I could not process that command."

    def _apply_pin_attempt(self, pin: str) -> dict:
        """Verify a UI-submitted PIN. A thin wrapper around the same
        verify/execute call the voice path uses (command_router.py's
        OS_PIN_CONFIRM dispatch) — this call is the single source of truth for
        attempt-counting, so it must run exactly once per submission.
        """
        from core.command_router import _execute_confirmed_payload
        from core.response_templates import render_template
        from core.session_memory import session_memory
        from os_control.confirmation import confirmation_manager

        language = session_memory.get_preferred_language() or "en"
        try:
            status, message, payload = confirmation_manager.verify_pin_and_execute(pin)
        except Exception:
            logger.exception("UI bridge pin_attempt verification failed")
            return {
                "status": "wrong",
                "message": "Sorry, I had an internal error.",
                "language": language,
                "attempts_remaining": confirmation_manager.pin_attempts_remaining(),
            }

        if status == "executed":
            try:
                _success, response_text, _meta = _execute_confirmed_payload(payload)
            except Exception:
                logger.exception("UI bridge pin_attempt execution failed")
                response_text = "Sorry, I had an internal error."
            session_memory.clear_pending_confirmation_token()
        elif status == "wrong":
            response_text = render_template("pin_wrong", language)
        elif status == "locked":
            session_memory.clear_pending_confirmation_token()
            response_text = render_template("pin_locked", language, message=message)
        else:
            session_memory.clear_pending_confirmation_token()
            response_text = render_template("missing_pending_confirmation", language)

        try:
            from audio.tts import speech_engine

            speech_engine.speak_async(response_text, language=language)
        except Exception:
            logger.debug("Failed to speak pin_attempt response", exc_info=True)

        return {
            "status": status,
            "message": response_text,
            "language": language,
            "attempts_remaining": confirmation_manager.pin_attempts_remaining(),
        }

    def _register_client(self, websocket: WebSocket) -> None:
        with self.lock:
            self.clients.add(websocket)
        logger.info("UI bridge client connected; clients=%s", len(self.clients))

    def _unregister_client(self, websocket: WebSocket) -> None:
        with self.lock:
            self.clients.discard(websocket)
        logger.info("UI bridge client disconnected; clients=%s", len(self.clients))

    def _on_state_changed(self, old_state, new_state) -> None:
        try:
            self.broadcast(
                make_event(
                    EVENT_STATE_CHANGED,
                    state=getattr(new_state, "value", str(new_state)),
                    previous=getattr(old_state, "value", str(old_state)),
                )
            )
        except Exception:
            logger.debug("Bridge state listener failed", exc_info=True)

    def _config_event(self) -> dict:
        try:
            from llm.ollama_client import get_runtime_model

            model = get_runtime_model(default=getattr(config, "LLM_MODEL", ""))
        except Exception:
            model = getattr(config, "LLM_MODEL", "")

        try:
            from core.persona import persona_manager

            persona = persona_manager.get_profile() or getattr(config, "PERSONA_DEFAULT", "")
        except Exception:
            persona = getattr(config, "PERSONA_DEFAULT", "")

        values = {
            "model": model,
            "model_tier": "auto" if bool(getattr(config, "LLM_AUTO_SELECT_MODEL", False)) else "configured",
            "wake_mode": getattr(config, "WAKE_WORD_MODE", ""),
            "feature_flags": dict(getattr(config, "FEATURE_FLAGS", {}) or {}),
            "stt_backend": getattr(config, "STT_BACKEND", ""),
            "tts_backend": getattr(config, "TTS_DEFAULT_BACKEND", ""),
            "persona": persona,
        }
        return make_event(EVENT_CONFIG, values=values)

    def _health_event(self) -> dict:
        # TODO: Replace this cheap bridge-local status with collect_diagnostics()
        # once the UI can tolerate slower health probes off the socket hot path.
        checks = [
            {
                "name": "ui_bridge",
                "status": "ok" if self.running else "degraded",
                "detail": f"clients={len(self.clients)} muted={self.muted}",
            }
        ]
        return make_event(EVENT_HEALTH, checks=checks)

    def broadcast(self, event: dict) -> None:
        if not self.running:
            return
        loop = self.loop
        if loop is None or loop.is_closed():
            return
        try:
            asyncio.run_coroutine_threadsafe(self._broadcast(event), loop)
        except Exception:
            logger.debug("Failed to schedule bridge broadcast", exc_info=True)

    async def _broadcast(self, event: dict) -> None:
        payload = to_json(event)
        with self.lock:
            clients = tuple(self.clients)
        stale_clients = []
        for websocket in clients:
            try:
                await websocket.send_text(payload)
            except Exception:
                logger.debug("Failed to send bridge event to client", exc_info=True)
                stale_clients.append(websocket)
        if stale_clients:
            with self.lock:
                for websocket in stale_clients:
                    self.clients.discard(websocket)


bridge = JarvisBridge()


def start_bridge() -> None:
    bridge.start()


def stop_bridge() -> None:
    bridge.stop()


def broadcast_event(event: dict) -> None:
    bridge.broadcast(event)
