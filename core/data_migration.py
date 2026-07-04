"""One-time migration of legacy root-level runtime artifacts into data/.

Older Jarvis installs wrote logs, the memory DB, the search index, the state
DB, and the KB/Chroma directories directly into the project root. Phase 1 of
the memory/file-organization plan consolidates all of that under a single
``data/`` tree (see core/config.py DATA_DIR and friends). This module moves
any artifact still sitting at its old root location into the new data/
location, so existing installs upgrade in place without losing history.
"""

from __future__ import annotations

import os
import shutil

from core.config import (
    ACTION_LOG_FILE,
    KB_STORAGE_DIR,
    LOG_FILE,
    MEMORY_DB_FILE,
    MEMORY_FILE,
    PROJECT_ROOT,
    SEARCH_INDEX_DB_FILE,
    STATE_DB_FILE,
    VECTOR_MEMORY_DIR,
)
from core.logger import get_logger

logger = get_logger("startup")


def _move_file(old_path: str, new_path: str) -> None:
    if not os.path.isfile(old_path) or os.path.isfile(new_path):
        return
    os.makedirs(os.path.dirname(new_path), exist_ok=True)
    shutil.move(old_path, new_path)
    logger.info("Migrated legacy file: %s -> %s", old_path, new_path)


def _move_sqlite(old_path: str, new_path: str) -> None:
    _move_file(old_path, new_path)
    for suffix in ("-wal", "-shm"):
        _move_file(old_path + suffix, new_path + suffix)


def _move_dir(old_path: str, new_path: str) -> None:
    if not os.path.isdir(old_path) or os.path.exists(new_path):
        return
    os.makedirs(os.path.dirname(new_path), exist_ok=True)
    shutil.move(old_path, new_path)
    logger.info("Migrated legacy directory: %s -> %s", old_path, new_path)


def migrate_legacy_paths() -> None:
    """Move any runtime artifact still at its old root location into data/.

    Idempotent: safe to call on every startup. Never overwrites an existing
    file/directory at the new location.
    """
    root = str(PROJECT_ROOT)

    _move_file(os.path.join(root, "jarvis.log"), LOG_FILE)
    for i in range(1, 10):
        _move_file(
            os.path.join(root, f"jarvis.log.{i}"),
            f"{LOG_FILE}.{i}",
        )
    _move_file(os.path.join(root, "jarvis_actions.log"), ACTION_LOG_FILE)

    _move_sqlite(os.path.join(root, "jarvis_memory.db"), MEMORY_DB_FILE)
    _move_file(os.path.join(root, "jarvis_memory.json"), MEMORY_FILE)

    _move_sqlite(os.path.join(root, "jarvis_index.db"), SEARCH_INDEX_DB_FILE)
    _move_sqlite(os.path.join(root, "jarvis_state.db"), STATE_DB_FILE)

    _move_dir(os.path.join(root, ".jarvis_kb"), KB_STORAGE_DIR)
    _move_dir(os.path.join(root, "data", "chroma_memory"), VECTOR_MEMORY_DIR)
