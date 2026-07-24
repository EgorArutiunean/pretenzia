from __future__ import annotations

import os
import shutil
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Callable, TypeVar


ValidationResult = TypeVar("ValidationResult")
_REFERENCE_UPDATE_LOCK = threading.Lock()


def install_validated_reference_file(
    upload_path: Path,
    target_path: Path,
    validator: Callable[[Path], ValidationResult],
    *,
    keep_backups: int = 5,
) -> tuple[ValidationResult, Path | None]:
    """Validate and atomically replace a reference file, retaining recent backups."""
    report = validator(upload_path)
    target_path.parent.mkdir(parents=True, exist_ok=True)

    with _REFERENCE_UPDATE_LOCK:
        backup_path: Path | None = None
        backup_dir = target_path.parent / "backups"
        if target_path.exists():
            if not target_path.is_file():
                raise ValueError(f"Путь рабочего справочника не является файлом: {target_path}")
            backup_dir.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            backup_path = backup_dir / f"{target_path.stem}_{timestamp}{target_path.suffix}"
            shutil.copy2(target_path, backup_path)

        staged_path = target_path.parent / f".{target_path.name}.{uuid.uuid4().hex}.tmp"
        try:
            shutil.copy2(upload_path, staged_path)
            os.replace(staged_path, target_path)
        finally:
            staged_path.unlink(missing_ok=True)

        upload_path.unlink(missing_ok=True)
        _prune_backups(backup_dir, target_path.stem, target_path.suffix, keep_backups)

    return report, backup_path


def _prune_backups(backup_dir: Path, stem: str, suffix: str, keep_backups: int) -> None:
    if keep_backups < 0 or not backup_dir.exists():
        return
    backups = sorted(
        backup_dir.glob(f"{stem}_*{suffix}"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    for old_backup in backups[keep_backups:]:
        old_backup.unlink(missing_ok=True)
