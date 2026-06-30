from __future__ import annotations

import argparse
import shutil
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

from app.config import PROJECT_ROOT


@dataclass(frozen=True)
class CleanupResult:
    deleted: list[Path]
    kept: list[Path]
    cutoff: datetime
    dry_run: bool


def _iter_run_dirs(runs_root: Path) -> list[Path]:
    if not runs_root.exists():
        return []
    return [
        path
        for path in runs_root.glob("user_*/*")
        if path.is_dir()
    ]


def cleanup_runs(
    runs_root: str | Path | None = None,
    *,
    older_than_days: int = 14,
    dry_run: bool = True,
    now: datetime | None = None,
) -> CleanupResult:
    root = Path(runs_root) if runs_root is not None else PROJECT_ROOT / "storage" / "runs"
    current_time = now or datetime.now()
    cutoff = current_time - timedelta(days=older_than_days)
    deleted: list[Path] = []
    kept: list[Path] = []

    for run_dir in _iter_run_dirs(root):
        mtime = datetime.fromtimestamp(run_dir.stat().st_mtime)
        if mtime >= cutoff:
            kept.append(run_dir)
            continue

        deleted.append(run_dir)
        if not dry_run:
            shutil.rmtree(run_dir, ignore_errors=True)

    return CleanupResult(deleted=deleted, kept=kept, cutoff=cutoff, dry_run=dry_run)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Удаление старых пользовательских запусков из storage/runs")
    parser.add_argument("--runs-root", default=str(PROJECT_ROOT / "storage" / "runs"), help="Корень storage/runs")
    parser.add_argument("--older-than-days", type=int, default=14, help="Удалять запуски старше N дней")
    parser.add_argument("--apply", action="store_true", help="Фактически удалить. Без флага только dry-run.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    result = cleanup_runs(
        runs_root=args.runs_root,
        older_than_days=args.older_than_days,
        dry_run=not args.apply,
    )
    mode = "dry-run" if result.dry_run else "apply"
    print(f"Cleanup mode: {mode}")
    print(f"Cutoff: {result.cutoff:%Y-%m-%d %H:%M:%S}")
    print(f"Deleted candidates: {len(result.deleted)}")
    print(f"Kept: {len(result.kept)}")
    for path in result.deleted:
        print(path)


if __name__ == "__main__":
    main()
