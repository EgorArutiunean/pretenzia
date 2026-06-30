from __future__ import annotations

import os
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path

from app.maintenance.cleanup_storage import cleanup_runs


class CleanupStorageTests(unittest.TestCase):
    def test_cleanup_runs_deletes_only_old_run_dirs_when_applied(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "runs"
            old_run = root / "user_1" / "2026-01-01_10-00-00"
            new_run = root / "user_1" / "2026-01-10_10-00-00"
            old_run.mkdir(parents=True)
            new_run.mkdir(parents=True)

            now = datetime(2026, 1, 20, 12, 0, 0)
            old_mtime = (now - timedelta(days=30)).timestamp()
            new_mtime = (now - timedelta(days=1)).timestamp()
            os.utime(old_run, (old_mtime, old_mtime))
            os.utime(new_run, (new_mtime, new_mtime))

            result = cleanup_runs(root, older_than_days=14, dry_run=False, now=now)

            self.assertEqual(result.deleted, [old_run])
            self.assertFalse(old_run.exists())
            self.assertTrue(new_run.exists())

    def test_cleanup_runs_dry_run_keeps_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "runs"
            old_run = root / "user_1" / "2026-01-01_10-00-00"
            old_run.mkdir(parents=True)
            now = datetime(2026, 1, 20, 12, 0, 0)
            old_mtime = (now - timedelta(days=30)).timestamp()
            os.utime(old_run, (old_mtime, old_mtime))

            result = cleanup_runs(root, older_than_days=14, dry_run=True, now=now)

            self.assertEqual(result.deleted, [old_run])
            self.assertTrue(old_run.exists())


if __name__ == "__main__":
    unittest.main()
