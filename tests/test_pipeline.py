from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from app.pipeline import run_excel_to_registry


class PipelineTests(unittest.TestCase):
    def test_run_excel_to_registry_rejects_missing_object_addresses_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            run_dir = Path(temp_dir) / "run"
            input_path = Path(temp_dir) / "raw.xlsx"
            input_path.write_bytes(b"not used")
            env = {
                "OBJECT_ADDRESSES_PATH": str(Path(temp_dir) / "missing.xlsx"),
                "MAX_UPLOAD_MB": "20",
            }

            with patch.dict(os.environ, env, clear=True):
                with self.assertRaisesRegex(FileNotFoundError, "OBJECT_ADDRESSES_PATH does not exist"):
                    run_excel_to_registry(str(input_path), str(run_dir))

    def test_run_excel_to_registry_rejects_object_addresses_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            run_dir = Path(temp_dir) / "run"
            input_path = Path(temp_dir) / "raw.xlsx"
            input_path.write_bytes(b"not used")
            env = {
                "OBJECT_ADDRESSES_PATH": temp_dir,
                "MAX_UPLOAD_MB": "20",
            }

            with patch.dict(os.environ, env, clear=True):
                with self.assertRaisesRegex(ValueError, "OBJECT_ADDRESSES_PATH must point"):
                    run_excel_to_registry(str(input_path), str(run_dir))


if __name__ == "__main__":
    unittest.main()
