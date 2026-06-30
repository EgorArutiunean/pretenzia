from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from app.config import load_settings


class ConfigTests(unittest.TestCase):
    def test_require_bot_rejects_empty_admin_ids(self) -> None:
        env = {
            "BOT_TOKEN": "123:test",
            "ADMIN_IDS": "",
            "MAX_UPLOAD_MB": "20",
        }
        with patch.dict(os.environ, env, clear=True):
            with self.assertRaisesRegex(RuntimeError, "ADMIN_IDS"):
                load_settings(require_bot=True)

    def test_admin_ids_are_parsed(self) -> None:
        env = {
            "BOT_TOKEN": "123:test",
            "ADMIN_IDS": "111, 222",
            "MAX_UPLOAD_MB": "25",
        }
        with patch.dict(os.environ, env, clear=True):
            settings = load_settings(require_bot=True)

        self.assertEqual(settings.admin_ids, {111, 222})
        self.assertEqual(settings.max_upload_mb, 25)

    def test_require_bot_rejects_object_addresses_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            env = {
                "BOT_TOKEN": "123:test",
                "ADMIN_IDS": "111",
                "OBJECT_ADDRESSES_PATH": str(Path(temp_dir)),
                "MAX_UPLOAD_MB": "20",
            }
            with patch.dict(os.environ, env, clear=True):
                with self.assertRaisesRegex(RuntimeError, "OBJECT_ADDRESSES_PATH must point"):
                    load_settings(require_bot=True)


if __name__ == "__main__":
    unittest.main()
