from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from app.shared.reference_files import install_validated_reference_file


class ReferenceFilesTests(unittest.TestCase):
    def test_install_validated_reference_file_keeps_previous_version(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            upload = temp / "upload.xlsx"
            target = temp / "references" / "base_data.xlsx"
            target.parent.mkdir(parents=True)
            target.write_bytes(b"old")
            upload.write_bytes(b"new")

            report, backup = install_validated_reference_file(
                upload,
                target,
                lambda path: path.stat().st_size,
            )

            self.assertEqual(report, 3)
            self.assertEqual(target.read_bytes(), b"new")
            self.assertFalse(upload.exists())
            self.assertIsNotNone(backup)
            self.assertEqual(backup.read_bytes(), b"old")

    def test_invalid_upload_does_not_replace_current_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            upload = temp / "upload.xlsx"
            target = temp / "base_data.xlsx"
            upload.write_bytes(b"invalid")
            target.write_bytes(b"current")

            with self.assertRaisesRegex(ValueError, "invalid"):
                install_validated_reference_file(
                    upload,
                    target,
                    lambda path: (_ for _ in ()).throw(ValueError("invalid")),
                )

            self.assertEqual(target.read_bytes(), b"current")
            self.assertTrue(upload.exists())


if __name__ == "__main__":
    unittest.main()
