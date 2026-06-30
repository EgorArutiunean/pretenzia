from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from docx import Document

from app.config import PROJECT_ROOT
from app.modules.claims.generate_claims_from_registry import generate_claims_zip_result


class ReleaseSmokeTests(unittest.TestCase):
    def test_registry_template_generates_claims_with_real_template(self) -> None:
        registry_path = PROJECT_ROOT / "registry_template.xlsx"
        template_path = PROJECT_ROOT / "app" / "modules" / "claims" / "claim_template.docx"

        with tempfile.TemporaryDirectory() as temp_dir:
            output_zip_path = Path(temp_dir) / "claims.zip"
            result = generate_claims_zip_result(
                registry_path=str(registry_path),
                template_path=str(template_path),
                output_zip_path=str(output_zip_path),
                claim_date="25.04.2026",
                payment_deadline="25.05.2026",
            )

            with ZipFile(output_zip_path) as archive:
                names = archive.namelist()
                archive.extract(names[0], temp_dir)

            generated_doc = Document(str(Path(temp_dir) / names[0]))
            generated_text = "\n".join(paragraph.text for paragraph in generated_doc.paragraphs)

        self.assertEqual(result.documents_count, 1)
        self.assertEqual(len(names), 1)
        self.assertNotIn("{{", generated_text)
        self.assertIn("Тестовый Должник", generated_text)


if __name__ == "__main__":
    unittest.main()
