from __future__ import annotations

import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from openpyxl import Workbook

from app.modules.claims.generate_claims_from_registry import (
    generate_claims_zip_result,
    read_registry,
)


def _build_registry(path: Path, rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Реестр"
    worksheet.append(["Лицевой счет", "ФИО", "Адрес", "Период задолженности", "Сумма долга"])
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


def _build_template(path: Path) -> None:
    document = Document()
    document.add_paragraph("Должник: {{ debtor_name }}")
    document.add_paragraph("Адрес: {{ address }}")
    document.add_paragraph("Сумма: {{ debt_amount }} ({{ debt_amount_words }})")
    document.add_paragraph("Объект: {{ object_address }}")
    document.save(path)


class ClaimsGenerationTests(unittest.TestCase):
    def test_read_registry_parses_positive_rows_only(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            registry_path = Path(temp_dir) / "registry.xlsx"
            _build_registry(
                registry_path,
                [
                    ["12970001", "Иванов Иван", "Москва, машиноместо № 0001", "01.01.2026-01.02.2026", 1000],
                    ["12970002", "Петров Петр", "Москва, машиноместо № 0002", "01.01.2026-01.02.2026", 0],
                ],
            )

            rows = read_registry(registry_path)

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].account_number, "12970001")
        self.assertEqual(rows[0].debt_amount, Decimal("1000.00"))

    def test_generate_claims_zip_result_creates_one_docx_per_registry_row(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            registry_path = temp / "registry.xlsx"
            template_path = temp / "template.docx"
            output_zip_path = temp / "claims.zip"
            _build_registry(
                registry_path,
                [
                    ["12970001", "Иванов Иван", "Москва, машиноместо № 0001", "01.01.2026-01.02.2026", 1000],
                    ["12970002", "Петров Петр", "Москва, машиноместо № 0002", "01.01.2026-01.02.2026", 2500.50],
                ],
            )
            _build_template(template_path)

            result = generate_claims_zip_result(
                registry_path=str(registry_path),
                template_path=str(template_path),
                output_zip_path=str(output_zip_path),
                claim_date="01.03.2026",
                payment_deadline="31.03.2026",
            )

            with ZipFile(output_zip_path) as archive:
                names = archive.namelist()

        self.assertEqual(result.documents_count, 2)
        self.assertEqual(result.total_amount, Decimal("3500.50"))
        self.assertEqual(len(names), 2)
        self.assertTrue(all(name.endswith(".docx") for name in names))

    def test_generate_claims_zip_result_rejects_empty_registry(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            registry_path = temp / "registry.xlsx"
            template_path = temp / "template.docx"
            _build_registry(registry_path, [])
            _build_template(template_path)

            with self.assertRaisesRegex(ValueError, "нет строк"):
                generate_claims_zip_result(
                    registry_path=str(registry_path),
                    template_path=str(template_path),
                    output_zip_path=str(temp / "claims.zip"),
                    claim_date="01.03.2026",
                    payment_deadline="31.03.2026",
                )


if __name__ == "__main__":
    unittest.main()
