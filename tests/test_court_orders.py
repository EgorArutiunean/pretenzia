from __future__ import annotations

import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from openpyxl import Workbook

from app.modules.court_orders.generate_court_orders_from_registry import (
    calculate_state_duty,
    generate_court_orders_zip_result,
    load_static_data,
)


def _build_registry(path: Path, rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Реестр"
    worksheet.append(["Лицевой счет", "ФИО", "Адрес", "Период задолженности", "Сумма долга"])
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


def _build_static_data(path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["Номер обьекта ", "Адрес", "Компания", "Реквезиты ", "Протокол", "Ставка", "Подсудность"])
    worksheet.append(
        [
            1297,
            "г. Москва, Тестовая улица, д. 1",
            "ТЕСТ ООО",
            (
                'Сокращенное наименование ООО "ТЕСТ"\n'
                "ИНН 7700000000\n"
                "КПП 770001001\n"
                "ОГРН 1200000000000\n"
                "Юридический адрес 127000, г. Москва, Тестовая ул., д. 1 "
                "Генеральный директор Иванов Иван Иванович"
            ),
            "протокол от 01 января 2026 года",
            950,
            "Мировому судье судебного участка № 1",
        ]
    )
    workbook.save(path)


def _build_template(path: Path) -> None:
    document = Document()
    document.add_paragraph("{{ court_name }}")
    document.add_paragraph("Взыскатель: {{ claimant_name }}")
    document.add_paragraph("Должник: {{ debtor_name }}")
    document.add_paragraph("Машино-место № {{ parking_place_number }}")
    document.add_paragraph("Задолженность: {{ debt_amount }}")
    document.add_paragraph("Госпошлина: {{ state_duty }}")
    document.save(path)


class CourtOrdersGenerationTests(unittest.TestCase):
    def test_load_static_data_accepts_current_workbook_headers(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            static_path = Path(temp_dir) / "static.xlsx"
            _build_static_data(static_path)

            data = load_static_data(static_path)

        self.assertIn("1297", data)
        self.assertEqual(data["1297"].claimant_name, 'ООО "ТЕСТ"')
        self.assertEqual(data["1297"].inn, "7700000000")

    def test_calculate_state_duty_for_court_order_uses_half_of_claim_fee(self) -> None:
        self.assertEqual(calculate_state_duty(Decimal("50225.00")), Decimal("2000.00"))
        self.assertEqual(calculate_state_duty(Decimal("150000.00")), Decimal("2750.00"))

    def test_generate_court_orders_zip_result_creates_docx_for_known_objects(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            registry_path = temp / "registry.xlsx"
            static_path = temp / "static.xlsx"
            template_path = temp / "template.docx"
            output_zip_path = temp / "court_orders.zip"
            _build_registry(
                registry_path,
                [
                    ["12970001", "Иванов Иван", "Москва, машиноместо № 0001", "01.01.2026 - 31.01.2026", 1000],
                    ["99990001", "Петров Петр", "Москва, машиноместо № 0001", "01.01.2026 - 31.01.2026", 2000],
                ],
            )
            _build_static_data(static_path)
            _build_template(template_path)

            result = generate_court_orders_zip_result(
                registry_path=str(registry_path),
                template_path=str(template_path),
                static_data_path=str(static_path),
                output_zip_path=str(output_zip_path),
                application_date="01.03.2026",
            )

            with ZipFile(output_zip_path) as archive:
                names = archive.namelist()
                archive.extract(names[0], temp)

            generated_doc = Document(str(temp / names[0]))
            generated_text = "\n".join(paragraph.text for paragraph in generated_doc.paragraphs)

        self.assertEqual(result.documents_count, 1)
        self.assertEqual(result.skipped_count, 1)
        self.assertEqual(result.total_debt_amount, Decimal("1000.00"))
        self.assertEqual(len(names), 1)
        self.assertIn("Иванов Иван", generated_text)
        self.assertIn("ООО \"ТЕСТ\"", generated_text)


if __name__ == "__main__":
    unittest.main()
