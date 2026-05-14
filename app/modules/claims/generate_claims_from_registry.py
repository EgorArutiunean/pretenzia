from __future__ import annotations

import argparse
import re
import shutil
import uuid
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook

from app.shared.file_utils import safe_filename
from app.shared.money_to_words import money_to_words
from app.shared.zip_utils import create_zip


REQUIRED_COLUMNS = [
    "Лицевой счет",
    "ФИО",
    "Адрес",
    "Период задолженности",
    "Сумма долга",
]

DEFAULT_CONTEXT = {
    "company_name": "ООО «ПАРКИНГ ЛАЙН»",
    "company_ogrn": "1207700184977",
    "company_inn": "7727444957",
    "company_kpp": "772701001",
    "company_post_address": "117042, город Москва, Чечёрский проезд, д. 120, помещ. 1/1",
    "company_email": "_______",
    "director_position": "Генеральный директор",
    "director_name": "Голубев И. А.",
}


@dataclass(frozen=True)
class ClaimRow:
    account_number: str
    debtor_name: str
    address: str
    debt_period: str
    debt_amount: Decimal

    @property
    def parking_place_number(self) -> str:
        digits = re.sub(r"\D+", "", self.account_number)
        return digits[-4:] if len(digits) >= 4 else digits


def parse_money(value: Any) -> Decimal:
    if value is None:
        return Decimal("0.00")
    if isinstance(value, Decimal):
        return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if isinstance(value, (int, float)):
        return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    text = str(value).strip().replace("\u00a0", " ")
    text = re.sub(r"[^0-9,\.\- ]+", "", text).replace(" ", "")
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    else:
        text = text.replace(",", ".")

    try:
        return Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"Не удалось распознать сумму долга: {value!r}") from exc


def format_money(value: Decimal) -> str:
    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    integer, frac = f"{value:.2f}".split(".")
    return f"{int(integer):,}".replace(",", " ") + f",{frac}"


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def read_registry(registry_path: str | Path, sheet_name: str | None = None) -> list[ClaimRow]:
    path = Path(registry_path)
    if not path.exists():
        raise FileNotFoundError(f"Excel-реестр не найден: {path}")

    wb = load_workbook(path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    header_row: int | None = None
    columns: dict[str, int] = {}
    required_by_norm = {normalize_header(col): col for col in REQUIRED_COLUMNS}

    for row_idx in range(1, min(ws.max_row, 30) + 1):
        found: dict[str, int] = {}
        for col_idx in range(1, ws.max_column + 1):
            norm = normalize_header(ws.cell(row_idx, col_idx).value)
            if norm in required_by_norm:
                found[required_by_norm[norm]] = col_idx
        if len(found) == len(REQUIRED_COLUMNS):
            header_row = row_idx
            columns = found
            break

    if header_row is None:
        visible_headers = [
            str(ws.cell(1, col_idx).value)
            for col_idx in range(1, ws.max_column + 1)
            if ws.cell(1, col_idx).value not in (None, "")
        ]
        raise ValueError(
            "В Excel-реестре не найдены обязательные колонки: "
            + ", ".join(REQUIRED_COLUMNS)
            + ". Найденные заголовки первой строки: "
            + (", ".join(visible_headers) if visible_headers else "нет")
        )

    rows: list[ClaimRow] = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        raw_values = {
            column: ws.cell(row_idx, col_idx).value
            for column, col_idx in columns.items()
        }
        if all(value in (None, "") for value in raw_values.values()):
            continue

        amount = parse_money(raw_values["Сумма долга"])
        if amount <= 0:
            continue

        rows.append(
            ClaimRow(
                account_number=str(raw_values["Лицевой счет"] or "").strip(),
                debtor_name=str(raw_values["ФИО"] or "").strip(),
                address=str(raw_values["Адрес"] or "").strip(),
                debt_period=str(raw_values["Период задолженности"] or "").strip().replace("-", " - "),
                debt_amount=amount,
            )
        )

    return rows


def _replace_text_in_paragraph(paragraph, replacements: dict[str, str]) -> None:
    if not paragraph.runs:
        return

    original = "".join(run.text for run in paragraph.runs)
    replaced = original
    for key, value in replacements.items():
        replaced = replaced.replace("{{ " + key + " }}", value)
        replaced = replaced.replace("{{" + key + "}}", value)

    if replaced != original:
        paragraph.runs[0].text = replaced
        for run in paragraph.runs[1:]:
            run.text = ""


def _replace_in_doc(doc: Document, replacements: dict[str, str]) -> None:
    for paragraph in doc.paragraphs:
        _replace_text_in_paragraph(paragraph, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_text_in_paragraph(paragraph, replacements)

    for section in doc.sections:
        for header_footer in (section.header, section.footer):
            for paragraph in header_footer.paragraphs:
                _replace_text_in_paragraph(paragraph, replacements)
            for table in header_footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            _replace_text_in_paragraph(paragraph, replacements)


def build_replacements(row: ClaimRow, claim_date: str, payment_deadline: str) -> dict[str, str]:
    context = dict(DEFAULT_CONTEXT)
    context.update(
        {
            "account_number": row.account_number,
            "debtor_name": row.debtor_name,
            "address": row.address,
            "debt_period": row.debt_period,
            "debt_amount": format_money(row.debt_amount),
            "debt_amount_words": money_to_words(row.debt_amount),
            "parking_place_number": row.parking_place_number,
            "claim_date": claim_date,
            "payment_deadline": payment_deadline,
        }
    )
    return {key: str(value) for key, value in context.items()}


def _project_root() -> Path:
    return Path(__file__).resolve().parents[3]


def generate_claims_zip(
    registry_path: str,
    template_path: str,
    output_zip_path: str,
    claim_date: str,
    payment_deadline: str,
) -> str:
    registry = Path(registry_path)
    template = Path(template_path)
    output_zip = Path(output_zip_path)

    if not template.exists():
        raise FileNotFoundError(f"Word-шаблон претензии не найден: {template}")

    rows = read_registry(registry)
    temp_root = _project_root() / "storage" / "temp"
    temp_root.mkdir(parents=True, exist_ok=True)
    temp_dir = temp_root / f"claims_{uuid.uuid4().hex}"
    temp_dir.mkdir(parents=True, exist_ok=False)

    try:
        generated_files: list[Path] = []
        used_names: set[str] = set()

        for index, row in enumerate(rows, start=1):
            doc = Document(str(template))
            _replace_in_doc(doc, build_replacements(row, claim_date, payment_deadline))

            base_name = safe_filename(f"{index:03d}_{row.account_number}_{row.debtor_name}")
            file_name = base_name + ".docx"
            counter = 2
            while file_name.lower() in used_names:
                file_name = f"{base_name}_{counter}.docx"
                counter += 1
            used_names.add(file_name.lower())

            file_path = temp_dir / file_name
            doc.save(file_path)
            generated_files.append(file_path)

        create_zip(generated_files, output_zip)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    return str(output_zip)


def parse_args() -> argparse.Namespace:
    today = date.today()
    default_deadline = today + timedelta(days=30)

    parser = argparse.ArgumentParser(description="Генерация Word-претензий из Excel-реестра и упаковка в ZIP")
    parser.add_argument("registry", help="Готовый Excel-реестр")
    parser.add_argument("--out", default="storage/output/claims.zip", help="Путь к итоговому ZIP-архиву")
    parser.add_argument("--template", default="app/modules/claims/claim_template.docx", help="Word-шаблон претензии")
    parser.add_argument("--claim-date", default=today.strftime("%d.%m.%Y"), help="Дата претензии")
    parser.add_argument("--payment-deadline", default=default_deadline.strftime("%d.%m.%Y"), help="Срок оплаты")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    zip_path = generate_claims_zip(
        registry_path=args.registry,
        template_path=args.template,
        output_zip_path=args.out,
        claim_date=args.claim_date,
        payment_deadline=args.payment_deadline,
    )
    count = len(read_registry(args.registry))
    print(f"Создано претензий: {count}")
    print(f"Архив: {zip_path}")


if __name__ == "__main__":
    main()
