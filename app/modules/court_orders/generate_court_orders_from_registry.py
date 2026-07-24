from __future__ import annotations

import argparse
import re
import shutil
import sys
import uuid
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook

from app.modules.claims.generate_claims_from_registry import (
    ClaimRow,
    format_money,
    normalize_header as normalize_registry_header,
    parse_money,
    read_registry,
)
from app.shared.file_utils import safe_filename
from app.shared.money_to_words import money_to_words
from app.shared.zip_utils import create_zip


REQUIRED_STATIC_COLUMNS = [
    "Номер объекта",
    "Адрес",
    "Компания",
]

STATIC_COLUMN_ALIASES = {
    "номер объекта": "Номер объекта",
    "номер обьекта": "Номер объекта",
    "объект": "Номер объекта",
    "обьект": "Номер объекта",
    "код объекта": "Номер объекта",
    "код обьекта": "Номер объекта",
    "адрес": "Адрес",
    "компания": "Компания",
    "реквизиты": "Реквизиты",
    "реквезиты": "Реквизиты",
    "протокол": "Протокол",
    "ставка": "Ставка",
    "подсудность": "Подсудность",
    "суд": "Подсудность",
    "судебный участок": "Подсудность",
    "адрес суда": "Адрес суда",
}


@dataclass(frozen=True)
class CourtOrderStaticData:
    object_code: str
    object_address: str
    company: str
    requisites: str
    protocol: str
    monthly_rate: Decimal | None
    court: str
    court_address: str

    @property
    def claimant_name(self) -> str:
        match = re.search(r'ООО\s+"([^"]+)"', self.requisites)
        if match:
            return f'ООО "{match.group(1)}"'
        return self.company

    @property
    def inn(self) -> str:
        return _extract_regex(self.requisites, r"ИНН\s+(\d+)")

    @property
    def kpp(self) -> str:
        return _extract_regex(self.requisites, r"КПП\s+(\d+)")

    @property
    def ogrn(self) -> str:
        return _extract_regex(self.requisites, r"ОГРН\s+(\d+)")

    @property
    def legal_address(self) -> str:
        return _extract_between(self.requisites, "Юридический адрес", "Генеральный директор")

    @property
    def director_name(self) -> str:
        text = _extract_between(self.requisites, "Генеральный директор", "Главный бухгалтер")
        if text:
            return text
        return _extract_regex(self.requisites, r"Генеральный директор\s+(.+?)(?:\s+Расчетный счет|$)")


@dataclass(frozen=True)
class CourtOrderGenerationResult:
    zip_path: str
    documents_count: int
    skipped_count: int
    total_debt_amount: Decimal


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower()).replace("ё", "е")


def object_code_from_account(account_number: str) -> str:
    digits = re.sub(r"\D+", "", account_number)
    if len(digits) < 4:
        return ""
    return digits[:4]


def read_optional_money_by_account(
    registry_path: str | Path,
    column_names: set[str],
) -> dict[str, Decimal]:
    path = Path(registry_path)
    workbook = load_workbook(path, data_only=True)
    worksheet = workbook.active
    expected = {normalize_registry_header(name) for name in column_names}

    account_col: int | None = None
    value_col: int | None = None
    header_row: int | None = None

    for row_idx in range(1, min(worksheet.max_row, 30) + 1):
        for col_idx in range(1, worksheet.max_column + 1):
            header = normalize_registry_header(worksheet.cell(row_idx, col_idx).value)
            if header == normalize_registry_header("Лицевой счет"):
                account_col = col_idx
            if header in expected:
                value_col = col_idx
        if account_col is not None and value_col is not None:
            header_row = row_idx
            break

    if header_row is None or account_col is None or value_col is None:
        return {}

    values: dict[str, Decimal] = {}
    for row_idx in range(header_row + 1, worksheet.max_row + 1):
        account = str(worksheet.cell(row_idx, account_col).value or "").strip()
        if not account:
            continue
        raw_value = worksheet.cell(row_idx, value_col).value
        if raw_value in (None, ""):
            continue
        values[account] = parse_money(raw_value)
    return values


def _extract_regex(text: str, pattern: str) -> str:
    match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
    return " ".join(match.group(1).split()) if match else ""


def _extract_between(text: str, start: str, end: str) -> str:
    pattern = re.escape(start) + r"\s+(.+?)\s+" + re.escape(end)
    return _extract_regex(text, pattern)


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


def _read_header_map(worksheet) -> dict[str, int]:
    header_map: dict[str, int] = {}
    for col_idx in range(1, worksheet.max_column + 1):
        normalized = normalize_header(worksheet.cell(1, col_idx).value)
        if not normalized:
            continue
        canonical = STATIC_COLUMN_ALIASES.get(normalized)
        if canonical:
            header_map[canonical] = col_idx

    missing = [column for column in REQUIRED_STATIC_COLUMNS if column not in header_map]
    if missing:
        raise ValueError(
            "В БЗ для судебных заявлений не найдены обязательные колонки: "
            + ", ".join(missing)
        )
    return header_map


def load_static_data(static_data_path: str | Path) -> dict[str, CourtOrderStaticData]:
    path = Path(static_data_path)
    if not path.exists():
        raise FileNotFoundError(f"БЗ для судебных заявлений не найдена: {path}")
    if not path.is_file():
        raise ValueError(f"БЗ для судебных заявлений должна быть файлом .xlsx: {path}")

    workbook = load_workbook(path, data_only=True)
    worksheet = workbook.active
    header_map = _read_header_map(worksheet)

    by_object: dict[str, CourtOrderStaticData] = {}
    defaults_by_company: dict[str, dict[str, Any]] = {}

    for row_idx in range(2, worksheet.max_row + 1):
        values = {
            column: worksheet.cell(row_idx, col_idx).value
            for column, col_idx in header_map.items()
        }
        object_code = re.sub(r"\D+", "", str(values.get("Номер объекта") or ""))
        company = str(values.get("Компания") or "").strip()
        if not object_code or not company:
            continue

        company_defaults = defaults_by_company.setdefault(company, {})
        for column in ("Реквизиты", "Протокол", "Ставка", "Подсудность", "Адрес суда"):
            if column in values and values[column] not in (None, ""):
                company_defaults[column] = values[column]
            elif column in company_defaults:
                values[column] = company_defaults[column]

        rate = None
        if values.get("Ставка") not in (None, ""):
            rate = parse_money(values["Ставка"])

        by_object[object_code.zfill(4)] = CourtOrderStaticData(
            object_code=object_code.zfill(4),
            object_address=str(values.get("Адрес") or "").strip(),
            company=company,
            requisites=str(values.get("Реквизиты") or "").strip(),
            protocol=str(values.get("Протокол") or "").strip(),
            monthly_rate=rate,
            court=str(values.get("Подсудность") or "").strip(),
            court_address=str(values.get("Адрес суда") or "").strip(),
        )

    if not by_object:
        raise ValueError("В БЗ для судебных заявлений нет строк с кодами объектов.")
    return by_object


def calculate_state_duty(amount: Decimal) -> Decimal:
    amount = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if amount <= 0:
        return Decimal("0.00")
    if amount <= Decimal("100000"):
        duty = max(amount * Decimal("0.04"), Decimal("4000.00"))
        return (duty / Decimal("2")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if amount <= Decimal("300000"):
        duty = Decimal("4000") + (amount - Decimal("100000")) * Decimal("0.03")
    elif amount <= Decimal("500000"):
        duty = Decimal("10000") + (amount - Decimal("300000")) * Decimal("0.025")
    elif amount <= Decimal("1000000"):
        duty = Decimal("15000") + (amount - Decimal("500000")) * Decimal("0.02")
    else:
        duty = Decimal("25000") + (amount - Decimal("1000000")) * Decimal("0.01")
    return (duty / Decimal("2")).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def build_replacements(
    row: ClaimRow,
    static_data: CourtOrderStaticData,
    application_date: str,
    penalty_amount: Decimal = Decimal("0.00"),
) -> dict[str, str]:
    total_claim_amount = row.debt_amount + penalty_amount
    state_duty = calculate_state_duty(total_claim_amount)
    context = {
        "court_name": static_data.court or "мировому судье судебного участка",
        "court_address": static_data.court_address,
        "claimant_name": static_data.claimant_name,
        "claimant_company": static_data.company,
        "claimant_requisites": static_data.requisites,
        "claimant_ogrn": static_data.ogrn,
        "claimant_inn": static_data.inn,
        "claimant_kpp": static_data.kpp,
        "claimant_legal_address": static_data.legal_address,
        "claimant_director_name": static_data.director_name,
        "debtor_name": row.debtor_name,
        "debtor_birth_date": "неизвестно",
        "debtor_birth_place": "неизвестно",
        "debtor_registration_address": "неизвестно",
        "debtor_passport": "неизвестны",
        "account_number": row.account_number,
        "object_code": static_data.object_code,
        "object_address": static_data.object_address or row.object_address,
        "parking_place_number": row.parking_place_number,
        "debt_period": row.debt_period,
        "debt_amount": format_money(row.debt_amount),
        "debt_amount_words": money_to_words(row.debt_amount),
        "penalty_amount": format_money(penalty_amount),
        "penalty_amount_words": money_to_words(penalty_amount),
        "total_claim_amount": format_money(total_claim_amount),
        "total_claim_amount_words": money_to_words(total_claim_amount),
        "state_duty": format_money(state_duty),
        "state_duty_words": money_to_words(state_duty),
        "monthly_rate": format_money(static_data.monthly_rate) if static_data.monthly_rate is not None else "",
        "protocol": static_data.protocol,
        "application_date": application_date,
    }
    return {key: str(value) for key, value in context.items()}


def _project_root() -> Path:
    return Path(__file__).resolve().parents[3]


def generate_court_orders_zip_result(
    registry_path: str,
    template_path: str,
    static_data_path: str,
    output_zip_path: str,
    application_date: str | None = None,
) -> CourtOrderGenerationResult:
    template = Path(template_path)
    output_zip = Path(output_zip_path)
    if not template.exists():
        raise FileNotFoundError(f"Word-шаблон судебного заявления не найден: {template}")

    rows = read_registry(registry_path)
    if not rows:
        raise ValueError("В реестре нет строк с положительной суммой долга для генерации заявлений.")

    static_by_object = load_static_data(static_data_path)
    penalty_by_account = read_optional_money_by_account(registry_path, {"Сумма пени", "Пени"})

    temp_root = _project_root() / "storage" / "temp"
    temp_root.mkdir(parents=True, exist_ok=True)
    temp_dir = temp_root / f"court_orders_{uuid.uuid4().hex}"
    temp_dir.mkdir(parents=True, exist_ok=False)

    generated_files: list[Path] = []
    used_names: set[str] = set()
    skipped_count = 0
    total_debt_amount = Decimal("0.00")
    effective_date = application_date or date.today().strftime("%d.%m.%Y")

    try:
        for index, row in enumerate(rows, start=1):
            object_code = object_code_from_account(row.account_number)
            static_data = static_by_object.get(object_code)
            if static_data is None:
                skipped_count += 1
                continue

            doc = Document(str(template))
            penalty_amount = penalty_by_account.get(row.account_number, Decimal("0.00"))
            _replace_in_doc(doc, build_replacements(row, static_data, effective_date, penalty_amount))

            base_name = safe_filename(f"{index:03d}_{row.account_number}_{row.debtor_name}_судебный_приказ")
            file_name = base_name + ".docx"
            counter = 2
            while file_name.lower() in used_names:
                file_name = f"{base_name}_{counter}.docx"
                counter += 1
            used_names.add(file_name.lower())

            file_path = temp_dir / file_name
            doc.save(file_path)
            generated_files.append(file_path)
            total_debt_amount += row.debt_amount

        if not generated_files:
            raise ValueError(
                "Не создано ни одного заявления: в БЗ не найдены коды объектов из лицевых счетов реестра."
            )

        create_zip(generated_files, output_zip)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    return CourtOrderGenerationResult(
        zip_path=str(output_zip),
        documents_count=len(generated_files),
        skipped_count=skipped_count,
        total_debt_amount=total_debt_amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP),
    )


def generate_court_orders_zip(
    registry_path: str,
    template_path: str,
    static_data_path: str,
    output_zip_path: str,
    application_date: str | None = None,
) -> str:
    return generate_court_orders_zip_result(
        registry_path=registry_path,
        template_path=template_path,
        static_data_path=static_data_path,
        output_zip_path=output_zip_path,
        application_date=application_date,
    ).zip_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Генерация Word-заявлений о вынесении судебного приказа из Excel-реестра"
    )
    parser.add_argument("registry", help="Готовый Excel-реестр")
    parser.add_argument("--out", default="storage/output/court_orders.zip", help="Путь к итоговому ZIP-архиву")
    parser.add_argument(
        "--template",
        default="app/modules/court_orders/court_order_template.docx",
        help="Word-шаблон заявления",
    )
    parser.add_argument(
        "--static-data",
        default="storage/court_orders_static_data.xlsx",
        help="БЗ для судебных заявлений",
    )
    parser.add_argument("--application-date", default=None, help="Дата заявления, по умолчанию сегодня")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    try:
        result = generate_court_orders_zip_result(
            registry_path=args.registry,
            template_path=args.template,
            static_data_path=args.static_data,
            output_zip_path=args.out,
            application_date=args.application_date,
        )
    except (FileNotFoundError, ValueError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc

    print(f"Создано заявлений: {result.documents_count}")
    if result.skipped_count:
        print(f"Пропущено строк без БЗ по объекту: {result.skipped_count}")
    print(f"Итоговая сумма долга: {format_money(result.total_debt_amount)}")
    print(f"Архив: {result.zip_path}")


if __name__ == "__main__":
    main()
