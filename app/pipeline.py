from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from app.config import PROJECT_ROOT, load_settings
from app.modules.claims.generate_claims_from_registry import generate_claims_zip
from app.modules.court_orders.generate_court_orders_from_registry import generate_court_orders_zip
from app.modules.excel_normalizer.build_debt_registry_template import build_registry


DEFAULT_CLAIM_TEMPLATE = PROJECT_ROOT / "app" / "modules" / "claims" / "claim_template.docx"
DEFAULT_COURT_ORDER_TEMPLATE = PROJECT_ROOT / "app" / "modules" / "court_orders" / "court_order_template.docx"
DEFAULT_COURT_ORDERS_BASE_DATA = PROJECT_ROOT / "storage" / "court_orders" / "base_data.xlsx"
DEFAULT_COURT_ORDERS_JURISDICTION = PROJECT_ROOT / "storage" / "court_orders" / "jurisdiction.xlsx"


def _ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def _default_claim_date() -> str:
    return date.today().strftime("%d.%m.%Y")


def _default_payment_deadline() -> str:
    return (date.today() + timedelta(days=30)).strftime("%d.%m.%Y")


def _resolve_optional_base_data() -> Path | None:
    settings = load_settings()
    configured = settings.court_orders_base_data_path
    if configured is not None:
        if not configured.exists():
            raise FileNotFoundError(
                f"COURT_ORDERS_BASE_DATA_PATH does not exist: {configured}"
            )
        if not configured.is_file():
            raise ValueError(
                f"COURT_ORDERS_BASE_DATA_PATH must point to an .xlsx file: {configured}"
            )
        return configured
    return DEFAULT_COURT_ORDERS_BASE_DATA if DEFAULT_COURT_ORDERS_BASE_DATA.exists() else None


def run_excel_to_registry(input_excel_path: str, run_dir: str) -> str:
    """
    ОНВ из 1С -> registry.xlsx.

    Calls the excel_normalizer module directly and returns the registry path.
    """
    settings = load_settings()
    run_path = Path(run_dir)
    registry_dir = _ensure_dir(run_path / "registry")
    output_path = registry_dir / "registry.xlsx"

    object_addresses = settings.object_addresses_path
    if object_addresses is None:
        default_object_addresses = PROJECT_ROOT / "object_addresses.xlsx"
        if default_object_addresses.exists():
            object_addresses = default_object_addresses
    elif not object_addresses.exists():
        raise FileNotFoundError(f"OBJECT_ADDRESSES_PATH does not exist: {object_addresses}")
    elif not object_addresses.is_file():
        raise ValueError(f"OBJECT_ADDRESSES_PATH must point to an .xlsx or .json file: {object_addresses}")

    base_data = _resolve_optional_base_data()
    build_registry(
        input_path=input_excel_path,
        output_path=str(output_path),
        object_addresses_path=str(object_addresses) if object_addresses else None,
        base_data_path=str(base_data) if base_data else None,
    )
    return str(output_path)


def run_registry_to_claims(
    registry_path: str,
    run_dir: str,
    *,
    claim_date: str | None = None,
    payment_deadline: str | None = None,
) -> str:
    """
    registry.xlsx -> claims.zip.

    Calls the claims module directly and returns the ZIP path.
    """
    settings = load_settings()
    run_path = Path(run_dir)
    output_dir = _ensure_dir(run_path / "output")
    output_zip_path = output_dir / "claims.zip"
    base_data = _resolve_optional_base_data()

    generate_claims_zip(
        registry_path=registry_path,
        template_path=str(DEFAULT_CLAIM_TEMPLATE),
        output_zip_path=str(output_zip_path),
        claim_date=claim_date or settings.claim_date or _default_claim_date(),
        payment_deadline=(
            payment_deadline
            or settings.payment_deadline
            or _default_payment_deadline()
        ),
        base_data_path=str(base_data) if base_data else None,
    )
    return str(output_zip_path)


def run_registry_to_court_orders(
    registry_path: str,
    run_dir: str,
    *,
    application_date: str | None = None,
) -> str:
    """
    registry.xlsx -> court_orders.zip.

    Calls the court_orders module directly and returns the ZIP path.
    """
    settings = load_settings()
    run_path = Path(run_dir)
    output_dir = _ensure_dir(run_path / "output")
    output_zip_path = output_dir / "court_orders.zip"

    base_data = settings.court_orders_base_data_path or DEFAULT_COURT_ORDERS_BASE_DATA
    jurisdiction = (
        settings.court_orders_jurisdiction_path
        or DEFAULT_COURT_ORDERS_JURISDICTION
    )
    if not base_data.exists():
        raise FileNotFoundError(
            f"Основная БЗ для судебных заявлений не найдена: {base_data}. "
            "Загрузите файл через /courtdata."
        )
    if not jurisdiction.exists():
        raise FileNotFoundError(
            f"БЗ подсудности не найдена: {jurisdiction}. "
            "Загрузите файл через /jurisdiction."
        )
    if not base_data.is_file():
        raise ValueError(f"Путь к основной БЗ должен указывать на .xlsx: {base_data}")
    if not jurisdiction.is_file():
        raise ValueError(f"Путь к БЗ подсудности должен указывать на .xlsx: {jurisdiction}")

    generate_court_orders_zip(
        registry_path=registry_path,
        template_path=str(DEFAULT_COURT_ORDER_TEMPLATE),
        base_data_path=str(base_data),
        jurisdiction_path=str(jurisdiction),
        output_zip_path=str(output_zip_path),
        application_date=application_date,
    )
    return str(output_zip_path)
