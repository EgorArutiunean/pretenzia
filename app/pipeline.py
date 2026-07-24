from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from app.config import PROJECT_ROOT, load_settings
from app.modules.claims.generate_claims_from_registry import generate_claims_zip
from app.modules.court_orders.generate_court_orders_from_registry import generate_court_orders_zip
from app.modules.excel_normalizer.build_debt_registry_template import build_registry


DEFAULT_CLAIM_TEMPLATE = PROJECT_ROOT / "app" / "modules" / "claims" / "claim_template.docx"
DEFAULT_COURT_ORDER_TEMPLATE = PROJECT_ROOT / "app" / "modules" / "court_orders" / "court_order_template.docx"
DEFAULT_COURT_ORDERS_STATIC_DATA = PROJECT_ROOT / "storage" / "court_orders_static_data.xlsx"


def _ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def _default_claim_date() -> str:
    return date.today().strftime("%d.%m.%Y")


def _default_payment_deadline() -> str:
    return (date.today() + timedelta(days=30)).strftime("%d.%m.%Y")


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

    build_registry(
        input_path=input_excel_path,
        output_path=str(output_path),
        object_addresses_path=str(object_addresses) if object_addresses else None,
    )
    return str(output_path)


def run_registry_to_claims(registry_path: str, run_dir: str) -> str:
    """
    registry.xlsx -> claims.zip.

    Calls the claims module directly and returns the ZIP path.
    """
    settings = load_settings()
    run_path = Path(run_dir)
    output_dir = _ensure_dir(run_path / "output")
    output_zip_path = output_dir / "claims.zip"

    generate_claims_zip(
        registry_path=registry_path,
        template_path=str(DEFAULT_CLAIM_TEMPLATE),
        output_zip_path=str(output_zip_path),
        claim_date=settings.claim_date or _default_claim_date(),
        payment_deadline=settings.payment_deadline or _default_payment_deadline(),
    )
    return str(output_zip_path)


def run_registry_to_court_orders(registry_path: str, run_dir: str) -> str:
    """
    registry.xlsx -> court_orders.zip.

    Calls the court_orders module directly and returns the ZIP path.
    """
    settings = load_settings()
    run_path = Path(run_dir)
    output_dir = _ensure_dir(run_path / "output")
    output_zip_path = output_dir / "court_orders.zip"

    static_data = settings.court_orders_static_data_path or DEFAULT_COURT_ORDERS_STATIC_DATA
    if not static_data.exists():
        raise FileNotFoundError(
            f"БЗ для судебных заявлений не найдена: {static_data}. "
            "Загрузите файл через /courtdata или задайте COURT_ORDERS_STATIC_DATA_PATH."
        )
    if not static_data.is_file():
        raise ValueError(f"COURT_ORDERS_STATIC_DATA_PATH must point to an .xlsx file: {static_data}")

    generate_court_orders_zip(
        registry_path=registry_path,
        template_path=str(DEFAULT_COURT_ORDER_TEMPLATE),
        static_data_path=str(static_data),
        output_zip_path=str(output_zip_path),
    )
    return str(output_zip_path)
