from __future__ import annotations

import os
from datetime import date, timedelta
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover - dotenv is optional for CLI-only use
    load_dotenv = None

from app.modules.claims.generate_claims_from_registry import generate_claims_zip
from app.modules.excel_normalizer.build_debt_registry_template import build_registry


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CLAIM_TEMPLATE = PROJECT_ROOT / "app" / "modules" / "claims" / "claim_template.docx"


class CourtOrdersNotImplementedError(NotImplementedError):
    """Raised while the court orders module is only an architectural stub."""


def _load_env() -> None:
    if load_dotenv is not None:
        load_dotenv(PROJECT_ROOT / ".env", override=True)


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
    _load_env()
    run_path = Path(run_dir)
    registry_dir = _ensure_dir(run_path / "registry")
    output_path = registry_dir / "registry.xlsx"

    object_addresses_path = os.getenv("OBJECT_ADDRESSES_PATH")
    object_addresses: Path | None = None
    if object_addresses_path:
        object_addresses = Path(object_addresses_path)
        if not object_addresses.is_absolute():
            object_addresses = PROJECT_ROOT / object_addresses
    else:
        default_object_addresses = PROJECT_ROOT / "object_addresses.xlsx"
        if default_object_addresses.exists():
            object_addresses = default_object_addresses

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
    _load_env()
    run_path = Path(run_dir)
    output_dir = _ensure_dir(run_path / "output")
    output_zip_path = output_dir / "claims.zip"

    generate_claims_zip(
        registry_path=registry_path,
        template_path=str(DEFAULT_CLAIM_TEMPLATE),
        output_zip_path=str(output_zip_path),
        claim_date=os.getenv("CLAIM_DATE", _default_claim_date()),
        payment_deadline=os.getenv("PAYMENT_DEADLINE", _default_payment_deadline()),
    )
    return str(output_zip_path)


def run_registry_to_court_orders(registry_path: str, run_dir: str) -> str:
    """
    registry.xlsx -> court_orders.zip.

    Stub for the future court_orders module. The source remains registry.xlsx.
    """
    raise CourtOrdersNotImplementedError(
        "Модуль судебных заявлений будет добавлен в следующей итерации. "
        "Источником данных будет registry.xlsx."
    )
