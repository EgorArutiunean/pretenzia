from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _parse_admin_ids(raw: str) -> set[int]:
    result: set[int] = set()
    for item in raw.split(","):
        item = item.strip()
        if not item:
            continue
        try:
            result.add(int(item))
        except ValueError as exc:
            raise RuntimeError(f"ADMIN_IDS contains a non-numeric Telegram user id: {item!r}") from exc
    return result


def _resolve_optional_path(raw: str | None) -> Path | None:
    if not raw:
        return None
    path = Path(raw)
    return path if path.is_absolute() else PROJECT_ROOT / path


@dataclass(frozen=True)
class Settings:
    bot_token: str | None
    admin_ids: set[int]
    claim_date: str | None
    payment_deadline: str | None
    object_addresses_path: Path | None
    telegram_ssl_verify: bool
    max_upload_mb: int


def load_settings(*, require_bot: bool = False) -> Settings:
    load_dotenv(PROJECT_ROOT / ".env", override=False)

    settings = Settings(
        bot_token=os.getenv("BOT_TOKEN") or None,
        admin_ids=_parse_admin_ids(os.getenv("ADMIN_IDS", "")),
        claim_date=os.getenv("CLAIM_DATE") or None,
        payment_deadline=os.getenv("PAYMENT_DEADLINE") or None,
        object_addresses_path=_resolve_optional_path(os.getenv("OBJECT_ADDRESSES_PATH")),
        telegram_ssl_verify=os.getenv("TELEGRAM_SSL_VERIFY", "true").strip().lower() not in {"0", "false", "no"},
        max_upload_mb=int(os.getenv("MAX_UPLOAD_MB", "20")),
    )

    if require_bot:
        if not settings.bot_token:
            raise RuntimeError("BOT_TOKEN is not set in .env")
        if not settings.admin_ids:
            raise RuntimeError("ADMIN_IDS must contain at least one Telegram user id for safe bot startup")
        if settings.object_addresses_path is not None and not settings.object_addresses_path.exists():
            raise RuntimeError(f"OBJECT_ADDRESSES_PATH does not exist: {settings.object_addresses_path}")

    return settings
