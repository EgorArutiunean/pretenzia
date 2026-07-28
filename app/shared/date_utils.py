from __future__ import annotations

from datetime import date, datetime, timedelta


DATE_FORMAT = "%d.%m.%Y"


def parse_operator_date(
    value: str | None,
    *,
    today: date | None = None,
) -> date:
    text = (value or "").strip().lower()
    if text == "сегодня":
        return today or date.today()
    try:
        return datetime.strptime(text, DATE_FORMAT).date()
    except ValueError as exc:
        raise ValueError("Введите дату в формате ДД.ММ.ГГГГ или слово «сегодня».") from exc


def parse_payment_deadline(
    value: str | None,
    *,
    claim_date: date,
) -> date:
    text = (value or "").strip().lower().replace(" ", "")
    if text == "+30":
        return claim_date + timedelta(days=30)
    deadline = parse_operator_date(value)
    if deadline < claim_date:
        raise ValueError("Срок оплаты не может быть раньше даты претензии.")
    return deadline


def format_date(value: date) -> str:
    return value.strftime(DATE_FORMAT)
