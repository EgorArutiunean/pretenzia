from __future__ import annotations

import unittest
from datetime import date

from app.shared.date_utils import (
    format_date,
    parse_operator_date,
    parse_payment_deadline,
)


class DateUtilsTests(unittest.TestCase):
    def test_operator_date_accepts_explicit_date_and_today(self) -> None:
        self.assertEqual(parse_operator_date("25.04.2026"), date(2026, 4, 25))
        self.assertEqual(
            parse_operator_date("сегодня", today=date(2026, 7, 28)),
            date(2026, 7, 28),
        )

    def test_payment_deadline_supports_plus_30(self) -> None:
        deadline = parse_payment_deadline(
            "+30",
            claim_date=date(2026, 4, 25),
        )

        self.assertEqual(format_date(deadline), "25.05.2026")

    def test_payment_deadline_rejects_date_before_claim(self) -> None:
        with self.assertRaisesRegex(ValueError, "не может быть раньше"):
            parse_payment_deadline(
                "24.04.2026",
                claim_date=date(2026, 4, 25),
            )

    def test_invalid_calendar_date_has_clear_error(self) -> None:
        with self.assertRaisesRegex(ValueError, "ДД.ММ.ГГГГ"):
            parse_operator_date("31.02.2026")


if __name__ == "__main__":
    unittest.main()
