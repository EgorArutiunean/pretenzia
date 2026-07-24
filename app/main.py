from __future__ import annotations

import argparse
import sys
from datetime import date, timedelta

from app.modules.claims.generate_claims_from_registry import format_money, generate_claims_zip_result
from app.modules.court_orders.generate_court_orders_from_registry import generate_court_orders_zip_result


def parse_args() -> argparse.Namespace:
    today = date.today()
    default_deadline = today + timedelta(days=30)

    parser = argparse.ArgumentParser(description="Генерация ZIP-архива документов из Excel-реестра")
    parser.add_argument("registry", help="Готовый Excel-реестр")
    parser.add_argument(
        "--mode",
        choices=("claims", "court-orders"),
        default="claims",
        help="Тип документов. По умолчанию claims, чтобы не ломать MVP-сценарий.",
    )
    parser.add_argument("--out", default="storage/output/claims.zip", help="Путь к итоговому ZIP-архиву")
    parser.add_argument("--template", default="app/modules/claims/claim_template.docx", help="Word-шаблон претензии")
    parser.add_argument("--claim-date", default=today.strftime("%d.%m.%Y"), help="Дата претензии")
    parser.add_argument("--payment-deadline", default=default_deadline.strftime("%d.%m.%Y"), help="Срок оплаты")
    parser.add_argument(
        "--static-data",
        default="storage/court_orders_static_data.xlsx",
        help="БЗ для судебных заявлений, используется только с --mode court-orders",
    )
    parser.add_argument("--application-date", default=None, help="Дата заявления, используется только с --mode court-orders")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    try:
        if args.mode == "court-orders":
            result = generate_court_orders_zip_result(
                registry_path=args.registry,
                template_path=(
                    args.template
                    if args.template != "app/modules/claims/claim_template.docx"
                    else "app/modules/court_orders/court_order_template.docx"
                ),
                static_data_path=args.static_data,
                output_zip_path=args.out,
                application_date=args.application_date,
            )
        else:
            result = generate_claims_zip_result(
                registry_path=args.registry,
                template_path=args.template,
                output_zip_path=args.out,
                claim_date=args.claim_date,
                payment_deadline=args.payment_deadline,
            )
    except (FileNotFoundError, ValueError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc

    if args.mode == "court-orders":
        print(f"Создано заявлений: {result.documents_count}")
        if result.skipped_count:
            print(f"Пропущено строк без БЗ по объекту: {result.skipped_count}")
        print(f"Итоговая сумма долга: {format_money(result.total_debt_amount)}")
    else:
        print(f"Создано претензий: {result.documents_count}")
        print(f"Итоговая сумма долга: {format_money(result.total_amount)}")
    print(f"Архив: {result.zip_path}")


if __name__ == "__main__":
    main()
