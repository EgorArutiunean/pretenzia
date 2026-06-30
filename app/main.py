from __future__ import annotations

import argparse
import sys
from datetime import date, timedelta

from app.modules.claims.generate_claims_from_registry import format_money, generate_claims_zip_result


def parse_args() -> argparse.Namespace:
    today = date.today()
    default_deadline = today + timedelta(days=30)

    parser = argparse.ArgumentParser(description="MVP: генерация ZIP-архива Word-претензий из Excel-реестра")
    parser.add_argument("registry", help="Готовый Excel-реестр")
    parser.add_argument("--out", default="storage/output/claims.zip", help="Путь к итоговому ZIP-архиву")
    parser.add_argument("--template", default="app/modules/claims/claim_template.docx", help="Word-шаблон претензии")
    parser.add_argument("--claim-date", default=today.strftime("%d.%m.%Y"), help="Дата претензии")
    parser.add_argument("--payment-deadline", default=default_deadline.strftime("%d.%m.%Y"), help="Срок оплаты")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    try:
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

    print(f"Создано претензий: {result.documents_count}")
    print(f"Итоговая сумма долга: {format_money(result.total_amount)}")
    print(f"Архив: {result.zip_path}")


if __name__ == "__main__":
    main()
