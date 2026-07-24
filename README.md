# pretenzia

Сервис автоматизации документооборота по задолженности.

## MVP

В текущей версии реализованы:

- `claims` — генерация Word-претензий из готового Excel-реестра и упаковка документов в ZIP;
- `court_orders` — генерация заявлений о вынесении судебного приказа из того же реестра.

## Конфигурация

Скопируйте `.env.example` в `.env` и заполните реальные значения:

```powershell
copy .env.example .env
```

Обязательные переменные для Telegram-бота:

- `BOT_TOKEN` — токен Telegram-бота.
- `ADMIN_IDS` — разрешенные Telegram user id через запятую. Бот не запускается без этого списка.
- `OBJECT_ADDRESSES_PATH` — путь к справочнику адресов объектов для нормализации ОНВ. Для Docker/Coolify используйте `/app/storage/object_addresses.xlsx`.
- `COURT_ORDERS_BASE_DATA_PATH` — основная БЗ объектов, компаний и реквизитов. Для Docker/Coolify используйте `/app/storage/court_orders/base_data.xlsx`.
- `COURT_ORDERS_JURISDICTION_PATH` — БЗ судебных участков. Для Docker/Coolify используйте `/app/storage/court_orders/jurisdiction.xlsx`.

Устаревшая переменная `COURT_ORDERS_STATIC_DATA_PATH` временно поддерживается как путь к основной БЗ.

Не задавайте `CLAIM_DATE` и `PAYMENT_DEADLINE` в production, если даты должны рассчитываться автоматически. По умолчанию дата претензии — текущий день, срок оплаты — текущий день + 30 дней. Эти переменные нужны только для ручного переопределения дат.

Для production оставляйте `TELEGRAM_SSL_VERIFY=true`.

## Проверка перед деплоем

```powershell
python -m unittest discover -v
python -m app.modules.excel_normalizer.build_debt_registry_template "ОНВ 1297.xlsx" --object-addresses "Справочник.xlsx" --out storage/registry/registry.xlsx
python -m app.main storage/registry/registry.xlsx --out storage/output/claims.zip
```

## Docker deploy

Перед запуском заполните `.env` и положите справочник адресов в persistent storage как `storage/object_addresses.xlsx`.

```powershell
docker compose build
docker compose up -d
docker compose logs -f bot
```

В Docker build context не попадают `.env`, ОНВ, справочники, `storage/` и ZIP-архивы.

Справочник адресов также можно обновить через Telegram-бота: команда `/dictionary` или кнопка `Обновить справочник адресов`. Загруженный файл сохраняется в `OBJECT_ADDRESSES_PATH`.

Справочники третьего модуля обновляются через Telegram-бота:

- `/courtdata` или кнопка `Обновить основную БЗ`;
- `/jurisdiction` или кнопка `Обновить подсудность`.

Перед заменой файл проверяется. Предыдущие версии сохраняются в `storage/court_orders/backups`, последние пять резервных копий каждого справочника остаются доступными для отката.

## Autodeploy

Автодеплой настроен через GitHub Actions: после успешного `CI` на ветке `master` workflow `Deploy` подключается к серверу по SSH и перезапускает `docker compose`.

Инструкция по подготовке сервера и GitHub Secrets: [docs/deploy.md](docs/deploy.md).

## Очистка персональных данных

По умолчанию команда показывает, какие пользовательские запуски старше 14 дней будут удалены:

```powershell
python -m app.maintenance.cleanup_storage --older-than-days 14
```

Фактическое удаление:

```powershell
python -m app.maintenance.cleanup_storage --older-than-days 14 --apply
```

Запуск:

```powershell
python -m app.main registry_template.xlsx --out storage/output/claims.zip
```

Прямой запуск модуля:

```powershell
python -m app.modules.claims.generate_claims_from_registry registry_template.xlsx --out storage/output/claims.zip
```

Дополнительные параметры для ручного запуска с фиксированными датами:

```powershell
--template app/modules/claims/claim_template.docx
--claim-date 25.04.2026
--payment-deadline 25.05.2026
```

Генерация заявлений о вынесении судебного приказа:

```powershell
python -m app.main registry_template.xlsx --mode court-orders --base-data storage/court_orders/base_data.xlsx --jurisdiction storage/court_orders/jurisdiction.xlsx --out storage/output/court_orders.zip
```

Прямой запуск модуля:

```powershell
python -m app.modules.court_orders.generate_court_orders_from_registry registry_template.xlsx --base-data storage/court_orders/base_data.xlsx --jurisdiction storage/court_orders/jurisdiction.xlsx --out storage/output/court_orders.zip
```

## Структура

- `app/modules/excel_normalizer` — будущая нормализация сырого Excel-отчета из 1С.
- `app/modules/claims` — MVP-модуль генерации претензий.
- `app/modules/court_orders` — генерация заявлений о вынесении судебного приказа.
- `app/shared` — общие утилиты.
- `storage/temp` — временные DOCX при сборке ZIP.
- `storage/output` — итоговые архивы.

## Excel Normalizer

Нормализатор поддерживает справочник адресов объектов в формате `.xlsx` или `.json`.

Лицевой счет должен состоять из 8 цифр:

- первые 4 цифры — `object_code`;
- последние 4 цифры — номер машиноместа.

Пример запуска:

```powershell
python -m app.modules.excel_normalizer.build_debt_registry_template "storage/input/onv_report.xlsx" --object-addresses object_addresses.xlsx --out storage/registry/registry.xlsx
```

Если адрес для `object_code` не найден, строка не попадает в основной лист `Реестр`, а записывается в лист `Ошибки`.
