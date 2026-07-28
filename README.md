# pretenzia

Сервис автоматизации документооборота по задолженности.

Пошаговая инструкция для оператора: [docs/user-guide.md](docs/user-guide.md).

## MVP

В текущей версии реализованы:

- `claims` — генерация Word-претензий из готового Excel-реестра и упаковка документов в ZIP;
- `court_orders` — генерация заявлений о вынесении судебного приказа из того же реестра.

## Работа через Telegram-бота

1. Отправьте `/start`.
2. В разделе `Обновление данных` загрузите изменившиеся справочники.
3. Создайте `registry.xlsx` из отчета ОНВ.
4. Выберите даты и передайте реестр в сценарий претензий или судебных заявлений.
5. Проверьте листы `Ошибки` и `Предупреждения` перед использованием документов.

Подробные требования к файлам и порядок проверки результата приведены в
[инструкции пользователя](docs/user-guide.md).

## Конфигурация

Скопируйте `.env.example` в `.env` и заполните реальные значения:

```powershell
copy .env.example .env
```

Обязательные переменные для Telegram-бота:

- `BOT_TOKEN` — токен Telegram-бота.
- `ADMIN_IDS` — Telegram ID администраторов через запятую. Они формируют документы и обновляют справочники.
- `USER_IDS` — Telegram ID обычных пользователей через запятую. Они могут только формировать документы.
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

Бот автоматически удаляет рабочие файлы старше 24 часов. Ручная проверка:

```powershell
python -m app.maintenance.cleanup_storage --older-than-hours 24
```

Фактическое удаление:

```powershell
python -m app.maintenance.cleanup_storage --older-than-hours 24 --apply
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
python -m app.modules.excel_normalizer.build_debt_registry_template "storage/input/onv_report.xlsx" --object-addresses object_addresses.xlsx --base-data storage/court_orders/base_data.xlsx --out storage/registry/registry.xlsx
```

Если адрес для `object_code` не найден, строка не попадает в основной лист `Реестр`, а записывается в лист `Ошибки`.
Компания, ИНН и код компании определяются по первым четырём цифрам лицевого счёта из основной БЗ.
