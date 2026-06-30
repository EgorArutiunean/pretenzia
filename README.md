# pretenzia

Сервис автоматизации документооборота по задолженности.

## MVP

В текущей версии реализован только модуль `claims`: генерация Word-претензий из готового Excel-реестра и упаковка документов в ZIP.

## Конфигурация

Скопируйте `.env.example` в `.env` и заполните реальные значения:

```powershell
copy .env.example .env
```

Обязательные переменные для Telegram-бота:

- `BOT_TOKEN` — токен Telegram-бота.
- `ADMIN_IDS` — разрешенные Telegram user id через запятую. Бот не запускается без этого списка.
- `OBJECT_ADDRESSES_PATH` — путь к справочнику адресов объектов для нормализации ОНВ.

Для production оставляйте `TELEGRAM_SSL_VERIFY=true`.

## Проверка перед деплоем

```powershell
python -m unittest discover -v
python -m app.modules.excel_normalizer.build_debt_registry_template "ОНВ 1297.xlsx" --object-addresses "Справочник.xlsx" --out storage/registry/registry.xlsx
python -m app.main storage/registry/registry.xlsx --out storage/output/claims.zip
```

Запуск:

```powershell
python -m app.main registry_template.xlsx --out storage/output/claims.zip
```

Прямой запуск модуля:

```powershell
python -m app.modules.claims.generate_claims_from_registry registry_template.xlsx --out storage/output/claims.zip
```

Дополнительные параметры:

```powershell
--template app/modules/claims/claim_template.docx
--claim-date 25.04.2026
--payment-deadline 25.05.2026
```

## Структура

- `app/modules/excel_normalizer` — будущая нормализация сырого Excel-отчета из 1С.
- `app/modules/claims` — MVP-модуль генерации претензий.
- `app/modules/court_orders` — будущая генерация заявлений о вынесении судебного приказа.
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
