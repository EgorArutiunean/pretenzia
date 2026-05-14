from __future__ import annotations

import logging
import os
from datetime import datetime
from pathlib import Path

from aiogram import Bot, F, Router
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import CallbackQuery, FSInputFile, Message
from openpyxl import load_workbook

from app.bot.keyboards import main_menu_keyboard
from app.pipeline import (
    CourtOrdersNotImplementedError,
    PROJECT_ROOT,
    run_excel_to_registry,
    run_registry_to_claims,
    run_registry_to_court_orders,
)


router = Router()
logger = logging.getLogger(__name__)


class DocumentFlow(StatesGroup):
    waiting_for_file = State()


ACTION_PROMPTS = {
    "normalize": "Загрузите Excel-файл ОНВ из 1С.",
    "claims": (
        "Загрузите готовый Excel-реестр с колонками: Лицевой счет, ФИО, "
        "Адрес, Период задолженности, Сумма долга."
    ),
}


HELP_TEXT = (
    "Форматы файлов:\n"
    "\n"
    "1. ОНВ из 1С -> registry.xlsx. Для адресов используется справочник из OBJECT_ADDRESSES_PATH.\n"
    "2. registry.xlsx -> claims.zip. Обязательные колонки: Лицевой счет, ФИО, Адрес, "
    "Период задолженности, Сумма долга.\n"
    "3. Судебные заявления будут генерироваться из registry.xlsx, не из архива претензий."
)


def _admin_ids() -> set[int]:
    raw = os.getenv("ADMIN_IDS", "")
    result: set[int] = set()
    for item in raw.split(","):
        item = item.strip()
        if item:
            try:
                result.add(int(item))
            except ValueError:
                logger.warning("Invalid ADMIN_IDS item: %s", item)
    return result


def _is_allowed(user_id: int | None) -> bool:
    if user_id is None:
        return False
    admin_ids = _admin_ids()
    if not admin_ids:
        return True
    return user_id in admin_ids


async def _deny_if_needed(message: Message) -> bool:
    if _is_allowed(message.from_user.id if message.from_user else None):
        return False
    await message.answer("Доступ запрещён.")
    return True


async def _deny_callback_if_needed(callback: CallbackQuery) -> bool:
    if _is_allowed(callback.from_user.id if callback.from_user else None):
        return False
    await callback.answer("Доступ запрещён.", show_alert=True)
    return True


def _create_run_dir(user_id: int) -> Path:
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    run_dir = PROJECT_ROOT / "storage" / "runs" / f"user_{user_id}" / timestamp
    for child in ("input", "registry", "output", "logs", "errors"):
        (run_dir / child).mkdir(parents=True, exist_ok=True)
    return run_dir


def _source_path(run_dir: Path, original_file_name: str | None) -> Path:
    suffix = Path(original_file_name or "").suffix or ".xlsx"
    return run_dir / "input" / f"source{suffix}"


def _registry_error_count(registry_path: str) -> int:
    workbook = load_workbook(registry_path, read_only=True, data_only=True)
    if "Ошибки" not in workbook.sheetnames:
        return 0

    worksheet = workbook["Ошибки"]
    return max(worksheet.max_row - 1, 0)


async def _send_menu(message: Message) -> None:
    await message.answer("Выберите действие:", reply_markup=main_menu_keyboard())


@router.message(CommandStart())
async def start(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id if message.from_user else None
    logger.info("Start from user_id=%s admin_ids=%s", user_id, sorted(_admin_ids()))
    if await _deny_if_needed(message):
        return
    await state.clear()
    await _send_menu(message)


@router.callback_query(F.data == "action:help")
async def show_help(callback: CallbackQuery) -> None:
    if await _deny_callback_if_needed(callback):
        return
    await callback.message.answer(HELP_TEXT, reply_markup=main_menu_keyboard())
    await callback.answer()


@router.callback_query(F.data == "action:court_orders")
async def court_orders_stub(callback: CallbackQuery, state: FSMContext) -> None:
    if await _deny_callback_if_needed(callback):
        return
    await state.clear()
    await callback.message.answer(
        "Модуль судебных заявлений будет добавлен в следующей итерации. "
        "Источником данных будет registry.xlsx.",
        reply_markup=main_menu_keyboard(),
    )
    await callback.answer()


@router.callback_query(F.data.in_({"action:normalize", "action:claims"}))
async def choose_action(callback: CallbackQuery, state: FSMContext) -> None:
    if await _deny_callback_if_needed(callback):
        return

    action = callback.data.split(":", 1)[1]
    await state.update_data(action=action)
    await state.set_state(DocumentFlow.waiting_for_file)
    await callback.message.answer(ACTION_PROMPTS[action])
    await callback.answer()


@router.message(DocumentFlow.waiting_for_file, F.document)
async def receive_document(message: Message, bot: Bot, state: FSMContext) -> None:
    if await _deny_if_needed(message):
        return

    data = await state.get_data()
    action = data.get("action")
    if action not in {"normalize", "claims"}:
        await state.clear()
        await message.answer("Выберите действие заново.", reply_markup=main_menu_keyboard())
        return

    document = message.document
    run_dir = _create_run_dir(message.from_user.id)
    input_path = _source_path(run_dir, document.file_name)

    await message.answer("Файл получен. Обрабатываю...")
    await bot.download(document, destination=input_path)

    try:
        if action == "normalize":
            result_path = run_excel_to_registry(str(input_path), str(run_dir))
            error_count = _registry_error_count(result_path)
            caption = "Готово: registry.xlsx"
            if error_count:
                caption += f"\nЕсть строки на листе «Ошибки»: {error_count}."
            await message.answer_document(
                FSInputFile(result_path),
                caption=caption,
            )
        elif action == "claims":
            result_path = run_registry_to_claims(str(input_path), str(run_dir))
            await message.answer_document(
                FSInputFile(result_path),
                caption="Готово: claims.zip",
            )
        else:
            run_registry_to_court_orders(str(input_path), str(run_dir))
    except CourtOrdersNotImplementedError as exc:
        await message.answer(str(exc), reply_markup=main_menu_keyboard())
    except Exception as exc:
        logger.exception("Processing failed. run_dir=%s action=%s", run_dir, action)
        await message.answer(
            f"Ошибка обработки файла: {exc}\nПапка запуска: {run_dir}",
            reply_markup=main_menu_keyboard(),
        )
        return
    finally:
        await state.clear()

    await _send_menu(message)


@router.message(DocumentFlow.waiting_for_file)
async def receive_non_document(message: Message) -> None:
    if await _deny_if_needed(message):
        return
    await message.answer("Пожалуйста, загрузите Excel-файл документом.")
