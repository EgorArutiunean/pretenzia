from __future__ import annotations

from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup


def main_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📊 Создать реестр из ОНВ",
                    callback_data="action:normalize",
                )
            ],
            [
                InlineKeyboardButton(
                    text="📄 Создать претензии из реестра",
                    callback_data="action:claims",
                )
            ],
            [
                InlineKeyboardButton(
                    text="⚖️ Создать заявления в суд",
                    callback_data="action:court_orders",
                )
            ],
            [
                InlineKeyboardButton(
                    text="🔄 Обновление данных",
                    callback_data="menu:data_updates",
                )
            ],
            [
                InlineKeyboardButton(
                    text="ℹ️ Инструкция",
                    callback_data="action:help",
                )
            ],
        ]
    )


def data_updates_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📍 Справочник адресов",
                    callback_data="action:dictionary",
                )
            ],
            [
                InlineKeyboardButton(
                    text="🏢 Основная БЗ",
                    callback_data="action:court_data",
                )
            ],
            [
                InlineKeyboardButton(
                    text="🏛️ Подсудность",
                    callback_data="action:jurisdiction",
                )
            ],
            [
                InlineKeyboardButton(
                    text="← Назад",
                    callback_data="menu:main",
                )
            ],
        ]
    )
