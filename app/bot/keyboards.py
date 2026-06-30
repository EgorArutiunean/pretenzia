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
                    text="📍 Обновить справочник адресов",
                    callback_data="action:dictionary",
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
                    text="ℹ️ Инструкция",
                    callback_data="action:help",
                )
            ],
        ]
    )
