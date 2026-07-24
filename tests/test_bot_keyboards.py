from __future__ import annotations

import unittest

from app.bot.keyboards import data_updates_menu_keyboard, main_menu_keyboard


def _callbacks(markup) -> list[str]:
    return [
        button.callback_data
        for row in markup.inline_keyboard
        for button in row
        if button.callback_data
    ]


class BotKeyboardsTests(unittest.TestCase):
    def test_main_menu_groups_reference_updates(self) -> None:
        callbacks = _callbacks(main_menu_keyboard())

        self.assertIn("menu:data_updates", callbacks)
        self.assertNotIn("action:dictionary", callbacks)
        self.assertNotIn("action:court_data", callbacks)
        self.assertNotIn("action:jurisdiction", callbacks)

    def test_data_updates_menu_contains_all_reference_actions_and_back(self) -> None:
        callbacks = _callbacks(data_updates_menu_keyboard())

        self.assertEqual(
            callbacks,
            [
                "action:dictionary",
                "action:court_data",
                "action:jurisdiction",
                "menu:main",
            ],
        )


if __name__ == "__main__":
    unittest.main()
