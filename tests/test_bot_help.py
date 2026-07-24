from __future__ import annotations

import unittest

from app.bot.handlers import HELP_TEXT


class BotHelpTests(unittest.TestCase):
    def test_help_text_fits_telegram_and_covers_all_workflows(self) -> None:
        self.assertLessEqual(len(HELP_TEXT), 4096)
        self.assertIn("Обновление данных", HELP_TEXT)
        self.assertIn("Создать реестр из ОНВ", HELP_TEXT)
        self.assertIn("Создать претензии из реестра", HELP_TEXT)
        self.assertIn("Создать заявления в суд", HELP_TEXT)
        self.assertIn("errors.xlsx", HELP_TEXT)
        self.assertIn("/start", HELP_TEXT)


if __name__ == "__main__":
    unittest.main()
