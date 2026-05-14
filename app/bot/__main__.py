from __future__ import annotations

import asyncio
import logging
import os

from aiogram import Bot, Dispatcher
from aiogram.fsm.storage.memory import MemoryStorage
from dotenv import load_dotenv

from app.bot.handlers import router
from app.pipeline import PROJECT_ROOT


async def main() -> None:
    load_dotenv(PROJECT_ROOT / ".env", override=True)
    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError("BOT_TOKEN is not set in .env")

    logging.basicConfig(level=logging.INFO)
    admin_ids = os.getenv("ADMIN_IDS", "")
    logging.info("Access mode: %s", "open to all users" if not admin_ids.strip() else f"ADMIN_IDS={admin_ids}")
    bot = Bot(token=token)
    dispatcher = Dispatcher(storage=MemoryStorage())
    dispatcher.include_router(router)
    await dispatcher.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
