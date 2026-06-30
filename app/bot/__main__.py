from __future__ import annotations

import asyncio
import logging

from aiogram import Bot, Dispatcher
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.fsm.storage.memory import MemoryStorage

from app.config import load_settings
from app.bot.handlers import router


async def main() -> None:
    settings = load_settings(require_bot=True)

    logging.basicConfig(level=logging.INFO)
    logging.info("Access mode: restricted to %d admin user(s)", len(settings.admin_ids))
    session = AiohttpSession()
    if not settings.telegram_ssl_verify:
        session._connector_init["ssl"] = False
        logging.warning("Telegram SSL verification is disabled by TELEGRAM_SSL_VERIFY=false")

    bot = Bot(token=settings.bot_token, session=session)
    dispatcher = Dispatcher(storage=MemoryStorage())
    dispatcher.include_router(router)
    await dispatcher.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
