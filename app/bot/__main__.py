from __future__ import annotations

import asyncio
import logging

from aiogram import Bot, Dispatcher
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.fsm.storage.memory import MemoryStorage

from app.config import load_settings
from app.bot.handlers import router
from app.maintenance.cleanup_storage import cleanup_runs


async def cleanup_loop() -> None:
    while True:
        try:
            result = await asyncio.to_thread(
                cleanup_runs,
                older_than_hours=24,
                dry_run=False,
            )
            if result.deleted:
                logging.info("Storage cleanup deleted %d expired run(s)", len(result.deleted))
        except Exception:
            logging.error("Storage cleanup failed", exc_info=False)
        await asyncio.sleep(60 * 60)


async def main() -> None:
    settings = load_settings(require_bot=True)

    logging.basicConfig(level=logging.INFO)
    logging.info(
        "Access mode: restricted to %d user(s), including %d admin(s)",
        len(settings.allowed_user_ids),
        len(settings.admin_ids),
    )
    session = AiohttpSession()
    if not settings.telegram_ssl_verify:
        session._connector_init["ssl"] = False
        logging.warning("Telegram SSL verification is disabled by TELEGRAM_SSL_VERIFY=false")

    bot = Bot(token=settings.bot_token, session=session)
    dispatcher = Dispatcher(storage=MemoryStorage())
    dispatcher.include_router(router)
    cleanup_task = asyncio.create_task(cleanup_loop())
    try:
        await dispatcher.start_polling(bot)
    finally:
        cleanup_task.cancel()
        await asyncio.gather(cleanup_task, return_exceptions=True)


if __name__ == "__main__":
    asyncio.run(main())
