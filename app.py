from __future__ import annotations

import asyncio
import sys

from aiogram import Bot, Dispatcher
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode

from bot.handlers import files, start
from core.config import config
from core.logging import get_logger, setup_logging


async def main() -> None:
    setup_logging(config.LOG_LEVEL)
    logger = get_logger(__name__)

    config.ensure_upload_dir()
    logger.info("Upload directory: %s", config.UPLOAD_DIR)

    bot = Bot(
        token=config.BOT_TOKEN,
        default=DefaultBotProperties(parse_mode=ParseMode.HTML),
    )

    dp = Dispatcher()
    dp.include_router(start.router)
    dp.include_router(files.router)

    logger.info("Bot starting…")
    try:
        await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())
    finally:
        await bot.session.close()
        logger.info("Bot stopped.")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        sys.exit(0)
