from __future__ import annotations

from aiogram import Bot, Router
from aiogram.filters import CommandStart
from aiogram.types import Message

from bot.handlers.files import pop_user_session
from bot.keyboards.common import main_keyboard

router = Router(name="start")

_WELCOME_TEXT = (
    "Отправьте резюме в формате <code>.docx</code>, <code>.pdf</code> или <code>.mhtml</code>, "
    "и я извлеку данные и сформирую Excel-таблицу."
)


@router.message(CommandStart())
async def cmd_start(message: Message, bot: Bot) -> None:
    user_id = message.from_user.id  # type: ignore[union-attr]
    confirm_ids = pop_user_session(user_id)
    if confirm_ids:
        from bot.handlers.files import _delete_confirm_messages  # local import avoids circular at module level
        await _delete_confirm_messages(bot, message.chat.id, confirm_ids)
    await message.answer(_WELCOME_TEXT, parse_mode="HTML", reply_markup=main_keyboard())
