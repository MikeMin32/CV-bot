from __future__ import annotations

from aiogram import Router
from aiogram.filters import CommandStart
from aiogram.types import Message

from bot.keyboards.common import main_keyboard

router = Router(name="start")

_WELCOME_TEXT = (
    "Отправьте резюме в формате <code>.docx</code>, и я извлеку данные "
    "и сформирую Excel-таблицу."
)


@router.message(CommandStart())
async def cmd_start(message: Message) -> None:
    await message.answer(_WELCOME_TEXT, parse_mode="HTML", reply_markup=main_keyboard())
