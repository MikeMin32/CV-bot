from __future__ import annotations

from aiogram.types import KeyboardButton, ReplyKeyboardMarkup

FINISH_TEXT = "✅ Завершить загрузку"
CLEAR_TEXT = "🗑 Очистить очередь"


def main_keyboard() -> ReplyKeyboardMarkup:
    """Persistent bottom keyboard shown on /start and kept throughout the session."""
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=FINISH_TEXT)],
            [KeyboardButton(text=CLEAR_TEXT)],
        ],
        resize_keyboard=True,
        persistent=True,
    )
