from __future__ import annotations

import asyncio
import shutil
from datetime import datetime
from pathlib import Path

from aiogram import Bot, F, Router
from aiogram.types import BufferedInputFile, Document, Message

from bot.keyboards.common import FINISH_TEXT
from core.config import config
from core.logging import get_logger
from services.excel_exporter import build_excel
from services.resume_extractor import extract_resume

logger = get_logger(__name__)
router = Router(name="files")

# In-memory session: user_id → list of saved file paths
_sessions: dict[int, list[Path]] = {}

# Tracks message_ids of "Файл загружен" confirmations per user for cleanup
_confirm_message_ids: dict[int, list[int]] = {}

MAX_FILE_BYTES = config.MAX_FILE_SIZE_MB * 1024 * 1024


# ---------------------------------------------------------------------------
# Document upload handler
# ---------------------------------------------------------------------------

@router.message(F.document)
async def handle_document(message: Message, bot: Bot) -> None:
    user_id = message.from_user.id  # type: ignore[union-attr]
    doc: Document = message.document  # type: ignore[assignment]

    # Validate extension
    file_name: str = doc.file_name or ""
    _ALLOWED = (".docx", ".pdf", ".mhtml", ".mht")
    if not any(file_name.lower().endswith(ext) for ext in _ALLOWED):
        await message.answer(
            "⚠️ Підтримувані формати: <code>.docx</code>, <code>.pdf</code>, <code>.mhtml</code>.",
            parse_mode="HTML",
        )
        return

    # Validate size
    if doc.file_size and doc.file_size > MAX_FILE_BYTES:
        await message.answer(
            f"⚠️ Файл слишком большой. Максимальный размер — {config.MAX_FILE_SIZE_MB} МБ."
        )
        return

    # Persist to disk
    config.ensure_upload_dir()
    save_dir = config.UPLOAD_DIR / str(user_id)
    save_dir.mkdir(parents=True, exist_ok=True)

    # Use original filename; make it unique if duplicate
    dest = save_dir / _unique_name(save_dir, file_name)

    try:
        tg_file = await bot.get_file(doc.file_id)
        await bot.download_file(tg_file.file_path, destination=str(dest))  # type: ignore[arg-type]
    except Exception as exc:
        logger.error("Download failed for %s (user=%d): %s", file_name, user_id, exc)
        await message.answer("❌ Не удалось загрузить файл. Попробуйте ещё раз.")
        return

    _sessions.setdefault(user_id, []).append(dest)
    count = len(_sessions[user_id])

    logger.info("Saved file: %s (user=%d, session_count=%d)", dest.name, user_id, count)

    confirm = await message.answer(
        f"Файл загружен: <b>{file_name}</b> (в очереди: {count})",
        parse_mode="HTML",
    )
    _confirm_message_ids.setdefault(user_id, []).append(confirm.message_id)


# ---------------------------------------------------------------------------
# Finish handler (ReplyKeyboard button sends a text message)
# ---------------------------------------------------------------------------

@router.message(F.text == FINISH_TEXT)
async def handle_finish(message: Message, bot: Bot) -> None:
    user_id = message.from_user.id  # type: ignore[union-attr]
    files = _sessions.pop(user_id, [])
    confirm_ids = _confirm_message_ids.pop(user_id, [])

    if not files:
        await message.answer(
            "⚠️ Очередь пуста. Сначала отправьте файлы <code>.docx</code>.",
            parse_mode="HTML",
        )
        return

    status_msg = await message.answer(
        f"⏳ Обрабатываю <b>{len(files)}</b> файл(ов)…",
        parse_mode="HTML",
    )

    # Parse all resumes (CPU-bound, but small enough for inline execution)
    resumes = []
    failed: list[str] = []
    for path in files:
        try:
            resume = await asyncio.get_event_loop().run_in_executor(
                None, extract_resume, path
            )
            resumes.append(resume)
        except Exception as exc:
            logger.error("Extraction failed for %s: %s", path.name, exc)
            failed.append(path.name)

    if not resumes:
        await status_msg.edit_text("❌ Не удалось обработать ни один из файлов.")
        _cleanup_user_dir(user_id)
        return

    # Build Excel with timestamped filename
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
    xlsx_name = f"candidates_{ts}.xlsx"
    output_path = config.UPLOAD_DIR / str(user_id) / xlsx_name
    try:
        await asyncio.get_event_loop().run_in_executor(
            None, build_excel, resumes, output_path
        )
    except Exception as exc:
        logger.error("Excel export failed (user=%d): %s", user_id, exc)
        await status_msg.edit_text("❌ Не удалось сформировать Excel-файл. Попробуйте ещё раз.")
        _cleanup_user_dir(user_id)
        return

    # Send file
    try:
        with open(output_path, "rb") as fh:
            xlsx_bytes = fh.read()

        caption_lines = [f"📊 <b>{xlsx_name}</b> — обработано резюме: {len(resumes)}."]
        if failed:
            caption_lines.append(f"\n⚠️ Не удалось обработать: {', '.join(failed)}")

        await bot.send_document(
            chat_id=message.chat.id,
            document=BufferedInputFile(xlsx_bytes, filename=xlsx_name),
            caption="\n".join(caption_lines),
            parse_mode="HTML",
        )
        await status_msg.delete()
        await _delete_confirm_messages(bot, message.chat.id, confirm_ids)
    except Exception as exc:
        logger.error("Send document failed (user=%d): %s", user_id, exc)
        await status_msg.edit_text("❌ Не удалось отправить файл. Попробуйте ещё раз.")
    finally:
        _cleanup_user_dir(user_id)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _unique_name(directory: Path, filename: str) -> str:
    """Append a counter suffix if a file with the same name already exists."""
    stem = Path(filename).stem
    suffix = Path(filename).suffix
    candidate = filename
    counter = 1
    while (directory / candidate).exists():
        candidate = f"{stem}_{counter}{suffix}"
        counter += 1
    return candidate


async def _delete_confirm_messages(bot: Bot, chat_id: int, message_ids: list[int]) -> None:
    """Silently delete all 'Файл загружен' confirmation messages."""
    for msg_id in message_ids:
        try:
            await bot.delete_message(chat_id=chat_id, message_id=msg_id)
        except Exception:
            pass  # already deleted or not accessible — ignore


def _cleanup_user_dir(user_id: int) -> None:
    user_dir = config.UPLOAD_DIR / str(user_id)
    try:
        shutil.rmtree(user_dir, ignore_errors=True)
        logger.info("Cleaned up temp dir for user=%d", user_id)
    except Exception as exc:
        logger.warning("Cleanup failed for user=%d: %s", user_id, exc)
