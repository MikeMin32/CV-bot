from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


class Config:
    BOT_TOKEN: str = os.environ["BOT_TOKEN"]
    UPLOAD_DIR: Path = Path(os.getenv("UPLOAD_DIR", "/tmp/uploads"))
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "INFO")
    MAX_FILE_SIZE_MB: int = int(os.getenv("MAX_FILE_SIZE_MB", "20"))

    @classmethod
    def ensure_upload_dir(cls) -> None:
        cls.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


config = Config()
