from __future__ import annotations

import re
import subprocess
from pathlib import Path

from core.logging import get_logger
from services.docx_parser import ParsedDocument, TextBlock

logger = get_logger(__name__)

# Matches a single all-caps Cyrillic/Latin word (name part on its own line in some PDFs)
_ALL_CAPS_WORD = re.compile(r"^[А-ЯЁЇІЄA-Z]{2,}$")


def parse_pdf(path: Path) -> ParsedDocument:
    """
    Extract text from a PDF file and return a ParsedDocument.

    Tries pdftotext (poppler) first for accurate layout analysis;
    falls back to pdfminer.six for text-layer extraction.
    Image-based PDFs return an empty ParsedDocument.
    """
    raw = _extract_via_pdftotext(path) or _extract_via_pdfminer(path)

    if not raw or not raw.strip():
        logger.warning(
            "No text extracted from %s — likely an image-based PDF", path.name
        )
        return ParsedDocument()

    text = _preprocess(raw)

    blocks: list[TextBlock] = []
    for para in re.split(r"\n{2,}", text):
        para = para.strip()
        if para:
            blocks.append(TextBlock(text=para, style="Normal"))

    return ParsedDocument(blocks=blocks)


# ---------------------------------------------------------------------------
# Extraction backends
# ---------------------------------------------------------------------------

def _extract_via_pdftotext(path: Path) -> str:
    """Use pdftotext (poppler) for layout-aware text extraction."""
    try:
        result = subprocess.run(
            ["pdftotext", str(path), "-"],
            capture_output=True,
            timeout=30,
        )
        if result.returncode == 0:
            return result.stdout.decode("utf-8", errors="replace")
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError) as exc:
        logger.debug("pdftotext not available or failed: %s", exc)
    return ""


def _extract_via_pdfminer(path: Path) -> str:
    """Fallback: use pdfminer.six for text extraction."""
    try:
        from pdfminer.high_level import extract_text  # type: ignore[import-untyped]

        return extract_text(str(path)) or ""
    except Exception as exc:
        logger.error("pdfminer failed for %s: %s", path.name, exc)
        return ""


# ---------------------------------------------------------------------------
# Preprocessing
# ---------------------------------------------------------------------------

def _preprocess(text: str) -> str:
    """
    Fix common PDF text artefacts before splitting into blocks.

    • Joins consecutive single all-caps Cyrillic/Latin lines into one name line
      (e.g. some Canva exports: "АЛЬБІНА\\nВЕДМІДЬОВА" → "Альбіна Ведмідьова").
    """
    lines = text.splitlines()
    result: list[str] = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if _ALL_CAPS_WORD.match(line) and i + 1 < len(lines):
            nxt = lines[i + 1].strip()
            if _ALL_CAPS_WORD.match(nxt):
                result.append(f"{line.capitalize()} {nxt.capitalize()}")
                i += 2
                continue
        result.append(lines[i])
        i += 1
    return "\n".join(result)
