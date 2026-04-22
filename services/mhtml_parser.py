from __future__ import annotations

"""
MHTML parser specialized for robota.ua candidate pages.

robota.ua saves pages as multipart MIME (MHTML / .mhtml).  The HTML inside
uses Angular-generated class names that are stable enough to use as anchors:

  • Candidate name  — <… class="santa-typo-h3 santa-text-black-700">
                       … <span … santahighlighter …>NAME</span>
  • Position title  — first meaningful text line after the name element
  • Phone / age / city — extracted by regex from the plain text of the page
                         (same patterns as the docx extractor)

The section "Резюме, схожі на вибране" marks the beginning of "similar
candidates" noise — we discard everything after it.
"""

import email
import re
from pathlib import Path

from core.logging import get_logger
from services.docx_parser import ParsedDocument, TextBlock

logger = get_logger(__name__)

# ── robota.ua HTML anchors ────────────────────────────────────────────────────

# Matches the opening tag of the h3 element that wraps the candidate name.
_NAME_H3 = re.compile(
    r'class="santa-typo-h3 santa-text-black-700">'
    r'.*?santahighlighter[^>]*>([^<]+)<',
    re.DOTALL,
)

# SVG path data: lines starting with an uppercase path command followed by a digit
_SVG_PATH = re.compile(r"^[A-Z]\d+\.\d+")

# UI button labels to skip when looking for the position title
_UI_NOISE = frozenset({
    "написати в чат",
    "запропонувати вакансію",
    "завантажити",
    "додати нотатку",
    "роздрукувати",
    "додати в helper",
    "/span>",
})

# Separates the candidate section from the "similar resumes" section
_SIMILAR_SECTION = "Резюме, схожі на вибране"


def parse_mhtml(path: Path) -> ParsedDocument:
    """
    Parse a robota.ua MHTML file into a ParsedDocument.

    The ParsedDocument contains:
      • blocks[0]  — a synthetic block with the candidate's name (and optionally
                     position on the next line), mirroring the work.ua docx format
                     so the existing resume_extractor logic can reuse it.
      • remaining blocks — plain-text lines from the page body for phone / age /
                           city detection.
    """
    try:
        html = _extract_html(path)
    except Exception as exc:
        logger.error("MHTML read failed for %s: %s", path.name, exc)
        return ParsedDocument()

    if not html:
        return ParsedDocument()

    # Clip noise after the "similar resumes" separator
    cutoff = html.find(_SIMILAR_SECTION)
    main_html = html[:cutoff] if cutoff > 0 else html

    blocks: list[TextBlock] = []

    # ── Name & position from HTML structure ───────────────────────────────────
    name = ""
    position = ""

    name_m = _NAME_H3.search(html)  # search full html (name is before cutoff anyway)
    if name_m:
        name = name_m.group(1).strip()
        position = _extract_position_after(html, name_m.end())

    # Build a synthetic "header" block so _extract_name_from_block / fallback work
    header_lines = [name] if name else []
    if position:
        header_lines.append(position)
    if header_lines:
        blocks.append(TextBlock(text="\n".join(header_lines), style="Heading 1"))

    # ── Body text for phone / age / city ─────────────────────────────────────
    body_text = _strip_tags(main_html)
    for line in body_text.splitlines():
        line = line.strip()
        if line and len(line) > 2:
            blocks.append(TextBlock(text=line, style="Normal"))

    return ParsedDocument(blocks=blocks)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _extract_html(path: Path) -> str:
    """Read MHTML and return the first text/html part, decoded."""
    with open(path, "rb") as fh:
        msg = email.message_from_bytes(fh.read())
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            payload = part.get_payload(decode=True)
            return payload.decode("utf-8", errors="replace")  # type: ignore[union-attr]
    return ""


def _strip_tags(html: str) -> str:
    """Remove all HTML tags, returning plain text with newlines at block boundaries."""
    # Replace block-level tags with newlines before stripping
    html = re.sub(r"<(?:div|p|li|br|h[1-6]|tr|td)[^>]*>", "\n", html, flags=re.IGNORECASE)
    return re.sub(r"<[^>]+>", "", html)


def _extract_position_after(html: str, offset: int) -> str:
    """
    Return the first meaningful text line that follows the name element.

    Skips SVG path data, short fragments, and known UI button labels.
    """
    after = html[offset:]
    plain = _strip_tags(after)
    for line in plain.splitlines():
        line = line.strip()
        if not line or len(line) < 4:
            continue
        if line.startswith("<"):
            continue
        if _SVG_PATH.match(line):
            continue
        if line.lower() in _UI_NOISE:
            continue
        return line
    return ""
