from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from core.logging import get_logger
from services.docx_parser import ParsedDocument, TextBlock, parse_docx
from services.pdf_parser import parse_pdf
from services.mhtml_parser import parse_mhtml

logger = get_logger(__name__)

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class ResumeData:
    name: str = ""
    phone: str = ""
    age: str = ""
    positions: str = ""
    source: str = ""
    source_file: str = ""
    parsed_at: str = ""


# ---------------------------------------------------------------------------
# Section header sets (Ukrainian)
# ---------------------------------------------------------------------------

_SECTION_HEADERS: frozenset[str] = frozenset({
    "контактна інформація",
    "досвід роботи",
    "освіта",
    "додаткова інформація",
    "додаткова освіта та сертифікати",
    "знання і навички",
    "знання мов",
    "бажані способи зв'язку",
    "бажані способи звʼязку",
    "сервіс пошуку роботи №1 в україні",
    "languages",
    "skills",
    "education",
    "experience",
    "work experience",
    "about",
    "contacts",
    "summary",
    "profile",
})

# Labels that appear before the value with a tab or colon
_PHONE_LABELS = re.compile(
    r"(?:телефон|phone|мобільний|tel\.?)\s*:?\t?(.+)",
    re.IGNORECASE,
)
_AGE_LABELS = re.compile(
    r"(?:вік|age)\s*:?\t?\s*(.+)",
    re.IGNORECASE,
)

# Explicit desired/target position labels (all supported languages).
# Matched only in the pre-experience portion of the document.
_POSITIONS_LABELS = re.compile(
    r"(?:"
    r"розглядає посади"
    r"|бажана посада"
    r"|бажані посади"
    r"|посада"
    r"|посади"
    r"|цільова посада"
    r"|желаемая должность"
    r"|цель"
    r"|desired position"
    r"|target position"
    r")\s*:?\t?(.+)",
    re.IGNORECASE,
)

# Section headers that mark the start of work experience.
# Once any of these is encountered, position extraction stops.
_EXPERIENCE_SECTION_HEADERS: frozenset[str] = frozenset({
    "досвід роботи",
    "опыт работы",
    "experience",
    "work experience",
    "employment history",
    "work history",
    "career history",
})

# Age value patterns
_AGE_VALUE = re.compile(r"(\d{1,3})\s*(?:\xa0|\s)*(?:рік|роки|років|years?|год)", re.IGNORECASE)
_AGE_BARE = re.compile(r"^(\d{1,3})$")

# Ukrainian phone normalization
_PHONE_RAW = re.compile(
    r"(?:\+?380|0)[\s\-]?(\d{2})[\s\-]?(\d{3})[\s\-]?(\d{2})[\s\-]?(\d{2})"
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _detect_source(path: Path) -> str:
    """Infer the job-platform source from the filename and extension."""
    name_lower = path.name.lower()
    suffix = path.suffix.lower()
    if "robota" in name_lower or suffix in (".mhtml", ".mht"):
        return "robota.ua"
    if "work" in name_lower:
        return "work.ua"
    if "hh" in name_lower:
        return "hh.ua"
    if "linkedin" in name_lower:
        return "LinkedIn"
    return ""


def _normalize_phone(raw: str) -> str:
    """Normalize any Ukrainian phone number to a plain digit string 380XXXXXXXXX."""
    cleaned = re.sub(r"[\s\-\(\)\+]", "", raw.replace("\xa0", ""))
    m = re.search(r"(?:380|0)(\d{9})", cleaned)
    if m:
        digits = m.group(1)  # 9 digits after country/leading 0
        return f"380{digits}"

    # Return stripped original if we cannot normalize
    return raw.strip()



def _extract_name_from_block(text: str) -> str:
    """
    work.ua format: 'Резюме від DD місяць YYYY\nПрізвище Ім'я По-батькові'
    Returns the name line (second line).
    """
    if "резюме від" in text.lower():
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        # Name is the line after the date line
        for i, line in enumerate(lines):
            if "резюме від" in line.lower() and i + 1 < len(lines):
                candidate = lines[i + 1]
                # Must look like a name: 2–4 capitalized Cyrillic/Latin words
                if _looks_like_name(candidate):
                    return candidate
    return ""


def _looks_like_name(text: str) -> bool:
    """Heuristic: 2–4 words, each starting with uppercase, mostly letters."""
    words = text.strip().split()
    if not (2 <= len(words) <= 4):
        return False
    return all(re.match(r"^[А-ЯЁЇІЄA-Z][а-яёїієa-z''\-A-Za-z]+$", w) for w in words)


def _extract_pdf_position_after_name(name: str, all_lines: list[str]) -> str:
    """
    robota.ua PDF layout: the position title is always the line immediately
    after the candidate's name.  Return that line if it passes sanity checks,
    otherwise return an empty string.
    """
    if not name:
        return ""
    try:
        idx = all_lines.index(name)
    except ValueError:
        return ""

    if idx + 1 >= len(all_lines):
        return ""

    candidate = all_lines[idx + 1].strip()
    if not candidate:
        return ""

    # Reject phone numbers
    cleaned = re.sub(r"[\s\-\(\)]", "", candidate.replace("\xa0", ""))
    if re.search(r"(?:\+?380|0)\d{9}", cleaned):
        return ""
    # Reject URLs / emails / viber links
    if re.search(r"(?:https?://|viber://|t\.me/|www\.|@)", candidate, re.IGNORECASE):
        return ""
    # Reject salary lines
    if re.search(r"\d.*грн|грн.*\d", candidate, re.IGNORECASE):
        return ""
    # Reject year/date-only lines
    if re.match(r"^\d{1,2}\s+\w+\s+\d{4}", candidate):
        return ""
    # Reject obvious service headers
    skip_prefixes = ("кандидат з", "резюме від", "сервіс", "http", "www")
    if any(candidate.lower().startswith(p) for p in skip_prefixes):
        return ""

    return candidate


def _collect_positions(blocks: list[TextBlock]) -> list[str]:
    """
    Collect ONLY the candidate's target / desired / current position.

    Priority sources, evaluated in document order, before the experience section:
      1. Heading 1 – resume title (highest priority).
      2. Explicit desired-position label line (Бажана посада, Посада, Цель, …).
      3. First non-section Heading 2 – professional headline near candidate name.

    Scanning stops as soon as a work-experience section header is detected.
    Job titles found inside Досвід роботи / Опыт работы / Experience are
    never included in the result.
    """
    positions: list[str] = []
    # Once True, skip the rest of the document for position extraction.
    in_experience_section = False
    # Guard: capture the headline (Heading 1 / Heading 2 title) only once.
    headline_captured = False

    for block in blocks:
        text = block.text.strip()
        lower = text.lower()
        style = block.style

        # ── Section-header detection ─────────────────────────────────────────
        if style in ("Heading 1", "Heading 2", "Heading 3"):
            if lower in _EXPERIENCE_SECTION_HEADERS:
                in_experience_section = True
                continue
            # Any other known section header: just skip the header line itself.
            if lower in _SECTION_HEADERS:
                continue

        # ── Freeze once inside the experience section ─────────────────────────
        if in_experience_section:
            continue

        # ── Source 1: Heading 1 resume title ─────────────────────────────────
        if style == "Heading 1":
            if text not in positions:
                positions.insert(0, text)
            headline_captured = True
            continue

        # ── Source 3: first non-section Heading 2 (professional headline) ────
        if style == "Heading 2" and lower not in _SECTION_HEADERS and not headline_captured:
            if text not in positions:
                positions.insert(0, text)
            headline_captured = True
            continue

        # ── Source 2: explicit desired-position label ─────────────────────────
        m = _POSITIONS_LABELS.match(text)
        if m:
            for pos in re.split(r"[,;/]", m.group(1)):
                p = pos.strip()
                if p and p not in positions:
                    positions.append(p)

    return positions


# ---------------------------------------------------------------------------
# Main extractor
# ---------------------------------------------------------------------------

def _read_mhtml_position(blocks: list[TextBlock]) -> str:
    """
    If the first block has style Heading 1 and contains two lines,
    the second line is the position title injected by the MHTML parser.
    """
    if not blocks:
        return ""
    first = blocks[0]
    if first.style != "Heading 1":
        return ""
    lines = [l.strip() for l in first.text.splitlines() if l.strip()]
    return lines[1] if len(lines) >= 2 else ""


def extract_resume(path: Path) -> ResumeData:
    """
    Parse a resume file (.docx / .pdf / .mhtml) and extract structured fields.
    Never raises — on any error returns a partially-filled ResumeData.
    """
    result = ResumeData(
        source=_detect_source(path),
        source_file=path.name,
        parsed_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )

    suffix = path.suffix.lower()
    try:
        if suffix == ".docx":
            doc = parse_docx(path)
        elif suffix == ".pdf":
            doc = parse_pdf(path)
        elif suffix in (".mhtml", ".mht"):
            doc = parse_mhtml(path)
        else:
            logger.error("Unsupported file format: %s", path.name)
            return result
    except Exception as exc:
        logger.error("Parsing failed for %s: %s", path.name, exc)
        return result

    blocks = doc.blocks
    all_lines = doc.all_lines

    # --- Name ---
    # For MHTML the mhtml_parser already put the name in blocks[0] (Heading 1).
    # For docx/pdf we use the work.ua "Резюме від" heuristic.
    for block in blocks:
        if block.style == "Heading 1":
            first_line = block.text.splitlines()[0].strip()
            if _looks_like_name(first_line):
                result.name = first_line
                break
        name = _extract_name_from_block(block.text)
        if name:
            result.name = name
            break

    # Fallback: first line that looks like a name (skip service lines)
    if not result.name:
        skip_prefixes = ("сервіс", "резюме", "http", "www")
        for line in all_lines:
            if any(line.lower().startswith(p) for p in skip_prefixes):
                continue
            if _looks_like_name(line):
                result.name = line
                break

    # --- Phone ---
    for line in all_lines:
        m = _PHONE_LABELS.match(line)
        if m:
            raw_phone = m.group(1).strip()
            result.phone = _normalize_phone(raw_phone)
            break

    if not result.phone:
        # Scan all lines; strip parens/spaces before matching so formats like
        # "+38 (098) 326-09-28" are found reliably.
        for line in all_lines:
            cleaned = re.sub(r"[\s\-\(\)]", "", line.replace("\xa0", ""))
            if re.search(r"(?:\+?380|0)\d{9}", cleaned):
                result.phone = _normalize_phone(line)
                break

    # --- Age ---
    for line in all_lines:
        m_label = _AGE_LABELS.match(line)
        if m_label:
            value_text = m_label.group(1).replace("\xa0", " ")
            m_val = _AGE_VALUE.search(value_text)
            if m_val:
                result.age = m_val.group(1)
                break
            m_bare = _AGE_BARE.match(value_text.strip())
            if m_bare:
                result.age = m_bare.group(1)
                break

    if not result.age:
        # Fallback: scan all lines for age patterns
        for line in all_lines:
            m = _AGE_VALUE.search(line.replace("\xa0", " "))
            if m:
                candidate = int(m.group(1))
                if 14 <= candidate <= 80:  # sanity check
                    result.age = str(candidate)
                    break

    # --- Positions ---
    # For MHTML the mhtml_parser stores the position as the second line of blocks[0].
    mhtml_position = _read_mhtml_position(blocks)
    if mhtml_position:
        result.positions = mhtml_position
    elif suffix == ".pdf" and result.name:
        # robota.ua PDFs: position is always the line immediately after the name.
        pdf_position = _extract_pdf_position_after_name(result.name, all_lines)
        if pdf_position:
            result.positions = pdf_position
        else:
            try:
                positions = _collect_positions(blocks)
                result.positions = "; ".join(positions)
            except Exception as exc:
                logger.warning("Position extraction failed for %s: %s", path.name, exc)
    else:
        try:
            positions = _collect_positions(blocks)
            result.positions = "; ".join(positions)
        except Exception as exc:
            logger.warning("Position extraction failed for %s: %s", path.name, exc)

    logger.info(
        "Extracted | file=%s name=%r phone=%r age=%r source=%r",
        path.name, result.name, result.phone, result.age, result.source,
    )
    return result
