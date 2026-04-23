from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta
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
    publication_date: datetime | None = None
    name: str = ""
    phone: str = ""
    age: str = ""
    positions: str = ""
    work_experience: str = ""
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

# ---------------------------------------------------------------------------
# Work-experience extraction helpers
# ---------------------------------------------------------------------------

# robota.ua / MHTML date range: "MM.YYYY - MM.YYYY", "05.2025 - до теперішнього часу"
_DATE_RANGE_ROBOTA = re.compile(
    r"\d{2}\.\d{4}\s*[-–]\s*(?:\d{2}\.\d{4}|до теперішнього часу|наш час)",
    re.IGNORECASE,
)

# work.ua DOCX date lines: "з MM.YYYY по MM.YYYY (X років)" / "з MM.YYYY по нині ..."
_DATE_RANGE_WORKUA_LINE = re.compile(
    r"^з\s+(\d{2}\.\d{4})\s+по\s+(\S+)",
    re.IGNORECASE,
)

# Custom resume date ranges: "DD.MM.YYYY – DD.MM.YYYY" or "YYYY – YYYY"
_DATE_RANGE_CUSTOM = re.compile(
    r"(?:\d{2}\.\d{2}\.\d{4}|\d{4})\s*[-–]\s*(?:\d{2}\.\d{2}\.\d{4}|\d{4})",
)

# Duration-only lines, e.g. "2 роки 3 місяці", "8 місяців"
_DURATION_LINE = re.compile(
    r"^\d+\s+(?:рік|роки|років|місяць|місяці|місяців)",
    re.IGNORECASE,
)

# robota.ua trigger: "Працювала/Працював в 1 компанії …"
# Masculine: Працював, Feminine: Працювала, Neutral: Працювало, Plural: Працювали
_EXPERIENCE_TRIGGER = re.compile(
    r"Працюва(?:в|ла|ло|ли)\s+в\s+\d+",
    re.IGNORECASE,
)

# Lines that mark the end of the experience section in robota.ua pages
_ROBOTA_EXP_STOP: frozenset[str] = frozenset({
    "ключова інформація",
    "навчалась в",
    "навчався в",
    "навчувся в",
    "освіта",
    "написати в чат",
    "завантажити",
    "найсвіжішу версію резюме",
    "кандидат з найбільшої бази",
    "до переліку резюме",
    "володіє мовами",
})

# Lines that mark the end of the experience section in custom PDFs
_CUSTOM_EXP_STOP: frozenset[str] = frozenset({
    "освіта",
    "навички",
    "контактна інформація",
    "мови",
    "про себе",
    "про мене",
    "summary",
    "education",
    "skills",
})


def _is_exp_noise(line: str) -> bool:
    """Return True for lines that carry no job-entry information."""
    if not line or line == "⬥":
        return True
    cl = re.sub(r"&\w+;", " ", line).lower().strip()
    if any(cl.startswith(m) for m in _ROBOTA_EXP_STOP):
        return True
    if line.startswith("http") or "viber://" in line or "t.me/" in line:
        return True
    if _DURATION_LINE.match(line):
        return True
    if re.match(r"^\d+$", line):      # bare numeric IDs
        return True
    return False


def _format_job_entry(title: str, company: str, dates: str) -> str:
    parts = [p.strip() for p in [title, company] if p.strip()]
    text = ", ".join(parts)
    if dates.strip():
        text += f" ({dates.strip()})"
    return text


def _parse_jobs_from_exp_lines(
    exp_lines: list[str],
    date_re: re.Pattern,
) -> str:
    """
    Shared engine: given cleaned experience lines and a date-range regex,
    extract (title, company, dates) for every job using a window + dedup
    approach.

    For each date-range line, the window [prev_date+1 … di-1] is filtered for
    noise, then deduplicated (first occurrence kept).  After dedup the last two
    items are title and company.  If the last item contains '|' it is treated as
    a combined "title|company" entry.  If there are ≥ 3 items the first one is
    treated as the industry label and, when the company ends with it, the label
    is stripped from the company string.
    """
    date_indices = [i for i, l in enumerate(exp_lines) if date_re.search(l)]
    if not date_indices:
        return ""

    jobs: list[str] = []
    for idx, di in enumerate(date_indices):
        date_str = exp_lines[di].strip()
        prev_date = date_indices[idx - 1] if idx > 0 else -1

        # Collect non-noise lines in the window, preserving forward order
        raw: list[str] = []
        for j in range(prev_date + 1, di):
            l = exp_lines[j].strip()
            if not _is_exp_noise(l):
                raw.append(l)

        # Deduplicate keeping first occurrence (this removes MHTML industry repeats)
        seen: set[str] = set()
        deduped: list[str] = []
        for l in raw:
            if l not in seen:
                seen.add(l)
                deduped.append(l)

        if not deduped:
            continue

        last = deduped[-1]

        # "Title|Company" combined line (e.g. some custom resumes)
        if "|" in last:
            parts = last.split("|", 1)
            title, company = parts[0].strip(), parts[1].strip()
        elif len(deduped) == 1:
            title, company = last, ""
        else:
            title = deduped[-2]
            company = last
            # Strip trailing industry label from company (robota.ua PDF format)
            if len(deduped) >= 3:
                industry = deduped[-3]
                if company.endswith(industry):
                    company = company[: -len(industry)].strip()

        if title:
            jobs.append(_format_job_entry(title, company, date_str))

    return "\n".join(jobs)


def _extract_work_experience_from_blocks(blocks: list[TextBlock]) -> str:
    """
    Extract work experience from work.ua DOCX-style blocks.
    Uses paragraph heading styles (Heading 2 / Heading 3) as structural markers.
    """
    jobs: list[str] = []
    in_exp = False
    current_title = ""

    for block in blocks:
        text = block.text.strip()
        lower = text.lower()
        style = block.style

        if style == "Heading 2":
            if lower == "досвід роботи":
                in_exp = True
                continue
            if in_exp:
                break   # next Heading 2 ends the experience section

        if not in_exp:
            continue

        if style == "Heading 3":
            current_title = text
        elif style == "Normal" and current_title:
            lines = [l.strip() for l in text.splitlines() if l.strip()]
            date_line = lines[0] if lines else ""
            company = lines[1] if len(lines) > 1 else ""
            m = _DATE_RANGE_WORKUA_LINE.match(date_line)
            if m:
                start = m.group(1)
                end_raw = m.group(2).lower()
                end = "до теперішнього часу" if end_raw == "нині" else m.group(2)
                date_str = f"{start} - {end}"
            else:
                date_str = date_line
            jobs.append(_format_job_entry(current_title, company, date_str))
            current_title = ""

    return "\n".join(jobs)


def _extract_work_experience_robota(all_lines: list[str]) -> str:
    """
    Extract work experience from robota.ua all_lines (PDF or MHTML).
    Triggered by "Працював(а) в N компанії/компаніях …".
    """
    trigger_idx = None
    for i, line in enumerate(all_lines):
        clean = re.sub(r"&\w+;", " ", line).strip()
        if _EXPERIENCE_TRIGGER.search(clean):
            trigger_idx = i + 1
            break
    if trigger_idx is None:
        return ""

    exp_lines: list[str] = []
    for line in all_lines[trigger_idx:]:
        cl = re.sub(r"&\w+;", " ", line).lower().strip()
        if any(cl.startswith(m) for m in _ROBOTA_EXP_STOP):
            break
        exp_lines.append(line.strip())

    return _parse_jobs_from_exp_lines(exp_lines, _DATE_RANGE_ROBOTA)


def _extract_work_experience_custom(all_lines: list[str]) -> str:
    """
    Extract work experience from non-platform PDFs that use a plain
    'ДОСВІД РОБОТИ' section header.
    """
    trigger_idx = None
    for i, line in enumerate(all_lines):
        if line.strip().lower() in ("досвід роботи", "experience", "work experience"):
            trigger_idx = i + 1
            break
    if trigger_idx is None:
        return ""

    exp_lines: list[str] = []
    for line in all_lines[trigger_idx:]:
        cl = line.strip().lower()
        if any(cl.startswith(m) for m in _CUSTOM_EXP_STOP):
            break
        exp_lines.append(line.strip())

    return _parse_jobs_from_exp_lines(exp_lines, _DATE_RANGE_CUSTOM)


# Age value patterns
_AGE_VALUE = re.compile(r"(\d{1,3})\s*(?:\xa0|\s)*(?:рік|роки|років|years?|год)", re.IGNORECASE)
_AGE_BARE = re.compile(r"^(\d{1,3})$")

# Ukrainian phone normalization
_PHONE_RAW = re.compile(
    r"(?:\+?380|0)[\s\-]?(\d{2})[\s\-]?(\d{3})[\s\-]?(\d{2})[\s\-]?(\d{2})"
)

# ---------------------------------------------------------------------------
# Date parsing
# ---------------------------------------------------------------------------

_UA_MONTHS: dict[str, int] = {
    # genitive (used in "16 квітня 2026 року" and "Резюме від 24 грудня 2025")
    "січня": 1, "лютого": 2, "березня": 3, "квітня": 4,
    "травня": 5, "червня": 6, "липня": 7, "серпня": 8,
    "вересня": 9, "жовтня": 10, "листопада": 11, "грудня": 12,
    # nominative (fallback)
    "січень": 1, "лютий": 2, "березень": 3, "квітень": 4,
    "травень": 5, "червень": 6, "липень": 7, "серпень": 8,
    "вересень": 9, "жовтень": 10, "листопад": 11, "грудень": 12,
}

_UA_DATE_RE = re.compile(
    r"(\d{1,2})\s+"
    r"(січня|лютого|березня|квітня|травня|червня|липня|серпня|вересня|жовтня|листопада|грудня"
    r"|січень|лютий|березень|квітень|травень|червень|липень|серпень|вересень|жовтень|листопад|грудень)"
    r"\s+(\d{4})",
    re.IGNORECASE,
)

# Matches relative MHTML dates like "2 дні тому", "1 місяць тому", "13 годин тому"
_RELATIVE_DATE_RE = re.compile(
    r"(\d+)\s+(годин(?:и)?|година|день|дні|днів|тиждень|тижні|тижнів|місяць|місяці|місяців)\s+тому",
    re.IGNORECASE,
)


def _parse_ua_date(text: str) -> datetime | None:
    """Parse 'DD місяць YYYY' from any text string."""
    m = _UA_DATE_RE.search(text)
    if not m:
        return None
    try:
        day = int(m.group(1))
        month = _UA_MONTHS.get(m.group(2).lower())
        year = int(m.group(3))
        if not month:
            return None
        return datetime(year, month, day)
    except ValueError:
        return None


def _resolve_relative_date(raw: str, reference: datetime) -> datetime | None:
    """Convert a relative Ukrainian date string to an absolute datetime."""
    normalized = re.sub(r"\s+", " ", raw).strip()
    m = _RELATIVE_DATE_RE.match(normalized)
    if not m:
        return None
    n, unit = int(m.group(1)), m.group(2).lower()
    if unit.startswith("год") or unit == "година":
        delta = timedelta(hours=n)
    elif unit in ("день", "дні", "днів"):
        delta = timedelta(days=n)
    elif unit.startswith("тиж"):
        delta = timedelta(weeks=n)
    elif unit.startswith("місяц"):
        delta = timedelta(days=n * 30)
    else:
        return None
    result = reference - delta
    return result.replace(hour=0, minute=0, second=0, microsecond=0)


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

    # Try the international prefix first (380XXXXXXXXX).
    # This must come before the leading-zero fallback so that a line like
    # "13.03.2001380508527316" doesn't accidentally anchor on the "0" inside
    # "2001" and return a wrong number.
    m = re.search(r"380(\d{9})", cleaned)
    if m:
        return f"380{m.group(1)}"

    # Fallback: local format 0XXXXXXXXX — require a non-digit before the "0"
    # to avoid matching a "0" that is part of a year (e.g. "2001").
    m = re.search(r"(?<!\d)0(\d{9})(?!\d)", cleaned)
    if m:
        return f"380{m.group(1)}"

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
        # Scan all lines; strip formatting chars before matching.
        # Use the same two-step priority as _normalize_phone: prefer "+?380"
        # over a bare "0" so a date like "13.03.2001" on the same line as
        # "+380XXXXXXXXX" doesn't trigger a false positive anchor.
        for line in all_lines:
            cleaned = re.sub(r"[\s\-\(\)]", "", line.replace("\xa0", ""))
            if re.search(r"\+?380\d{9}", cleaned) or re.search(r"(?<!\d)0\d{9}(?!\d)", cleaned):
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

    # --- Work experience ---
    try:
        if suffix == ".docx":
            result.work_experience = _extract_work_experience_from_blocks(blocks)
        elif result.source == "robota.ua":
            result.work_experience = _extract_work_experience_robota(all_lines)
        else:
            result.work_experience = _extract_work_experience_custom(all_lines)
    except Exception as exc:
        logger.warning("Work experience extraction failed for %s: %s", path.name, exc)

    # --- Publication date ---
    if suffix in (".mhtml", ".mht") and doc.raw_date_hint:
        try:
            ref = datetime.fromtimestamp(path.stat().st_mtime)
        except OSError:
            ref = datetime.now()
        result.publication_date = _resolve_relative_date(doc.raw_date_hint, ref)
    else:
        # For DOCX and PDF: scan the first 15 lines for a Ukrainian date.
        # DOCX work.ua: "Резюме від 24 грудня 2025"
        # PDF robota.ua: "16 квітня 2026 року" (appears near the top)
        for line in all_lines[:15]:
            d = _parse_ua_date(line)
            if d:
                result.publication_date = d
                break

    exp_preview = result.work_experience.replace("\n", " | ")[:80] if result.work_experience else "—"
    logger.info(
        "Extracted | file=%s name=%r phone=%r age=%r source=%r date=%s exp=%r",
        path.name, result.name, result.phone, result.age, result.source,
        result.publication_date.strftime("%Y-%m-%d") if result.publication_date else "—",
        exp_preview,
    )
    return result
