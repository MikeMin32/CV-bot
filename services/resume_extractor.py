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
    "досвід",
    "зайнятість",
})

# Common section / sidebar headers that can appear alone on a line and
# must not be confused with job titles, company names, or candidate names.
# All entries are stored in lowercase for case-insensitive comparison.
_SIDE_HEADERS: frozenset[str] = frozenset({
    "особисті дані",
    "особиста інформація",
    "контакти",
    "контактна інформація",
    "контактні дані",
    "зайнятість",
    "освіта",
    "додаткова освіта",
    "додаткова освіта та сертифікати",
    "навички",
    "ключові навички",
    "професійні навички",
    "hard skills",
    "soft skills",
    "мови",
    "мова",
    "знання мов",
    "хобі",
    "інтереси",
    "професійні якості",
    "особисті якості",
    "сертифікати",
    "рекомендації",
    "додаткова інформація",
    "додаткова діяльність",
    "про себе",
    "про мене",
    "коротко про мене",
    "короткий зміст",
    "курси",
    "що я вмію",
    "резюме",
    "контакти:",
})

# All-caps section headers that should never be treated as candidate names.
_NOT_NAME_HEADERS: frozenset[str] = frozenset({
    "резюме",
    "досвід роботи",
    "досвід",
    "освіта",
    "навички",
    "про себе",
    "про мене",
    "контакти",
    "контактна інформація",
    "короткий зміст",
    "мова",
    "мови",
    "курси",
    "сертифікати",
    "хобі",
    "інтереси",
    "професійні якості",
    "особисті якості",
    "додаткова інформація",
    "додаткова діяльність",
    "додаткова освіта",
    "що я вмію",
    "ключова інформація",
    "ключові навички",
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

# Abbreviated Ukrainian month names used by many custom resume builders.
_UA_SHORT_MONTH = r"(?:січ|лют|бер|квіт|трав|черв|лип|серп|вер|жовт|лист|груд)"

# A single date endpoint: DD.MM.YYYY, MM.YYYY, YYYY, or "мес YYYY".
_DATE_POINT = (
    rf"(?:\d{{2}}\.\d{{2}}\.\d{{4}}"
    rf"|\d{{2}}\.\d{{4}}"
    rf"|\d{{4}}"
    rf"|{_UA_SHORT_MONTH}\s+\d{{4}})"
)

# Open-ended / ongoing markers that can appear on the right side of a range.
_DATE_POINT_OPEN = (
    rf"(?:{_DATE_POINT}"
    r"|Нинішній"
    r"|нині"
    r"|теперішній\s+час"
    r"|до\s+теперішнього\s+часу"
    r"|наш\s+час"
    r"|present"
    r"|now)"
)

# Custom resume date ranges used in many non-platform CVs.  Handles:
#   "DD.MM.YYYY – DD.MM.YYYY", "YYYY – YYYY",
#   "лют 2021 - серп 2024", "вер 2020 - лип 2024",
#   "квіт 2024 - Нинішній", "2021 - present".
_DATE_RANGE_CUSTOM = re.compile(
    rf"{_DATE_POINT}\s*[-–—]\s*{_DATE_POINT_OPEN}",
    re.IGNORECASE,
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

# Lines that mark the end of the experience section in custom PDFs.
# We stop scanning once we encounter one of these headers after the trigger.
_CUSTOM_EXP_STOP: frozenset[str] = frozenset({
    "освіта",
    "навички",
    "ключові навички",
    "професійні навички",
    "контактна інформація",
    "контакти",
    "мови",
    "мова",
    "знання мов",
    "про себе",
    "про мене",
    "summary",
    "education",
    "skills",
    "hard skills",
    "soft skills",
    "сертифікати",
    "рекомендації",
    "додаткова діяльність",
    "додаткова освіта",
    "курси",
    "хобі",
    "інтереси",
    "особисті якості",
    "професійні якості",
    "languages",
    "certificates",
    "additional information",
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
    if cl in _SIDE_HEADERS:
        return True
    return False


# Punctuation that typically ends a bullet/description line.
_DESC_ENDS = (".", "!", "?", ":", ";")
_DESC_STARTS = ("•", "●", "▪", "–", "—", "-", "·")


def _is_description_line(line: str) -> bool:
    """
    Heuristic for description / responsibility lines that should be ignored
    when collecting job-identifier candidates (title / company).
    """
    if not line:
        return True
    s = line.strip()
    if s.startswith(_DESC_STARTS):
        return True
    if s.endswith(_DESC_ENDS):
        return True
    return False


def _format_job_entry(title: str, company: str, dates: str) -> str:
    parts = [p.strip() for p in [title, company] if p.strip()]
    text = ", ".join(parts)
    if dates.strip():
        text += f" ({dates.strip()})"
    return text


def _collect_above(
    exp_lines: list[str],
    di: int,
    prev_di: int,
) -> list[str]:
    """
    Walk backwards from di-1 down to prev_di+1, collecting identifier
    candidates (non-noise, non-description lines).  Stops at the first
    description or bullet line so that descriptions from the *previous* job
    do not leak into this job's candidates.

    Returns the candidates in their original document order.
    """
    items: list[str] = []
    for j in range(di - 1, prev_di, -1):
        line = exp_lines[j].strip()
        if not line or _is_exp_noise(line):
            continue
        if _is_description_line(line):
            break
        items.append(line)
    items.reverse()
    return items


def _collect_below(
    exp_lines: list[str],
    di: int,
    next_di: int,
    limit: int = 2,
) -> list[str]:
    """
    Walk forward from di+1, collecting up to ``limit`` identifier candidates.
    Stops at the first description/bullet line so that descriptions of the
    *current* job do not leak into its identifier set.
    """
    items: list[str] = []
    for j in range(di + 1, next_di):
        line = exp_lines[j].strip()
        if not line or _is_exp_noise(line):
            continue
        if _is_description_line(line):
            break
        items.append(line)
        if len(items) >= limit:
            break
    return items


def _dedupe(items: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for it in items:
        if it and it not in seen:
            seen.add(it)
            out.append(it)
    return out


def _detect_job_format(
    exp_lines: list[str],
    date_indices: list[int],
) -> str:
    """
    Detect the structural ordering used by the document, based on the first
    date range encountered:

      • "before_complete" — identifiers appear BEFORE the date line
        (e.g. work.ua / robota.ua / Альбіна-style: title+company then date).
      • "company_after"   — only the title appears before the date;
        the company appears on the line directly after
        (e.g. some Canva CV templates: Title / Date / Company / description).
      • "date_first"      — date appears FIRST, then title, then company
        (e.g. Васенко-style templates).
    """
    if not date_indices:
        return "before_complete"

    first_di = date_indices[0]
    next_di = date_indices[1] if len(date_indices) > 1 else len(exp_lines)
    above = _collect_above(exp_lines, first_di, -1)
    below = _collect_below(exp_lines, first_di, next_di, limit=2)

    if len(above) >= 2:
        return "before_complete"
    if len(above) == 1 and below:
        return "company_after"
    if not above and len(below) >= 2:
        return "date_first"
    # Fallback: assume the classic "identifiers before date" layout.
    return "before_complete"


def _split_title_company_pipe(text: str) -> tuple[str, str]:
    """Split a 'Title|Company' combined line (some custom templates)."""
    parts = text.split("|", 1)
    return parts[0].strip(), parts[1].strip()


def _parse_jobs_from_exp_lines(
    exp_lines: list[str],
    date_re: re.Pattern,
) -> str:
    """
    Extract (title, company, dates) tuples for every job.  Supports three
    structural orderings — ``before_complete`` (title/company before date,
    as used by work.ua / robota.ua), ``company_after`` (title before date,
    company directly after), and ``date_first`` (date precedes title/company).
    The format is auto-detected from the first date range and then reused
    for the whole document.
    """
    date_indices = [i for i, l in enumerate(exp_lines) if date_re.search(l)]
    if not date_indices:
        return ""

    fmt = _detect_job_format(exp_lines, date_indices)
    jobs: list[str] = []

    for idx, di in enumerate(date_indices):
        date_str = exp_lines[di].strip()
        prev_di = date_indices[idx - 1] if idx > 0 else -1
        next_di = date_indices[idx + 1] if idx + 1 < len(date_indices) else len(exp_lines)

        above = _dedupe(_collect_above(exp_lines, di, prev_di))
        below = _collect_below(exp_lines, di, next_di, limit=2)

        title = ""
        company = ""

        if fmt == "company_after":
            # Title sits directly above the date; company sits directly below.
            if above:
                title = above[-1]
            if below:
                company = below[0]
            # Rare fallback for the first job in templates where the very
            # first title also precedes only one line: use "|" split if present.
            if title and "|" in title and not company:
                title, company = _split_title_company_pipe(title)

        elif fmt == "date_first":
            if below:
                title = below[0]
            if len(below) >= 2:
                company = below[1]
            if title and "|" in title and not company:
                title, company = _split_title_company_pipe(title)

        else:  # before_complete
            if above:
                last = above[-1]
                if "|" in last:
                    title, company = _split_title_company_pipe(last)
                elif len(above) == 1:
                    title = last
                else:
                    title = above[-2]
                    company = last
                    # Strip the industry suffix repeated on the company line
                    # (robota.ua PDF quirk, e.g. "ТОВ «…» Роздрібна торгівля").
                    if len(above) >= 3:
                        industry = above[-3]
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


_CUSTOM_EXP_TRIGGERS: frozenset[str] = frozenset({
    "досвід роботи",
    "досвід",
    "опыт работы",
    "опыт",
    "experience",
    "work experience",
    "employment history",
    "work history",
    "career history",
    "зайнятість",
    "робочий досвід",
    "професійний досвід",
})


def _extract_work_experience_custom(all_lines: list[str]) -> str:
    """
    Extract work experience from non-platform PDFs that use a plain section
    header such as ``ДОСВІД РОБОТИ``, ``Досвід``, ``Зайнятість``, ``Experience``.
    """
    trigger_idx = None
    for i, line in enumerate(all_lines):
        cl = line.strip().lower().rstrip(":")
        if cl in _CUSTOM_EXP_TRIGGERS:
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

# Labels that precede a birth-date value (inline or on the next line).
_BIRTH_LABELS: tuple[str, ...] = (
    "дата народження",
    "день народження",
    "народився",
    "народжений",
    "народжена",
    "дата рождения",
    "date of birth",
    "birthdate",
    "d.o.b.",
    "dob",
)

# Matches DD.MM.YYYY / DD/MM/YYYY / DD-MM-YYYY (with strict 4-digit year).
_DOTTED_DATE_RE = re.compile(
    r"(?<!\d)(\d{1,2})[./\-](\d{1,2})[./\-](\d{4})(?!\d)"
)

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


def _parse_dotted_date(text: str) -> datetime | None:
    """Parse 'DD.MM.YYYY' (or '/'/'-' separators) from any text string."""
    m = _DOTTED_DATE_RE.search(text)
    if not m:
        return None
    try:
        day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if not (1 <= month <= 12 and 1 <= day <= 31):
            return None
        return datetime(year, month, day)
    except ValueError:
        return None


def _parse_any_date(text: str) -> datetime | None:
    """Parse either a Ukrainian-word date or a DD.MM.YYYY dotted date."""
    return _parse_ua_date(text) or _parse_dotted_date(text)


def _age_from_birth(birth: datetime, ref: datetime | None = None) -> int:
    """Compute the age in whole years from a birth date."""
    today = (ref or datetime.now()).date()
    b = birth.date()
    age = today.year - b.year
    if (today.month, today.day) < (b.month, b.day):
        age -= 1
    return age


def _extract_birth_date(all_lines: list[str]) -> datetime | None:
    """
    Return the candidate's birth date, if it can be located.

    Strategy:
      1. Explicit label ("Дата народження", "Date of birth", …) followed by
         a date — either inline on the same line or on the next non-empty line.
      2. Fallback: scan every line for a date whose year is at most
         ``current_year - 15`` (heuristic: working-age candidates are born
         before that cut-off, so any such date is almost certainly a DOB
         rather than a publication / education / work-experience date).

    Both Ukrainian-word dates ("17 жовтня 2005") and dotted dates
    ("13.03.2001" / "02/04/2003") are supported.
    """
    now_year = datetime.now().year
    max_birth_year = now_year - 15  # candidate must be ≥ 15 years old
    min_birth_year = now_year - 90  # sanity upper bound

    def _accept(d: datetime | None) -> datetime | None:
        if d and min_birth_year <= d.year <= max_birth_year:
            return d
        return None

    # 1. Label-based lookup (inline value first, then the following line).
    for i, line in enumerate(all_lines):
        low = line.strip().lower()
        if not any(low.startswith(lbl) for lbl in _BIRTH_LABELS):
            continue
        d = _accept(_parse_any_date(line))
        if d:
            return d
        # Look ahead a couple of lines for the date value.
        for j in range(i + 1, min(i + 3, len(all_lines))):
            d = _accept(_parse_any_date(all_lines[j]))
            if d:
                return d

    # 2. Fallback: first plausible birth-year date anywhere in the document.
    for line in all_lines:
        d = _accept(_parse_any_date(line))
        if d:
            return d
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


_NAME_WORD_MIXED = re.compile(r"^[А-ЯЁЇІЄA-Z][а-яёїієa-z''\-A-Za-z]+$")
# ALL-CAPS variant used by some Canva / custom PDF templates (e.g. "ВІКТОРІЯ").
_NAME_WORD_UPPER = re.compile(r"^[А-ЯЁЇІЄA-Z]{2,}(?:['ʼ\-][А-ЯЁЇІЄA-Z]+)*$")


def _looks_like_name(text: str) -> bool:
    """Heuristic: 2–4 words that look like Cyrillic/Latin name parts."""
    s = text.strip()
    if not s:
        return False
    if s.lower() in _NOT_NAME_HEADERS:
        return False
    words = s.split()
    if not (2 <= len(words) <= 4):
        return False
    for w in words:
        if not (_NAME_WORD_MIXED.match(w) or _NAME_WORD_UPPER.match(w)):
            return False
    return True


def _normalize_name(name: str) -> str:
    """Convert ALL-CAPS names to Title Case; leave mixed-case names untouched."""
    s = name.strip()
    if not s:
        return s
    words = s.split()
    if all(w.isupper() for w in words if any(c.isalpha() for c in w)):
        return " ".join(w.capitalize() for w in words)
    return s


_POSITION_SKIP_PREFIXES: tuple[str, ...] = (
    "кандидат з", "резюме від", "сервіс", "http", "www",
    "телефон", "email", "e-mail", "телеграм", "telegram", "контакт",
)


def _is_contact_or_service_line(line: str) -> bool:
    """True if the line is clearly contact info or a known service header."""
    s = line.strip()
    if not s:
        return True
    cl = s.lower()
    if cl in _SIDE_HEADERS or cl in _NOT_NAME_HEADERS:
        return True
    cleaned = re.sub(r"[\s\-\(\)]", "", s.replace("\xa0", ""))
    if re.search(r"(?:\+?380|0)\d{9}", cleaned):
        return True
    if re.search(r"(?:https?://|viber://|t\.me/|www\.|@)", s, re.IGNORECASE):
        return True
    if re.search(r"\d.*грн|грн.*\d", s, re.IGNORECASE):
        return True
    if _UA_DATE_RE.search(s):
        return True
    if re.match(r"^\d{1,2}\s+\w+\s+\d{4}", s):
        return True
    if re.search(r"\d{4}\s*р\.", s):
        return True
    if any(cl.startswith(p) for p in _POSITION_SKIP_PREFIXES):
        return True
    # Bare city markers like "м. Вінниця" / "м.Київ"
    if re.match(r"^м\.\s*[А-ЯЁЇІЄA-Z]", s):
        return True
    # Gender / marital status single-word lines
    if cl in ("жіночий", "чоловічий", "жіноча", "чоловіча",
             "неодружений", "неодружена", "одружений", "одружена",
             "неодружений/неодружена"):
        return True
    return False


def _join_position_lines(lines: list[str]) -> str:
    """Join wrapped position lines, collapsing whitespace and trailing punctuation."""
    text = " ".join(lines)
    text = re.sub(r"\s*,\s*", ", ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text.rstrip(",; ")


def _extract_pdf_position_from_blocks(name: str, blocks: list[TextBlock]) -> str:
    """
    Look up the candidate's name inside each text block and return the
    professional headline that follows it *within the same block*.

    Many modern CV templates put the name and the (possibly multi-line)
    position in the same paragraph.  We keep collecting lines as long as the
    next one starts with a lowercase letter — a reliable signal for a wrapped
    continuation of the same headline — or until we hit a contact/service
    line or another section header.
    """
    if not name:
        return ""
    for block in blocks:
        lines = [l.strip() for l in block.text.splitlines() if l.strip()]
        if name not in lines:
            continue
        idx = lines.index(name)
        after = lines[idx + 1:]
        # We only trust the FIRST block that contains the candidate's name —
        # later occurrences are usually repeats inside sidebar contact blocks
        # (e.g. under "ПІБ" / "Адреса" labels) and produce false positives.
        if not after:
            return ""
        first = after[0]
        if _is_contact_or_service_line(first):
            return ""
        collected = [first]
        for nxt in after[1:]:
            if not nxt or _is_contact_or_service_line(nxt):
                break
            # Wrapped continuation: starts with a lowercase Cyrillic/Latin letter.
            if nxt[:1].isalpha() and nxt[:1].islower():
                collected.append(nxt)
                if len(collected) >= 5:
                    break
                continue
            break
        return _join_position_lines(collected)
    return ""


def _extract_pdf_position_after_name(name: str, all_lines: list[str]) -> str:
    """
    Fallback for the flat-text case: return the line immediately after the
    candidate's name when it passes basic sanity checks.
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
    if not candidate or _is_contact_or_service_line(candidate):
        return ""
    return candidate


# ---------------------------------------------------------------------------
# Filename-based position fallback
# ---------------------------------------------------------------------------

_ROBOTA_FILENAME_RE = re.compile(
    r"^(?P<body>.+?)_id_\d+_robota(?:_ua)?$",
    re.IGNORECASE,
)


def _position_from_filename(filename: str, name: str) -> str:
    """
    Best-effort extraction of the resume headline from common filename
    conventions used by Ukrainian job boards and users exporting from them:

    • robota.ua:  ``<Position>_<Name>_id_<digits>_robota[_ua].pdf``
    • work.ua:    ``Резюме_—_<Position>,_<Name>.pdf``
      (position may contain commas — we treat the *last* comma-separated
       piece that looks like a person name as the candidate's name.)
    """
    if not filename:
        return ""
    stem = filename
    for ext in (".pdf", ".PDF", ".docx", ".DOCX", ".mhtml", ".mht"):
        if stem.endswith(ext):
            stem = stem[: -len(ext)]
            break

    m = _ROBOTA_FILENAME_RE.match(stem)
    if m:
        body = m.group("body")
        if name:
            # Strip the name words from the tail, one by one (order-insensitive).
            for word in reversed(name.split()):
                suffix = "_" + word
                if body.lower().endswith(suffix.lower()):
                    body = body[: -len(suffix)]
        return body.replace("_", " ").strip()

    # work.ua-style: anything after "Резюме — "
    if stem.lower().startswith("резюме"):
        body = re.sub(r"^резюме\s*[_\s]*[—–\-]\s*[_\s]*", "",
                      stem, count=1, flags=re.IGNORECASE)
        body = body.replace("_", " ").strip()
        parts = [p.strip() for p in body.split(",") if p.strip()]
        if len(parts) >= 2 and _looks_like_name(parts[-1]):
            return ", ".join(parts[:-1])
        return body

    return ""


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
    raw_name = ""
    for block in blocks:
        if block.style == "Heading 1":
            first_line = block.text.splitlines()[0].strip()
            if _looks_like_name(first_line):
                raw_name = first_line
                break
        name = _extract_name_from_block(block.text)
        if name:
            raw_name = name
            break

    # Fallback: first line that looks like a name (skip service lines)
    if not raw_name:
        skip_prefixes = ("сервіс", "резюме", "http", "www", "кандидат")
        for line in all_lines:
            if any(line.lower().startswith(p) for p in skip_prefixes):
                continue
            if _looks_like_name(line):
                raw_name = line
                break

    # Store the canonical (Title Case) form, but keep ``raw_name`` for lookups
    # so that ALL-CAPS resumes still match the exact line in the document.
    result.name = _normalize_name(raw_name)

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

    if not result.age:
        # Last-chance: derive age from a birth date in the document.
        # Example sources: "Дата народження\n2 квітня 2003 р." (Canva/work.ua
        # templates) or an inline DOB right after the name (robota.ua PDFs).
        birth = _extract_birth_date(all_lines)
        if birth is not None:
            age_years = _age_from_birth(birth)
            if 15 <= age_years <= 90:
                result.age = str(age_years)

    # --- Positions ---
    # For MHTML the mhtml_parser stores the position as the second line of blocks[0].
    mhtml_position = _read_mhtml_position(blocks)
    if mhtml_position:
        result.positions = mhtml_position
    elif suffix == ".pdf":
        pdf_position = ""
        # 1. Preferred: multi-line headline that sits in the same paragraph
        #    as the candidate's name (works for most modern CV templates).
        if raw_name:
            pdf_position = _extract_pdf_position_from_blocks(raw_name, blocks)
        # 2. Fallback: classic robota.ua layout — line directly after the name.
        if not pdf_position and raw_name:
            pdf_position = _extract_pdf_position_after_name(raw_name, all_lines)
        # 3. Last resort: recognisable filename patterns (robota.ua / work.ua).
        if not pdf_position:
            pdf_position = _position_from_filename(path.name, result.name)
        # 4. Heading-based fallback (kept for any PDFs that expose heading styles).
        if not pdf_position:
            try:
                positions = _collect_positions(blocks)
                pdf_position = "; ".join(positions)
            except Exception as exc:
                logger.warning("Position extraction failed for %s: %s", path.name, exc)
        result.positions = pdf_position
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
            # Classic robota.ua export (MHTML / legacy PDF) uses the
            # "Працював(а) в N компанії…" trigger.  Newer robota.ua PDF
            # variants don't — fall back to the generic custom parser so
            # the candidate still gets work-experience data.
            exp = _extract_work_experience_robota(all_lines)
            if not exp:
                exp = _extract_work_experience_custom(all_lines)
            result.work_experience = exp
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
        #
        # Guard: many CV templates also show the candidate's birth date near
        # the top (e.g. "17 жовтня 2005 р.").  We therefore reject any date
        # that is clearly too old to be a publication date.
        now_year = datetime.now().year
        for line in all_lines[:15]:
            d = _parse_ua_date(line)
            if d and (now_year - 3) <= d.year <= (now_year + 1):
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
