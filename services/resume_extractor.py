from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from core.logging import get_logger
from services.docx_parser import ParsedDocument, TextBlock, parse_docx

logger = get_logger(__name__)

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class ResumeData:
    name: str = ""
    phone: str = ""
    city: str = ""
    age: str = ""
    positions: str = ""
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

# Current residence city — "місто роботи" is explicitly excluded via negative lookahead.
_CURRENT_CITY_LABELS = re.compile(
    r"(?:місто\s+проживання|місто(?!\s+роботи)|city|location)\s*:?\t?(.+)",
    re.IGNORECASE,
)

# Target / desired work city.
# Covers both label-value ("Готовий працювати: Вінниця") and
# phrase ("Готовий працювати у Вінниці") forms in one pattern.
_TARGET_CITY_LABELS = re.compile(
    r"(?:"
    r"місто\s+роботи"
    r"|бажане\s+місто"
    r"|бажаний\s+регіон"
    r"|desired\s+(?:work\s+)?(?:city|location)"
    r"|work\s+(?:city|location)"
    r"|preferred\s+(?:city|location)"
    # "Готовий/Готова/Готові працювати", "до переїзду", "до роботи"
    r"|готов(?:ий|а|і)?\s+(?:до\s+)?(?:переїзд[уі]|роботи|працювати)"
    r")\s*:?\t?\s*(.+)",
    re.IGNORECASE,
)

# Fallback phrase form without a colon separator:
# "Готова/Готовий працювати у Вінниці", "Готова до переїзду у Харків".
_READY_TO_WORK_PHRASE = re.compile(
    r"готов(?:ий|а|і)?\s+(?:до\s+)?(?:переїзд[уі]|роботи|працювати)\s+(?:в|у)\s+(.+)",
    re.IGNORECASE,
)
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

# Known Ukrainian cities for fallback city detection
_UA_CITIES: frozenset[str] = frozenset({
    "київ", "kyiv", "харків", "одеса", "дніпро", "запоріжжя", "львів",
    "кривий ріг", "миколаїв", "маріуполь", "луганськ", "вінниця",
    "херсон", "полтава", "чернігів", "черкаси", "суми", "житомир",
    "рівне", "івано-франківськ", "тернопіль", "хмельницький", "ужгород",
    "луцьк", "чернівці", "кропивницький", "турбів", "бровари",
})

# Locative / dative / accusative → nominative for common Ukrainian cities.
# Used to normalize city names extracted from phrases like "у Вінниці".
_UA_CITY_LOCATIVE: dict[str, str] = {
    "вінниці": "Вінниця",   "вінницю": "Вінниця",
    "одесі": "Одеса",       "одесу": "Одеса",
    "полтаві": "Полтава",   "полтаву": "Полтава",
    "черкасах": "Черкаси",
    "сумах": "Суми",
    "чернівцях": "Чернівці",
    "харкові": "Харків",
    "дніпрі": "Дніпро",    "дніпрові": "Дніпро",
    "києві": "Київ",        "києву": "Київ",
    "львові": "Львів",
    "запоріжжі": "Запоріжжя",
    "херсоні": "Херсон",
    "миколаєві": "Миколаїв", "миколаєву": "Миколаїв",
    "луганську": "Луганськ",
    "маріуполі": "Маріуполь",
    "луцьку": "Луцьк",
    "рівному": "Рівне",     "рівні": "Рівне",
    "тернополі": "Тернопіль",
    "хмельницькому": "Хмельницький",
    "ужгороді": "Ужгород",
    "чернігові": "Чернігів", "чернігову": "Чернігів",
    "житомирі": "Житомир",
    "кропивницькому": "Кропивницький",
    "броварах": "Бровари",
    "турбові": "Турбів",
    "кривому розі": "Кривий Ріг",
    "івано-франківську": "Івано-Франківськ",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize_phone(raw: str) -> str:
    """Normalize any Ukrainian phone number to +38 0XX XXX-XX-XX format."""
    m = _PHONE_RAW.search(raw.replace("\xa0", "").replace(" ", "").replace("-", ""))
    if m:
        # Groups already captured without separators — re-parse from cleaned string
        pass

    # Try again on the original with relaxed grouping
    cleaned = re.sub(r"[\s\-\(\)]", "", raw.replace("\xa0", ""))
    m2 = re.search(r"(?:\+?380|0)(\d{9})", cleaned)
    if m2:
        digits = m2.group(1)  # 9 digits after country/leading 0
        return f"+38 0{digits[:2]} {digits[2:5]}-{digits[5:7]}-{digits[7:9]}"

    # Return stripped original if we cannot normalize
    return raw.strip()


def _normalize_city_name(raw: str) -> str:
    """
    Normalize a city name to nominative form.

    Handles:
    - Leading prepositions: "у Вінниці" → "Вінниці" → "Вінниця"
    - Locative/dative/accusative inflections via _UA_CITY_LOCATIVE lookup
    - Already-nominative names present in _UA_CITIES
    - Unknown cities returned as-is (cleaned)
    """
    cleaned = raw.strip().strip(",.;:")
    # Strip leading Ukrainian prepositions "у"/"в" before the city name.
    cleaned = re.sub(r"^(?:у|в)\s+", "", cleaned, flags=re.IGNORECASE).strip()
    lower = cleaned.lower()
    if lower in _UA_CITY_LOCATIVE:
        return _UA_CITY_LOCATIVE[lower]
    if lower in _UA_CITIES:
        return cleaned.capitalize()
    return cleaned


def _format_city_field(current_city: str, target_city: str) -> str:
    """
    Compose the final city field from residence and target work location.

    target_city is the primary (job-target) value.
    current_city is shown in brackets as secondary context.

      Both present, different → "TargetCity (CurrentCity)"
      Both present, same      → "TargetCity"
      Only one present        → that value alone
    """
    current = current_city.strip()
    target = target_city.strip()
    if not current and not target:
        return ""
    if not target:
        return current
    if not current or current.lower() == target.lower():
        return target
    return f"{target} ({current})"


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

def extract_resume(path: Path) -> ResumeData:
    """
    Parse a .docx file and extract structured resume fields.
    Never raises — on any error returns a partially-filled ResumeData.
    """
    result = ResumeData(
        source_file=path.name,
        parsed_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )

    try:
        doc = parse_docx(path)
    except Exception as exc:
        logger.error("parse_docx failed for %s: %s", path.name, exc)
        return result

    blocks = doc.blocks
    all_lines = doc.all_lines

    # --- Name ---
    for block in blocks:
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
        # Scan all lines for a phone-like pattern
        for line in all_lines:
            if _PHONE_RAW.search(line.replace("\xa0", "")):
                result.phone = _normalize_phone(line)
                break

    # --- City ---
    current_city = ""
    target_city = ""

    for line in all_lines:
        if not current_city:
            m = _CURRENT_CITY_LABELS.match(line)
            if m:
                current_city = m.group(1).strip()
                logger.debug("City | current matched line=%r → %r", line, current_city)
        if not target_city:
            m = _TARGET_CITY_LABELS.match(line)
            if m:
                target_city = _normalize_city_name(m.group(1))
                logger.debug("City | target matched line=%r → %r", line, target_city)
        if not target_city:
            m = _READY_TO_WORK_PHRASE.match(line)
            if m:
                target_city = _normalize_city_name(m.group(1))
                logger.debug("City | ready-to-work phrase line=%r → %r", line, target_city)

    # Fallback: scan for a known city name to use as current residence.
    if not current_city:
        for line in all_lines:
            lower = line.lower()
            for city in _UA_CITIES:
                if city in lower:
                    current_city = city.capitalize()
                    logger.debug("City | fallback scan line=%r → %r", line, current_city)
                    break
            if current_city:
                break

    logger.debug(
        "City | file=%s  current_city=%r  target_city=%r",
        path.name, current_city, target_city,
    )
    result.city = _format_city_field(current_city, target_city)

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
    try:
        positions = _collect_positions(blocks)
        result.positions = "; ".join(positions)
    except Exception as exc:
        logger.warning("Position extraction failed for %s: %s", path.name, exc)

    logger.info(
        "Extracted | file=%s name=%r phone=%r city=%r age=%r",
        path.name, result.name, result.phone, result.city, result.age,
    )
    return result
