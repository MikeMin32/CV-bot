from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from core.logging import get_logger
from services.resume_extractor import ResumeData

logger = get_logger(__name__)

COLUMNS: list[str] = [
    "Посада",
    "ПІБ",
    "Вік",
    "Джерело",
    "Номер телефону",
]

# Header styling
_HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
_HEADER_FILL = PatternFill(fill_type="solid", fgColor="2F5496")
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Minimum column widths (characters)
_MIN_COL_WIDTHS: dict[str, int] = {
    "Посада": 45,
    "ПІБ": 28,
    "Вік": 8,
    "Джерело": 16,
    "Номер телефону": 22,
}


def _auto_fit_column(ws, col_idx: int, header: str, values: list[str]) -> None:
    letter = get_column_letter(col_idx)
    max_len = max(
        [len(header)] + [len(str(v)) for v in values]
    )
    min_width = _MIN_COL_WIDTHS.get(header, 15)
    ws.column_dimensions[letter].width = max(min_width, min(max_len + 4, 60))


def build_excel(resumes: list[ResumeData], output_path: Path) -> None:
    """Write a list of ResumeData objects to an .xlsx file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Candidates"
    ws.freeze_panes = "A2"  # freeze header row

    # Write header
    for col_idx, header in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGN
    ws.row_dimensions[1].height = 22

    # Write data rows
    row_values: dict[str, list[str]] = {col: [] for col in COLUMNS}
    for row_idx, resume in enumerate(resumes, start=2):
        values = [
            resume.positions,
            resume.name,
            resume.age,
            resume.source,
            resume.phone,
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            row_values[COLUMNS[col_idx - 1]].append(str(value))

    # Auto-fit columns
    for col_idx, header in enumerate(COLUMNS, start=1):
        _auto_fit_column(ws, col_idx, header, row_values[header])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))
    logger.info("Excel saved: %s (%d rows)", output_path, len(resumes))
