"""Parse FFQ Excel matrix into Nutrition rows."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple, Union

from openpyxl import load_workbook

from .frequencies import (
    frequency_code_for_header,
    is_accepted_mark,
    is_blank_cell,
    normalize_header,
)
from .mapping import FoodMapping, load_food_mapping

OUTPUT_COLUMNS = [
    "Unique Patient ID",
    "Questionnaire date",
    "Portion size",
    "Frequency consumption",
    "Food group",
    "Food subgroup",
    "Food item",
]

FREQ_COL_START = 1  # column B (0-indexed)
FREQ_COL_END = 11  # exclusive → columns B–K

_A1_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


@dataclass
class NutritionRow:
    unique_patient_id: str
    questionnaire_date: str
    portion_size: str
    frequency_consumption: str
    food_group: str
    food_subgroup: str
    food_item: str

    def as_dict(self) -> Dict[str, str]:
        return {
            "Unique Patient ID": self.unique_patient_id,
            "Questionnaire date": self.questionnaire_date,
            "Portion size": self.portion_size,
            "Frequency consumption": self.frequency_consumption,
            "Food group": self.food_group,
            "Food subgroup": self.food_subgroup,
            "Food item": self.food_item,
        }


@dataclass
class ExtractResult:
    rows: List[NutritionRow] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    patient_id: str = ""
    questionnaire_date: str = ""
    marked_count: int = 0
    omitted_never: int = 0
    omitted_blank: int = 0
    omitted_na_mapping: int = 0
    omitted_odd_mark: int = 0

    @property
    def ok(self) -> bool:
        return not self.errors


class ExtractError(Exception):
    """Fatal FFQ extraction error."""


def parse_a1(address: str) -> Tuple[int, int]:
    """Parse an A1-style address into 0-based (row, col)."""
    text = (address or "").strip().upper()
    match = _A1_RE.match(text)
    if not match:
        raise ExtractError(
            f"Invalid cell address {address!r}. Expected A1-style (e.g. A1, B3)."
        )
    col_letters, row_digits = match.groups()
    col = 0
    for ch in col_letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    col -= 1
    row = int(row_digits) - 1
    if row < 0 or col < 0:
        raise ExtractError(f"Invalid cell address {address!r}.")
    return row, col


def get_cell(rows: Sequence[Sequence[object]], address: str) -> object:
    """Return the value at an A1 address, or None if out of bounds."""
    row_idx, col_idx = parse_a1(address)
    if row_idx >= len(rows):
        return None
    row = rows[row_idx]
    if col_idx >= len(row):
        return None
    return row[col_idx]


def resolve_sheet(workbook, sheet_name: Optional[str] = None):
    """Return the named sheet, or the first sheet when name is blank."""
    names = list(workbook.sheetnames)
    if not names:
        raise ExtractError("Workbook has no worksheets")
    name = (sheet_name or "").strip()
    if not name:
        return workbook[names[0]]
    if name not in workbook.sheetnames:
        raise ExtractError(
            f"Sheet {name!r} not found. Available: {workbook.sheetnames}"
        )
    return workbook[name]


def _format_date(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    return text


def _cell_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def resolve_patient_id(value: object, *, cell: str = "A3") -> str:
    """Require a non-blank patient ID in the configured cell."""
    text = _cell_text(value)
    if text:
        return text
    raise ExtractError(f"Patient ID not found in cell {cell}")


def resolve_questionnaire_date(value: object, *, cell: str = "A1") -> str:
    """Require a non-blank questionnaire date in the configured cell."""
    text = _format_date(value)
    if text:
        return text
    raise ExtractError(f"Questionnaire date not found in cell {cell}")


def _parse_portion(section_label: str) -> str:
    if " - " not in section_label:
        # e.g. "The following on bread or vegetables (teaspoon):"
        lowered = section_label.lower()
        if "teaspoon" in lowered:
            return "teaspoon"
        return ""
    return section_label.split(" - ", 1)[1].strip()


def _is_section_header(row: Sequence[object]) -> bool:
    if len(row) <= FREQ_COL_START:
        return False
    return normalize_header(row[FREQ_COL_START]) == "never"


def _extract_frequency_headers(row: Sequence[object]) -> List[object]:
    return list(row[FREQ_COL_START:FREQ_COL_END])


def _select_frequency(
    food_item: str,
    cells: Sequence[object],
    headers: Sequence[object],
    result: ExtractResult,
) -> Optional[str]:
    """Return schema frequency code or None if the row should be omitted."""
    marked_indexes: List[int] = []
    odd_marks: List[str] = []

    for idx, cell in enumerate(cells):
        if is_blank_cell(cell):
            continue
        if is_accepted_mark(cell):
            marked_indexes.append(idx)
        else:
            odd_marks.append(str(cell).strip())

    if len(marked_indexes) > 1:
        msg = (
            f"Multiple frequency marks for food item {food_item!r} "
            f"at columns {[headers[i] for i in marked_indexes]!r}"
        )
        result.errors.append(msg)
        return None

    if odd_marks and not marked_indexes:
        msg = (
            f"Unrecognized mark(s) {odd_marks!r} for food item {food_item!r}; "
            "row omitted"
        )
        result.warnings.append(msg)
        result.omitted_odd_mark += 1
        return None

    if odd_marks and marked_indexes:
        msg = (
            f"Unrecognized extra mark(s) {odd_marks!r} alongside a valid mark "
            f"for food item {food_item!r}"
        )
        result.warnings.append(msg)

    if not marked_indexes:
        result.omitted_blank += 1
        return None

    header = headers[marked_indexes[0]]
    if normalize_header(header) == "never":
        result.omitted_never += 1
        return None

    try:
        code = frequency_code_for_header(header)
    except ValueError as exc:
        result.errors.append(str(exc))
        return None

    if code is None:
        result.omitted_never += 1
        return None
    return code


def extract_ffq(
    workbook_path: Union[str, Path],
    mapping_path: Optional[Union[str, Path]] = None,
    *,
    sheet_name: str = "",
    date_cell: str = "A1",
    patient_id_cell: str = "A3",
    food_mapping: Optional[Dict[str, FoodMapping]] = None,
    read_only: bool = False,
) -> ExtractResult:
    """
    Read an FFQ workbook and return Nutrition rows + diagnostics.

    Use read_only=True for batch/ZIP workloads (lower memory).
    Blank sheet_name selects the first worksheet.
    """
    path = Path(workbook_path)
    mapping = food_mapping if food_mapping is not None else load_food_mapping(mapping_path)
    result = ExtractResult()

    # Validate cell addresses early so bad hyperparams fail clearly.
    parse_a1(date_cell)
    parse_a1(patient_id_cell)

    wb = load_workbook(path, data_only=True, read_only=read_only)
    try:
        ws = resolve_sheet(wb, sheet_name)
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    if len(rows) < 3:
        raise ExtractError("FFQ sheet has fewer than 3 rows")

    try:
        questionnaire_date = resolve_questionnaire_date(
            get_cell(rows, date_cell), cell=date_cell.strip().upper() or "A1"
        )
    except ExtractError as exc:
        result.errors.append(str(exc))
        return result
    result.questionnaire_date = questionnaire_date

    try:
        patient_id = resolve_patient_id(
            get_cell(rows, patient_id_cell),
            cell=patient_id_cell.strip().upper() or "A3",
        )
    except ExtractError as exc:
        result.errors.append(str(exc))
        return result
    result.patient_id = patient_id

    current_portion = ""
    current_headers: List[object] = []

    for row_idx, row in enumerate(rows[3:], start=4):
        if not row or row[0] is None or not str(row[0]).strip():
            continue
        label = str(row[0]).strip()

        if _is_section_header(row):
            current_portion = _parse_portion(label)
            current_headers = _extract_frequency_headers(row)
            continue

        if not current_headers:
            result.warnings.append(
                f"Row {row_idx}: food item {label!r} before any section header; skipped"
            )
            continue

        food = mapping.get(label)
        if food is None:
            result.errors.append(
                f"Row {row_idx}: food item {label!r} not found in mapping file"
            )
            continue

        cells = list(row[FREQ_COL_START:FREQ_COL_END])
        # Pad if short
        if len(cells) < len(current_headers):
            cells.extend([None] * (len(current_headers) - len(cells)))

        code = _select_frequency(label, cells, current_headers, result)
        if code is None:
            continue

        if food.is_na:
            msg = (
                f"Row {row_idx}: food item {label!r} mapped to NA "
                f"({food.comment or 'not in French questionnaire'}); omitted"
            )
            result.warnings.append(msg)
            result.omitted_na_mapping += 1
            continue

        result.marked_count += 1
        result.rows.append(
            NutritionRow(
                unique_patient_id=patient_id,
                questionnaire_date=questionnaire_date,
                portion_size=current_portion,
                frequency_consumption=code,
                food_group=food.food_group,
                food_subgroup=food.food_subgroup,
                food_item=food.food_item,
            )
        )

    return result
