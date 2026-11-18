import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

from ffq_extractor.batch import list_excel_members, process_input, process_zip
from ffq_extractor.parse import OUTPUT_COLUMNS

ROOT = Path(__file__).resolve().parents[1]
MAPPING = ROOT / "data" / "Matching_nutritional_variables.xlsx"

FREQ_HEADERS = [
    "Never",
    "Less than once a month",
    "Once a month",
    "Once a week",
    "Twice a week",
    "5 times a week",
    "Once a day",
    "Twice a day",
    "4 times a day",
    "6 or more times a day",
]


def _mark_row(food: str, header_name: str):
    cells = [None] * 10
    cells[FREQ_HEADERS.index(header_name)] = "X"
    return (food, *cells)


def build_marked_ffq(path: Path, patient_id: str, food: str, header: str) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append([datetime(2026, 9, 1)])
    ws.append(["Participant study number: ", None])
    ws.append([patient_id])
    ws.append(
        (
            "Meat, fish, poultry - Medium serving (size deck of cards)",
            *FREQ_HEADERS,
            "Comment",
        )
    )
    ws.append(_mark_row(food, header))
    wb.save(path)
    return path


def build_blank_patient_ffq(path: Path) -> Path:
    """FFQ without patient id → extraction error."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append([datetime(2026, 9, 1)])
    ws.append(["Participant study number: ", None])
    ws.append([None])
    ws.append(
        (
            "Meat, fish, poultry - Medium serving (size deck of cards)",
            *FREQ_HEADERS,
            "Comment",
        )
    )
    ws.append(_mark_row("Beef (roast / steak)", "Once a day"))
    wb.save(path)
    return path


def test_list_excel_members_skips_junk(tmp_path):
    zip_path = tmp_path / "mixed.zip"
    good = tmp_path / "good.xlsx"
    build_marked_ffq(good, "MAI_1", "Beef (roast / steak)", "Once a day")
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.write(good, arcname="patients/p1.xlsx")
        zf.writestr("__MACOSX/._p1.xlsx", b"junk")
        zf.writestr("notes.txt", "ignore me")
        zf.writestr(".hidden.xlsx", b"not a real workbook")
    members = list_excel_members(zip_path)
    assert members == ["patients/p1.xlsx"]


def test_process_zip_merges_and_logs_failures(tmp_path):
    a = build_marked_ffq(
        tmp_path / "a.xlsx", "MAI_A", "Beef (roast / steak)", "Once a day"
    )
    b = build_marked_ffq(
        tmp_path / "b.xlsx", "MAI_B", "Mashed potatoes", "Once a month"
    )
    bad = build_blank_patient_ffq(tmp_path / "bad.xlsx")

    zip_path = tmp_path / "batch.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.write(a, arcname="cohort/a.xlsx")
        zf.write(b, arcname="cohort/b.xlsx")
        zf.write(bad, arcname="cohort/bad.xlsx")
        zf.writestr("__MACOSX/._a.xlsx", b"x")

    out = tmp_path / "nutrition.csv"
    errors: list[tuple[str, str]] = []

    def _on_error(member: str, message: str) -> None:
        errors.append((member, message))

    result = process_zip(
        zip_path,
        out,
        mapping_path=MAPPING,
        workers=2,
        progress_every=1,
        error_callback=_on_error,
    )

    assert result.files_total == 3
    assert result.files_ok == 2
    assert result.files_failed == 1
    assert result.rows_written == 2
    assert any("bad.xlsx" in m for m in result.failed_members)
    assert any("Patient ID" in msg for _, msg in errors)
    assert not list(tmp_path.glob("*.errors.log"))
    assert not list(tmp_path.glob("*.part.csv"))

    df = pd.read_csv(out)
    assert list(df.columns) == OUTPUT_COLUMNS
    assert set(df["Unique Patient ID"]) == {"MAI_A", "MAI_B"}
    assert set(df["Food item"]) == {"Beef (roast / steak)", "Mashed potatoes"}


def test_process_input_single_xlsx(tmp_path):
    xlsx = build_marked_ffq(
        tmp_path / "one.xlsx", "MAI_ONE", "Beef (roast / steak)", "6 or more times a day"
    )
    out = tmp_path / "out.csv"
    result = process_input(xlsx, out, mapping_path=MAPPING)
    assert result.ok
    assert result.files_ok == 1
    assert result.rows_written == 1
    df = pd.read_csv(out)
    assert df.iloc[0]["Frequency consumption"] == "6+/D"
    assert not list(tmp_path.glob("*.errors.log"))


def test_process_input_single_xlsx_fail_fast(tmp_path):
    xlsx = build_blank_patient_ffq(tmp_path / "bad.xlsx")
    out = tmp_path / "out.csv"
    errors: list[tuple[str, str]] = []
    result = process_input(
        xlsx,
        out,
        mapping_path=MAPPING,
        error_callback=lambda m, msg: errors.append((m, msg)),
    )
    assert not result.ok
    assert result.files_failed == 1
    assert any("bad.xlsx" in m for m, _ in errors)
    assert any("Patient ID" in msg for _, msg in errors)
    assert not list(tmp_path.glob("*.errors.log"))


def test_process_zip_invalid_date_cell_reports_filename(tmp_path):
    a = build_marked_ffq(
        tmp_path / "a.xlsx", "MAI_A", "Beef (roast / steak)", "Once a day"
    )
    zip_path = tmp_path / "batch.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.write(a, arcname="cohort/a.xlsx")

    errors: list[tuple[str, str]] = []
    result = process_zip(
        zip_path,
        tmp_path / "nutrition.csv",
        mapping_path=MAPPING,
        date_cell="A4",
        patient_id_cell="A3",
        workers=1,
        error_callback=lambda m, msg: errors.append((m, msg)),
    )

    assert result.files_ok == 0
    assert result.files_failed == 1
    assert any("a.xlsx" in m for m in result.failed_members)
    assert len(errors) == 1
    member, message = errors[0]
    assert "a.xlsx" in member
    assert "not a valid date" in message
    assert "A4" in message


def test_read_only_parse_matches(tmp_path):
    from ffq_extractor.parse import extract_ffq

    xlsx = build_marked_ffq(
        tmp_path / "ro.xlsx", "MAI_RO", "Beef (roast / steak)", "Once a week"
    )
    a = extract_ffq(xlsx, mapping_path=MAPPING, read_only=False)
    b = extract_ffq(xlsx, mapping_path=MAPPING, read_only=True)
    assert a.ok and b.ok
    assert [r.as_dict() for r in a.rows] == [r.as_dict() for r in b.rows]
