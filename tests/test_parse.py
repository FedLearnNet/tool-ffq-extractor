from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

from ffq_extractor.export import rows_to_dataframe, write_csv
from ffq_extractor.mapping import load_food_mapping
from ffq_extractor.parse import (
    OUTPUT_COLUMNS,
    ExtractError,
    extract_ffq,
    parse_a1,
    resolve_patient_id,
    resolve_questionnaire_date,
    _parse_portion,
)

ROOT = Path(__file__).resolve().parents[1]
MAPPING = ROOT / "data" / "Matching_nutritional_variables.xlsx"
BLANK = ROOT / "tests" / "fixtures" / "blank_ffq.xlsx"

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


def test_parse_a1():
    assert parse_a1("A1") == (0, 0)
    assert parse_a1("a3") == (2, 0)
    assert parse_a1("B3") == (2, 1)
    assert parse_a1("AA10") == (9, 26)


def test_parse_a1_invalid():
    with pytest.raises(ExtractError):
        parse_a1("3A")
    with pytest.raises(ExtractError):
        parse_a1("")


def test_resolve_patient_id_from_cell():
    assert resolve_patient_id("MAI_42", cell="A3") == "MAI_42"


def test_resolve_patient_id_missing():
    with pytest.raises(ExtractError, match="Patient ID not found in cell A3"):
        resolve_patient_id(None, cell="A3")
    with pytest.raises(ExtractError, match="Patient ID not found in cell A3"):
        resolve_patient_id("  ", cell="A3")


def test_resolve_questionnaire_date_missing():
    with pytest.raises(ExtractError, match="Questionnaire date not found in cell A1"):
        resolve_questionnaire_date(None, cell="A1")


@pytest.mark.parametrize(
    "label,expected",
    [
        (
            "Meat, fish, poultry - Medium serving (size deck of cards)",
            "Medium serving (size deck of cards)",
        ),
        ("Dairy products and fat", ""),
        ("The following on bread or vegetables (teaspoon):", "teaspoon"),
    ],
)
def test_parse_portion(label, expected):
    assert _parse_portion(label) == expected


def _mark_row(food: str, header_name: str):
    cells = [None] * 10
    cells[FREQ_HEADERS.index(header_name)] = "X"
    return (food, *cells)


def build_synthetic_ffq(path: Path, *, sheet_title: str = "Sheet1") -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append([datetime(2026, 9, 1)])
    ws.append(["Participant study number: ", None])
    ws.append(["MAI_TEST01"])
    ws.append(
        (
            "Meat, fish, poultry - Medium serving (size deck of cards)",
            *FREQ_HEADERS,
            "Comment",
        )
    )
    # marked 6+/D
    ws.append(_mark_row("Beef (roast / steak)", "6 or more times a day"))
    # Never marked → omit
    ws.append(_mark_row("Beef: stew", "Never"))
    # blank → omit
    ws.append(("Pork (roast / chops / escalopes)", *([None] * 10)))
    # multi-mark → error
    multi = [None] * 10
    multi[FREQ_HEADERS.index("Once a week")] = "X"
    multi[FREQ_HEADERS.index("Once a day")] = "X"
    ws.append(("Lamb (roast / chops / stew)", *multi))
    # odd mark → omit with warning
    odd = [None] * 10
    odd[FREQ_HEADERS.index("Once a month")] = "NA"
    ws.append(("Bacon / Ham", *odd))

    ws.append(
        (
            "Potatoes, rice and pasta - Medium serving = a cupful",
            *FREQ_HEADERS,
            "Comment",
        )
    )
    ws.append(_mark_row("Mashed potatoes", "Once a month"))

    # NA-mapped item
    ws.append(
        (
            "Soups, sauces and spreads",
            *FREQ_HEADERS,
            "Comment",
        )
    )
    ws.append(_mark_row("Pickles, chutney (tablespoon)", "Once a week"))

    wb.save(path)
    return path


def test_blank_ffq_requires_patient_id(tmp_path):
    from openpyxl import load_workbook

    out = tmp_path / "blank_no_id.xlsx"
    wb = load_workbook(BLANK)
    ws = wb.active
    ws["A3"] = None
    wb.save(out)
    wb.close()

    result = extract_ffq(out, mapping_path=MAPPING)
    assert not result.ok
    assert any("Patient ID not found in cell A3" in e for e in result.errors)


def test_blank_ffq_with_patient_in_a3(tmp_path):
    from openpyxl import load_workbook

    out = tmp_path / "blank_with_id.xlsx"
    wb = load_workbook(BLANK)
    ws = wb.active
    ws["A3"] = "MAI_BLANK"
    if ws["A1"].value is None:
        ws["A1"] = datetime(2026, 1, 1)
    wb.save(out)
    wb.close()

    result = extract_ffq(out, mapping_path=MAPPING)
    assert result.ok
    assert result.patient_id == "MAI_BLANK"
    assert result.rows == []
    df = rows_to_dataframe(result.rows)
    assert list(df.columns) == OUTPUT_COLUMNS


def test_blank_ffq_fixture_uses_a3_label_as_id():
    """Vendored blank has 'MAI_' in A3; single-cell mode treats that as the ID."""
    result = extract_ffq(BLANK, mapping_path=MAPPING)
    assert result.ok
    assert result.patient_id == "MAI_"
    assert result.rows == []


def test_missing_date_cell_errors(tmp_path):
    path = tmp_path / "no_date.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append([None])
    ws.append([None])
    ws.append(["MAI_1"])
    wb.save(path)

    result = extract_ffq(path, mapping_path=MAPPING)
    assert not result.ok
    assert any("Questionnaire date not found in cell A1" in e for e in result.errors)


def test_patient_id_from_custom_cell(tmp_path):
    path = tmp_path / "custom_id.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append([datetime(2026, 9, 1)])
    ws.append([None])
    ws.append([None, "MAI_CUSTOM"])
    wb.save(path)

    result = extract_ffq(path, mapping_path=MAPPING, patient_id_cell="B3")
    assert result.ok
    assert result.patient_id == "MAI_CUSTOM"


def test_first_sheet_default(tmp_path):
    xlsx = build_synthetic_ffq(tmp_path / "renamed.xlsx", sheet_title="FFQ Data")
    result = extract_ffq(xlsx, mapping_path=MAPPING, sheet_name="")
    assert result.patient_id == "MAI_TEST01"
    assert any(r.food_item == "Beef (roast / steak)" for r in result.rows)


def test_named_sheet_missing(tmp_path):
    xlsx = build_synthetic_ffq(tmp_path / "one.xlsx")
    with pytest.raises(ExtractError, match="Sheet 'Missing' not found"):
        extract_ffq(xlsx, mapping_path=MAPPING, sheet_name="Missing")


def test_synthetic_extraction(tmp_path):
    xlsx = build_synthetic_ffq(tmp_path / "synthetic.xlsx")
    result = extract_ffq(xlsx, mapping_path=MAPPING)

    assert any("Multiple frequency marks" in e for e in result.errors)
    assert result.omitted_never == 1
    assert result.omitted_odd_mark == 1
    assert result.omitted_na_mapping == 1

    # Despite multi-mark error, other valid rows are collected
    by_item = {r.food_item: r for r in result.rows}
    assert by_item["Beef (roast / steak)"].frequency_consumption == "6+/D"
    assert (
        by_item["Beef (roast / steak)"].portion_size
        == "Medium serving (size deck of cards)"
    )
    assert by_item["Beef (roast / steak)"].food_group == (
        "Meat/fish/eggs and substitutes"
    )
    assert by_item["Beef (roast / steak)"].food_subgroup == "Meat excluding offal"

    mashed = by_item["Mashed potatoes"]
    assert mashed.frequency_consumption == "1/M"
    assert mashed.portion_size == "Medium serving = a cupful"
    assert mashed.food_group == "Starchy foods"
    assert mashed.food_subgroup == "Potatoe"
    assert mashed.unique_patient_id == "MAI_TEST01"
    assert mashed.questionnaire_date == "2026-09-01"

    assert "Pickles, chutney (tablespoon)" not in by_item
    assert "Beef: stew" not in by_item


def test_write_csv_golden(tmp_path):
    xlsx = build_synthetic_ffq(tmp_path / "synthetic.xlsx")
    result = extract_ffq(xlsx, mapping_path=MAPPING)
    csv_path = write_csv(result.rows, tmp_path / "out.csv")
    df = pd.read_csv(csv_path)
    assert list(df.columns) == OUTPUT_COLUMNS
    assert set(df["Food item"]) == {"Beef (roast / steak)", "Mashed potatoes"}


def test_na_mapping_is_warning_not_error(tmp_path):
    xlsx = build_synthetic_ffq(tmp_path / "synthetic.xlsx")
    result = extract_ffq(xlsx, mapping_path=MAPPING)
    assert result.omitted_na_mapping == 1
    assert any("mapped to NA" in w for w in result.warnings)
    assert not any("mapped to NA" in e for e in result.errors)


def test_mapping_loaded_once_external():
    mapping = load_food_mapping(MAPPING)
    assert "Beef (roast / steak)" in mapping
