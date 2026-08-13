from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from agent.intake import parse_sample_list


def _write_sample_list(path: Path, rows: list[list]) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    header = [
        "test_step_id",
        "sample_id",
        "identifying_details",
        "population_description",
        "population_size",
        "selection_method",
        "po_number",
        "branch",
    ]
    ws.append(header)
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


@pytest.fixture
def sample_list_xlsx(tmp_path: Path) -> Path:
    return _write_sample_list(
        tmp_path / "samples.xlsx",
        rows=[
            ["TS-4.2", "S01", "October accrual", "Monthly accruals FY2026", 12, "random", "", ""],
            ["TS-1.1", "S01", "PO 48213", "All POs > $5,000", 340, "random", "48213", "35210"],
            ["TS-1.1", "S02", "PO 48214", "All POs > $5,000", 340, "random", "48214", "35211"],
        ],
    )


def test_parses_one_manifest_per_test_step(sample_list_xlsx: Path):
    manifests = parse_sample_list(sample_list_xlsx)
    assert set(manifests) == {"TS-4.2", "TS-1.1"}


def test_sample_size_is_computed_not_read(sample_list_xlsx: Path):
    manifests = parse_sample_list(sample_list_xlsx)
    assert manifests["TS-1.1"].sample_size == 2
    assert manifests["TS-4.2"].sample_size == 1


def test_extra_columns_become_key_fields(sample_list_xlsx: Path):
    manifests = parse_sample_list(sample_list_xlsx)
    s01 = manifests["TS-1.1"].samples[0]
    assert s01.key_fields == {"po_number": "48213", "branch": "35210"}


def test_no_extra_columns_gives_none_key_fields(sample_list_xlsx: Path):
    manifests = parse_sample_list(sample_list_xlsx)
    accrual_sample = manifests["TS-4.2"].samples[0]
    assert accrual_sample.key_fields is None


def test_population_info_taken_from_first_row_per_step(sample_list_xlsx: Path):
    manifests = parse_sample_list(sample_list_xlsx)
    assert manifests["TS-1.1"].population_description == "All POs > $5,000"
    assert manifests["TS-1.1"].population_size == 340
    assert manifests["TS-1.1"].selection_method == "random"


def test_missing_required_column_raises(tmp_path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["test_step_id", "sample_id"])  # missing identifying_details etc.
    ws.append(["TS-4.2", "S01"])
    path = tmp_path / "bad.xlsx"
    wb.save(path)

    with pytest.raises(ValueError, match="missing required column"):
        parse_sample_list(path)


def test_invalid_selection_method_raises(tmp_path: Path):
    path = _write_sample_list(
        tmp_path / "bad_method.xlsx",
        rows=[["TS-4.2", "S01", "October accrual", "Monthly", 12, "vibes", "", ""]],
    )
    with pytest.raises(ValueError, match="selection_method"):
        parse_sample_list(path)


def test_blank_rows_are_skipped(tmp_path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(
        [
            "test_step_id",
            "sample_id",
            "identifying_details",
            "population_description",
            "population_size",
            "selection_method",
        ]
    )
    ws.append(["TS-4.2", "S01", "October accrual", "Monthly", 12, "random"])
    ws.append([None, None, None, None, None, None])
    path = tmp_path / "with_blank.xlsx"
    wb.save(path)

    manifests = parse_sample_list(path)
    assert manifests["TS-4.2"].sample_size == 1
