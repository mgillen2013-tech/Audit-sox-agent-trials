"""Parses the CY sample list -- the intake input from design doc section 1
that previously had no code behind it (SamplePopulationManifest existed as
a schema, but nothing built it from an actual file).

Expected file: one Excel sheet, one row per sampled item.

Required columns (header row, matched case-insensitively, whitespace
trimmed):
    test_step_id          -- which test step this sample belongs to
    sample_id              -- "S01".."S25", unique within a test step
    identifying_details    -- free text describing the item
    population_description -- same value repeated on every row for a given
                               test_step_id, e.g. "All POs > $5,000 issued
                               Oct 2025-Sep 2026"
    selection_method        -- one of: random, haphazard, judgmental, all_items

Optional column:
    population_size         -- leave blank if unknown

Any other columns (e.g. po_number, branch, date, amount) are captured per
row as SampleItem.key_fields, so search_cy_support can be pointed at a
specific field instead of a fuzzy description match.

sample_size is NOT a column -- it's computed as the row count per
test_step_id, so it can never drift out of sync with the actual sample list
the way a manually-typed count could.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl

from agent.schemas import SampleItem, SamplePopulationManifest

_REQUIRED_COLUMNS = {
    "test_step_id",
    "sample_id",
    "identifying_details",
    "population_description",
    "selection_method",
}
_KNOWN_COLUMNS = _REQUIRED_COLUMNS | {"population_size"}


def parse_sample_list(path: str | Path) -> dict[str, SamplePopulationManifest]:
    """Returns {test_step_id: SamplePopulationManifest}."""
    path = Path(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows())
    if not rows:
        raise ValueError(f"{path.name}: sheet is empty")

    header_row = rows[0]
    col_index: dict[str, int] = {}
    for idx, cell in enumerate(header_row):
        if cell.value is None:
            continue
        col_index[str(cell.value).strip().lower()] = idx

    missing = _REQUIRED_COLUMNS - set(col_index)
    if missing:
        raise ValueError(
            f"{path.name}: missing required column(s) {sorted(missing)}. "
            f"Found columns: {sorted(col_index)}"
        )

    key_field_columns = {name: idx for name, idx in col_index.items() if name not in _KNOWN_COLUMNS}

    def cell_str(row, col_name: str) -> str | None:
        idx = col_index.get(col_name)
        if idx is None or idx >= len(row):
            return None
        value = row[idx].value
        if value is None or value == "":
            return None
        return str(value).strip()

    grouped: dict[str, list[SampleItem]] = {}
    population_info: dict[str, dict[str, str | None]] = {}

    for row in rows[1:]:
        if all(c.value in (None, "") for c in row):
            continue

        test_step_id = cell_str(row, "test_step_id")
        sample_id = cell_str(row, "sample_id")
        identifying_details = cell_str(row, "identifying_details")
        if not test_step_id or not sample_id or not identifying_details:
            raise ValueError(
                f"{path.name}: row {row[0].row} is missing test_step_id, sample_id, "
                f"or identifying_details"
            )

        key_fields = {name: cell_str(row, name) for name in key_field_columns}
        key_fields = {k: v for k, v in key_fields.items() if v is not None} or None

        grouped.setdefault(test_step_id, []).append(
            SampleItem(
                sample_id=sample_id,
                test_step_id=test_step_id,
                identifying_details=identifying_details,
                key_fields=key_fields,
            )
        )

        if test_step_id not in population_info:
            pop_size_str = cell_str(row, "population_size")
            selection_method = cell_str(row, "selection_method")
            if selection_method not in ("random", "haphazard", "judgmental", "all_items"):
                raise ValueError(
                    f"{path.name}: test_step_id {test_step_id!r} has selection_method "
                    f"{selection_method!r}, must be one of random/haphazard/judgmental/all_items"
                )
            population_info[test_step_id] = {
                "population_description": cell_str(row, "population_description"),
                "population_size": int(pop_size_str) if pop_size_str else None,
                "selection_method": selection_method,
            }

    manifests: dict[str, SamplePopulationManifest] = {}
    for test_step_id, samples in grouped.items():
        info = population_info[test_step_id]
        manifests[test_step_id] = SamplePopulationManifest(
            test_step_id=test_step_id,
            population_description=info["population_description"] or "",
            population_size=info["population_size"],
            sample_size=len(samples),
            selection_method=info["selection_method"],
            samples=samples,
        )
    return manifests
