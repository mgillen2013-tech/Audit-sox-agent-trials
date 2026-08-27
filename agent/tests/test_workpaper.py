"""Workpaper generation tests: build from real ConclusionOutput objects and
read the produced files back (openpyxl / pdfplumber) rather than trusting
that "no exception" means a usable document.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pdfplumber
import pytest

from agent.loop import AuditLogEntry
from agent.schemas import ConclusionOutput, EvidenceCitation, ModelMetadata, SampleCoverage
from agent.workpaper import build_workpaper, workpaper_path_for


@pytest.fixture
def spec() -> dict:
    return {
        "control_id": "C-14",
        "control_objective_ref": "CO-4",
        "control_objective_text": "Accruals are recorded completely and accurately.",
        "test_steps": [
            {"test_step_id": "TS-4.2", "test_step_text": "Recalculate the accrual and agree to the GL."},
            {"test_step_id": "TS-9.9", "test_step_text": "A step that failed mid-run."},
        ],
    }


@pytest.fixture
def results() -> dict:
    conclusion = ConclusionOutput(
        test_step_id="TS-4.2",
        control_objective_ref="CO-4",
        conclusion="satisfied",
        narrative="Recalculated accrual ties to GL with no variance.",
        evidence_citations=[
            EvidenceCitation(
                evidence_id="ev_0001",
                source_file="Recon_Oct2026.xlsx",
                location="Sheet1!A1:B2",
                quote_or_summary="Recalculation ties to GL export with no variance",
                relevance="CY support for the accrual recalculation",
            )
        ],
        procedures_performed=["Recalculated the accrual", "Agreed the balance to the GL export"],
        ipe_completeness_accuracy_status="not_applicable",
        ipe_completeness_accuracy_evidence=[],
        exceptions=[],
        additional_support_requests=["Q4 support for the November accrual"],
        confidence="high",
        confidence_rationale="Single clear source, full sample coverage.",
        sample_coverage=SampleCoverage(
            total_required=1, total_found=1, missing=[], coverage_pct=100.0, complete=True
        ),
        model_metadata=ModelMetadata(
            model="claude-opus-5", prompt_version="v1", timestamp="2026-08-27T12:00:00+00:00", tool_call_count=3
        ),
    )
    failed = {
        "error": "test step 'TS-9.9' exceeded the 50,000-token budget after 4 turn(s)",
        "reason": "token_budget_exceeded",
        "tokens_used": 51_200,
        "turns_used": 4,
        "audit_log": [
            AuditLogEntry(
                turn=1,
                tool_name="search_cy_support",
                tool_use_id="t1",
                input={"query": "november accrual reconciliation", "top_k": 5},
                output={"results": []},
                is_error=False,
                timestamp="2026-08-27T12:01:00+00:00",
            )
        ],
    }
    return {"TS-4.2": {"conclusion": conclusion, "audit_log": []}, "TS-9.9": failed}


def test_path_matches_py_file_type(tmp_path: Path):
    assert workpaper_path_for("PY_C14.xlsm", "C-14", tmp_path).suffix == ".xlsx"
    assert workpaper_path_for("PY_C14.pdf", "C-14", tmp_path).suffix == ".pdf"
    # Control ids with filesystem-hostile characters can't break the filename.
    assert "/" not in workpaper_path_for("PY.pdf", "C/14 (AP)", tmp_path).name


def test_xlsx_workpaper_contents(spec: dict, results: dict, tmp_path: Path):
    path = build_workpaper(spec, results, "PY_Testing_C14.xlsx", tmp_path)

    wb = openpyxl.load_workbook(path)
    assert set(wb.sheetnames) == {"Summary", "TS-4.2", "TS-9.9"}

    def sheet_text(name: str) -> str:
        return " ".join(
            str(cell.value) for row in wb[name].iter_rows() for cell in row if cell.value is not None
        )

    summary = sheet_text("Summary")
    assert "DRAFT" in summary
    assert "C-14" in summary
    assert "Satisfied" in summary
    assert "INCOMPLETE" in summary  # the failed step is visible on the summary, not silently missing

    step = sheet_text("TS-4.2")
    assert "Recalculated accrual ties to GL with no variance." in step
    assert "ev_0001" in step
    assert "Q4 support for the November accrual" in step
    assert "claude-opus-5" in step

    failed = sheet_text("TS-9.9")
    assert "INCOMPLETE" in failed
    assert "token_budget_exceeded" in failed
    assert "november accrual reconciliation" in failed  # searches attempted before the abort are documented


def test_pdf_workpaper_contents(spec: dict, results: dict, tmp_path: Path):
    path = build_workpaper(spec, results, "PY_Testing_C14.pdf", tmp_path)

    assert path.suffix == ".pdf"
    with pdfplumber.open(path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    assert "DRAFT" in text
    assert "C-14" in text
    assert "Satisfied" in text
    assert "INCOMPLETE" in text
    assert "ev_0001" in text
