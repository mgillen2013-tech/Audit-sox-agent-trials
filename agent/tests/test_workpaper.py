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


def _make_support_pdf(path: Path) -> None:
    from reportlab.lib.pagesizes import letter as letter_size
    from reportlab.pdfgen import canvas

    c = canvas.Canvas(str(path), pagesize=letter_size)
    c.drawString(72, 700, "Approved / Vendor Number 21022452 / Cost Center 0510")
    c.drawString(72, 680, "Payment Amount 11,678.47 dated 12/04/2025")
    c.showPage()
    c.save()


@pytest.fixture
def annotated_setup(spec: dict, results: dict, tmp_path: Path) -> tuple[dict, dict, Path]:
    # A real support file on disk whose deterministic re-extraction yields
    # ev_0001 for the page text -- the id the citation below points at.
    support = tmp_path / "support"
    support.mkdir()
    _make_support_pdf(support / "approval.pdf")
    spec = {**spec, "cy_support_files": ["approval.pdf"]}

    conclusion: ConclusionOutput = results["TS-4.2"]["conclusion"]
    annotated = conclusion.model_copy(
        update={
            "evidence_citations": [
                EvidenceCitation(
                    evidence_id="ev_0001",
                    source_file="approval.pdf",
                    location="approval.pdf p.1",
                    quote_or_summary="Approved / Vendor Number 21022452 / Cost Center 0510",
                    relevance="Approval evidence for the sampled payment",
                )
            ]
        }
    )
    results = {"TS-4.2": {"conclusion": annotated, "audit_log": []}}
    return spec, results, support


def test_xlsx_workpaper_embeds_annotated_exhibit(annotated_setup, tmp_path: Path):
    spec, results, support = annotated_setup
    path = build_workpaper(spec, results, "PY_Testing_C14.xlsx", tmp_path, support_dir=support)

    wb = openpyxl.load_workbook(path)
    ws = wb["TS-4.2"]
    text = " ".join(str(c.value) for r in ws.iter_rows() for c in r if c.value is not None)
    assert "Tickmark" in text
    assert "approval.pdf p.1" in text

    # Exhibits live on their own sheet -- inline, full-page renders pushed
    # the conclusion ~110 rows down the step sheet.
    assert len(ws._images) == 0
    assert "see the 'TS-4.2 - Exhibits' sheet" in text

    ex_ws = wb["TS-4.2 - Exhibits"]
    assert len(ex_ws._images) == 1
    ex_text = " ".join(str(c.value) for r in ex_ws.iter_rows() for c in r if c.value is not None)
    assert "Evidence exhibits" in ex_text
    assert "Tickmark letters match" in ex_text


def test_pdf_workpaper_embeds_annotated_exhibit(annotated_setup, tmp_path: Path):
    spec, results, support = annotated_setup
    path = build_workpaper(spec, results, "PY_Testing_C14.pdf", tmp_path, support_dir=support)

    with pdfplumber.open(path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        image_count = sum(len(page.images) for page in pdf.pages)
    assert "Evidence exhibits" in text
    assert image_count >= 1


def test_workpaper_without_support_dir_is_text_only(annotated_setup, tmp_path: Path):
    # No support_dir (or a failed re-extraction) must degrade to the plain
    # citation table, never break the build.
    spec, results, _ = annotated_setup
    path = build_workpaper(spec, results, "PY_Testing_C14.xlsx", tmp_path)
    wb = openpyxl.load_workbook(path)
    assert len(wb["TS-4.2"]._images) == 0


def test_excel_cited_evidence_gets_text_excerpt(spec: dict, results: dict, tmp_path: Path):
    # Excel-sourced citations can't be screenshotted -- the cited range's
    # extracted rows are excerpted into the workpaper instead.
    support = tmp_path / "support"
    support.mkdir()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Item", "Amount"])
    ws.append(["October accrual", "482110"])
    wb.save(support / "Recon_Oct2026.xlsx")
    spec = {**spec, "cy_support_files": ["Recon_Oct2026.xlsx"]}

    path = build_workpaper(spec, results, "PY_Testing_C14.xlsx", tmp_path, support_dir=support)
    out = openpyxl.load_workbook(path)
    text = " ".join(str(c.value) for r in out["TS-4.2 - Exhibits"].iter_rows() for c in r if c.value is not None)
    assert "Exhibit A" in text
    assert "October accrual | 482110" in text


def test_step_sheet_puts_answers_before_supporting_detail(spec: dict, results: dict, tmp_path: Path):
    # A reviewer opening the sheet must see the verdict, coverage, IPE
    # status, exceptions and open requests before scrolling into the
    # narrative and evidence table -- inline exhibits previously pushed all
    # of that ~110 rows down the page.
    path = build_workpaper(spec, results, "PY_Testing_C14.xlsx", tmp_path)
    ws = openpyxl.load_workbook(path)["TS-4.2"]

    rows = {}
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.isupper() and len(cell.value) > 3:
                rows.setdefault(cell.value, cell.row)

    assert rows["CONCLUSION"] < rows["EXCEPTIONS"] < rows["ADDITIONAL SUPPORT REQUESTED"]
    assert rows["ADDITIONAL SUPPORT REQUESTED"] < rows["DOCUMENTATION"]
    assert rows["DOCUMENTATION"] < rows["PROCEDURES PERFORMED"] < rows["EVIDENCE CITED"]
    assert rows["EVIDENCE CITED"] < rows["PREPARED BY"]


def test_empty_exceptions_says_none_not_silence(spec: dict, results: dict, tmp_path: Path):
    # "no exceptions" and "we didn't look" must not be indistinguishable in
    # a workpaper -- an empty section is a real audit assertion.
    path = build_workpaper(spec, results, "PY_Testing_C14.xlsx", tmp_path)
    ws = openpyxl.load_workbook(path)["TS-4.2"]
    text = " ".join(str(c.value) for r in ws.iter_rows() for c in r if c.value is not None)
    assert "None noted." in text


def test_summary_sheet_carries_ipe_and_open_request_counts(spec: dict, results: dict, tmp_path: Path):
    path = build_workpaper(spec, results, "PY_Testing_C14.xlsx", tmp_path)
    ws = openpyxl.load_workbook(path)["Summary"]
    text = " ".join(str(c.value) for r in ws.iter_rows() for c in r if c.value is not None)
    assert "IPE status" in text
    assert "Open requests" in text
    assert "not_applicable" in text  # from the fixture's conclusion


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
