from __future__ import annotations

import pytest

from agent.schemas import (
    CheckSampleCoverageError,
    CheckSampleCoverageInput,
    EvidenceItem,
    FlagExceptionInput,
    RequestAdditionalSupportInput,
    SearchCySupportInput,
)
from agent.tools import ToolContext, check_sample_coverage, flag_exception, request_additional_support, search_cy_support


@pytest.fixture
def ctx() -> ToolContext:
    items = [
        EvidenceItem(
            evidence_id="ev_001",
            source_file="Recon.xlsx",
            source_type="excel_cell",
            location="Sheet1!C4",
            extracted_text="Accrual recalculation ties to GL export with no variance.",
            extraction_confidence=1.0,
            preview_ref="Recon.xlsx!Sheet1!C4",
        ),
        EvidenceItem(
            evidence_id="ev_002",
            source_file="POs.xlsx",
            source_type="excel_table",
            location="Sheet1!A1:D5",
            extracted_table=[["PO", "Amount"], ["48213", "12450"]],
            extraction_confidence=1.0,
            preview_ref="POs.xlsx!Sheet1!A1:D5",
        ),
        EvidenceItem(
            evidence_id="ev_003",
            source_file="scan.pdf",
            source_type="image_ocr",
            location="scan.pdf p.1",
            extraction_confidence=0.0,
            preview_ref="scan.pdf p.1",
        ),
    ]
    return ToolContext(evidence_items=items)


def test_search_finds_relevant_item(ctx: ToolContext):
    out = search_cy_support(SearchCySupportInput(query="accrual GL reconciliation"), ctx)
    assert out.results
    assert out.results[0].evidence_id == "ev_001"
    assert "ev_001" in ctx.evidence_ids_returned_by_search


def test_search_no_match_returns_empty(ctx: ToolContext):
    out = search_cy_support(SearchCySupportInput(query="zzz nonexistent zzz"), ctx)
    assert out.results == []


def test_search_respects_evidence_types_filter(ctx: ToolContext):
    out = search_cy_support(
        SearchCySupportInput(query="accrual PO amount", evidence_types=["excel_table"]), ctx
    )
    assert all(r.evidence_id != "ev_001" for r in out.results)


def test_search_unindexable_ocr_placeholder_is_not_a_match(ctx: ToolContext):
    out = search_cy_support(SearchCySupportInput(query="scan.pdf"), ctx)
    assert all(r.evidence_id != "ev_003" for r in out.results)


def test_check_sample_coverage_complete():
    out = check_sample_coverage(
        CheckSampleCoverageInput(required_sample_ids=["S01", "S02"], found_evidence_ids=["S01", "S02"])
    )
    assert out.total_required == 2
    assert out.total_found == 2
    assert out.missing == []
    assert out.complete is True
    assert out.coverage_pct == 100.0


def test_check_sample_coverage_partial():
    out = check_sample_coverage(
        CheckSampleCoverageInput(required_sample_ids=["S01", "S02", "S03"], found_evidence_ids=["S01"])
    )
    assert out.total_found == 1
    assert out.missing == ["S02", "S03"]
    assert out.complete is False
    assert round(out.coverage_pct, 2) == 33.33


def test_check_sample_coverage_no_sample_list():
    out = check_sample_coverage(CheckSampleCoverageInput(required_sample_ids=[], found_evidence_ids=["S01"]))
    assert isinstance(out, CheckSampleCoverageError)
    assert out.error == "no_sample_list_found"


def test_flag_exception_records_and_returns_id(ctx: ToolContext):
    out = flag_exception(
        FlagExceptionInput(
            test_step_id="TS-4.2", description="Missing approval", evidence_ids=["ev_002"], severity="high"
        ),
        ctx,
    )
    assert out.exception_id in ctx.exceptions
    assert ctx.exceptions[out.exception_id].severity == "high"


def test_request_additional_support_records_and_returns_id(ctx: ToolContext):
    out = request_additional_support(
        RequestAdditionalSupportInput(test_step_id="TS-4.2", description="Need clearer scan", reason="illegible"),
        ctx,
    )
    assert out.request_id in ctx.support_requests
    assert ctx.support_requests[out.request_id].reason == "illegible"
