from __future__ import annotations

import pytest

from agent.schemas import (
    EvidenceItem,
    FlagExceptionInput,
    RequestAdditionalSupportInput,
    SearchCySupportInput,
)
from agent.tools import ToolContext, flag_exception, locate_value, request_additional_support, search_cy_support


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


# --------------------------------------------------------------------------
# Response size -- the single largest recurring cost in a run, since results
# land in the conversation and are re-sent on every subsequent turn.
# --------------------------------------------------------------------------


def _big_ctx(n_items: int) -> ToolContext:
    return ToolContext(
        evidence_items=[
            EvidenceItem(
                evidence_id=f"ev_{i:03d}",
                source_file="big.xlsx",
                source_type="excel_table",
                location=f"Sheet1!A{i}",
                extracted_text="accrual " + ("x" * 5_000),
                extraction_confidence=1.0,
                preview_ref="p",
            )
            for i in range(1, n_items + 1)
        ]
    )


def test_search_response_is_budgeted_as_a_whole_not_per_result():
    # top_k=8 x a 2,000-char snippet cap was ~4,000 tokens parked in
    # history permanently. The budget is shared across results instead.
    ctx = _big_ctx(8)
    out = search_cy_support(SearchCySupportInput(query="accrual", top_k=8), ctx)
    total = sum(len(r.snippet) for r in out.results)
    assert len(out.results) == 8
    assert total <= 8_000, f"one search returned {total} chars"


def test_a_targeted_search_still_returns_near_full_text():
    # The incentive has to run the right way: ask precisely, get more.
    ctx = _big_ctx(8)
    narrow = search_cy_support(SearchCySupportInput(query="accrual", top_k=2), ctx)
    broad = search_cy_support(SearchCySupportInput(query="accrual", top_k=8), ctx)
    assert len(narrow.results[0].snippet) > len(broad.results[0].snippet)


def test_top_k_is_bounded():
    # Unbounded, one call could park arbitrarily much in the conversation.
    with pytest.raises(Exception):
        SearchCySupportInput(query="accrual", top_k=500)


def test_locate_value_is_strict_about_digits_loose_about_punctuation(ctx: ToolContext):
    # "11,193.00" printed and "11193.00" in the model's head are the same
    # number -- being strict there would send it hunting for another anchor
    # when the one it had was fine. A DIFFERENT number must still miss.
    from agent.schemas import LocateValueInput

    ids = list(ctx._evidence_by_id)
    target = ctx._evidence_by_id[ids[0]]
    target.extracted_text = "Invoice 35713082 paid 11,193.00 on 02/11/2026"

    assert locate_value(LocateValueInput(evidence_id=ids[0], value="11193.00"), ctx).found
    assert locate_value(LocateValueInput(evidence_id=ids[0], value="11,193.00"), ctx).found
    assert not locate_value(LocateValueInput(evidence_id=ids[0], value="11,194.00"), ctx).found
