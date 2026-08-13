"""Tool-loop tests against a fake Anthropic client -- no network call, no API
key needed. These prove the orchestration logic (tool dispatch, audit log,
forced-close nudging, the fabrication guard, and the required_sample_ids
server-side override) independent of whether the model is actually reachable.
Running run_test_step() against the real API is a separate, manual check
once a key is configured.
"""

from __future__ import annotations

import pytest

from agent.loop import run_test_step, TestStepRequest
from agent.schemas import ConclusionOutput, EvidenceItem, SamplePopulationManifest, SampleItem
from agent.tests.conftest import FakeClient, fake_response as response, text_block as text, tool_use


@pytest.fixture
def evidence_items() -> list[EvidenceItem]:
    return [
        EvidenceItem(
            evidence_id="ev_001",
            source_file="Recon.xlsx",
            source_type="excel_cell",
            location="Sheet1!C4",
            extracted_text="Accrual recalculation, ending balance $482,110 ties to GL export.",
            extraction_confidence=1.0,
            preview_ref="Recon.xlsx!Sheet1!C4",
        ),
        EvidenceItem(
            evidence_id="ev_002",
            source_file="POs.xlsx",
            source_type="excel_table",
            location="Sheet1!A1:D5",
            extracted_text="Purchase order approval log, unrelated vendor invoices.",
            extraction_confidence=1.0,
            preview_ref="POs.xlsx!Sheet1!A1:D5",
        ),
    ]


@pytest.fixture
def sample_manifest() -> SamplePopulationManifest:
    return SamplePopulationManifest(
        test_step_id="TS-4.2",
        population_description="All month-end accrual reconciliations",
        sample_size=1,
        selection_method="all_items",
        samples=[SampleItem(sample_id="S01", test_step_id="TS-4.2", identifying_details="October accrual")],
    )


@pytest.fixture
def request_() -> TestStepRequest:
    return TestStepRequest(
        test_step_id="TS-4.2",
        control_id="C-14",
        control_objective_ref="CO-4",
        control_objective_text="Accruals are recorded completely and accurately.",
        test_step_text="Recalculate the accrual and agree to the GL.",
        py_conclusion_text="Satisfied. No exceptions noted.",
    )


def _submit_conclusion_input(**overrides) -> dict:
    base = dict(
        test_step_id="TS-4.2",
        control_objective_ref="CO-4",
        conclusion="satisfied",
        narrative="Recalculated accrual ties to GL with no variance.",
        evidence_citations=[
            {
                "evidence_id": "ev_001",
                "source_file": "Recon.xlsx",
                "location": "Sheet1!C4",
                "quote_or_summary": "Accrual recalculation, ending balance $482,110",
                "relevance": "Recalculated figure used to test the accrual",
            }
        ],
        procedures_performed=["recalculation"],
        ipe_completeness_accuracy_status="not_applicable",
        ipe_completeness_accuracy_evidence=[],
        exceptions=[],
        additional_support_requests=[],
        confidence="high",
        confidence_rationale="Native Excel source, full sample coverage.",
    )
    base.update(overrides)
    return base


def test_happy_path_search_then_coverage_then_submit(evidence_items, sample_manifest, request_):
    client = FakeClient(
        responses=[
            response([tool_use("t1", "search_cy_support", {"query": "accrual recalculation", "top_k": 5})]),
            response(
                [
                    tool_use(
                        "t2",
                        "check_sample_coverage",
                        {"required_sample_ids": [], "found_evidence_ids": ["S01"]},
                    )
                ]
            ),
            response([tool_use("t3", "submit_conclusion", _submit_conclusion_input())]),
        ]
    )

    conclusion, audit_log = run_test_step(request_, evidence_items, sample_manifest, client)

    assert isinstance(conclusion, ConclusionOutput)
    assert conclusion.conclusion == "satisfied"
    assert conclusion.sample_coverage is not None
    assert conclusion.sample_coverage.complete is True
    assert conclusion.model_metadata.tool_call_count == 3
    assert len(audit_log) == 3
    assert [e.tool_name for e in audit_log] == [
        "search_cy_support",
        "check_sample_coverage",
        "submit_conclusion",
    ]
    assert all(not e.is_error for e in audit_log)


def test_required_sample_ids_overridden_server_side(evidence_items, sample_manifest, request_):
    # The model passes a made-up required_sample_ids list; the loop must
    # ignore it and use the real manifest (["S01"]) instead.
    client = FakeClient(
        responses=[
            response(
                [
                    tool_use(
                        "t1",
                        "check_sample_coverage",
                        {"required_sample_ids": ["BOGUS-1", "BOGUS-2"], "found_evidence_ids": ["S01"]},
                    )
                ]
            ),
            response(
                [
                    tool_use(
                        "t2",
                        "submit_conclusion",
                        _submit_conclusion_input(),
                    )
                ]
            ),
        ]
    )
    # Need ev_001 to have been "returned by search" for submit_conclusion to
    # pass the fabrication check -- but this test skips search_cy_support on
    # purpose. Use insufficient_evidence instead, which needs no citations.
    client.responses[1] = response(
        [
            tool_use(
                "t2",
                "submit_conclusion",
                _submit_conclusion_input(conclusion="insufficient_evidence", evidence_citations=[]),
            )
        ]
    )

    conclusion, audit_log = run_test_step(request_, evidence_items, sample_manifest, client)

    coverage_entry = audit_log[0]
    assert coverage_entry.output["total_required"] == 1  # from the real manifest, not the bogus list
    assert coverage_entry.output["missing"] == []
    assert conclusion.conclusion == "insufficient_evidence"


def test_fabricated_citation_is_rejected_and_model_can_retry(evidence_items, sample_manifest, request_):
    client = FakeClient(
        responses=[
            # First attempt cites an evidence_id never returned by search -- rejected.
            response(
                [
                    tool_use(
                        "t1",
                        "submit_conclusion",
                        _submit_conclusion_input(
                            evidence_citations=[
                                {
                                    "evidence_id": "ev_999",
                                    "source_file": "Recon.xlsx",
                                    "location": "Sheet1!C4",
                                    "quote_or_summary": "made up",
                                    "relevance": "made up",
                                }
                            ]
                        ),
                    )
                ]
            ),
            # Retry: search first, then cite a real evidence_id.
            response([tool_use("t2", "search_cy_support", {"query": "accrual", "top_k": 5})]),
            response([tool_use("t3", "submit_conclusion", _submit_conclusion_input())]),
        ]
    )

    conclusion, audit_log = run_test_step(request_, evidence_items, sample_manifest, client)

    assert audit_log[0].tool_name == "submit_conclusion"
    assert audit_log[0].is_error is True
    assert "never returned by search_cy_support" in audit_log[0].output["error"]
    assert conclusion.conclusion == "satisfied"


def test_tool_call_count_matches_audit_log_even_after_invalid_input(evidence_items, sample_manifest, request_):
    # A live run caught this: a submit_conclusion attempt that fails pydantic
    # validation (not just the fabrication check -- e.g. the IPE
    # status/evidence pairing rule) used to take an early-return path in
    # _execute_tool that skipped the tool_call_count increment entirely, so
    # model_metadata.tool_call_count came back lower than len(audit_log).
    client = FakeClient(
        responses=[
            response(
                [
                    tool_use(
                        "t1",
                        "submit_conclusion",
                        _submit_conclusion_input(
                            ipe_completeness_accuracy_status="validated",
                            ipe_completeness_accuracy_evidence=[],  # invalid: violates the pairing rule
                        ),
                    )
                ]
            ),
            response(
                [
                    tool_use(
                        "t2",
                        "submit_conclusion",
                        _submit_conclusion_input(conclusion="insufficient_evidence", evidence_citations=[]),
                    )
                ]
            ),
        ]
    )

    conclusion, audit_log = run_test_step(request_, evidence_items, sample_manifest, client)

    assert len(audit_log) == 2
    assert audit_log[0].is_error is True
    assert conclusion.model_metadata.tool_call_count == 2


def test_no_tool_call_is_nudged_not_accepted_as_final(evidence_items, sample_manifest, request_):
    client = FakeClient(
        responses=[
            response([text("Let me think about this.")], stop_reason="end_turn"),
            response(
                [
                    tool_use(
                        "t1",
                        "submit_conclusion",
                        _submit_conclusion_input(conclusion="insufficient_evidence", evidence_citations=[]),
                    )
                ]
            ),
        ]
    )

    conclusion, audit_log = run_test_step(request_, evidence_items, sample_manifest, client)

    assert conclusion.conclusion == "insufficient_evidence"
    # First call produced no tool_use -- a nudge message should have been
    # inserted before the model's second attempt.
    second_call_messages = client.calls[1]["messages"]
    assert any(
        isinstance(m["content"], str) and "Every turn must end in a tool call" in m["content"]
        for m in second_call_messages
    )


def test_max_iterations_exhausted_raises(evidence_items, sample_manifest, request_):
    client = FakeClient(
        responses=[response([text("stalling")], stop_reason="end_turn") for _ in range(5)]
    )
    with pytest.raises(RuntimeError, match="did not reach submit_conclusion"):
        run_test_step(request_, evidence_items, sample_manifest, client, max_iterations=3)
