"""Tool-loop tests against a fake Anthropic client -- no network call, no API
key needed. These prove the orchestration logic (tool dispatch, audit log,
forced-close nudging, the fabrication guard, and the required_sample_ids
server-side override) independent of whether the model is actually reachable.
Running run_test_step() against the real API is a separate, manual check
once a key is configured.
"""

from __future__ import annotations

import pytest

from agent.loop import IncompleteRunError, build_user_turn, run_test_step, TestStepRequest
from agent.schemas import ConclusionOutput, EvidenceItem, SamplePopulationManifest, SampleItem
from agent.tests.conftest import FakeClient, fake_response as response, fake_usage, text_block as text, tool_use


def _message_text(content) -> str:
    """content is a bare string, or (after the cache-breakpoint pass) a list
    of blocks -- normalize to a searchable string either way.
    """
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        return " ".join(block.get("text", "") for block in content if isinstance(block, dict))
    return ""


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


def test_py_conclusion_text_is_optional_and_defaults_empty():
    request = TestStepRequest(
        test_step_id="TS-4.2",
        control_id="C-14",
        control_objective_ref="CO-4",
        control_objective_text="Accruals are recorded completely and accurately.",
        test_step_text="Recalculate the accrual and agree to the GL.",
    )
    assert request.py_conclusion_text == ""


def test_build_user_turn_omits_py_conclusion_line_when_blank():
    request = TestStepRequest(
        test_step_id="TS-4.2",
        control_id="C-14",
        control_objective_ref="CO-4",
        control_objective_text="Accruals are recorded completely and accurately.",
        test_step_text="Recalculate the accrual and agree to the GL.",
    )
    turn = build_user_turn(request)
    assert "PY conclusion: not separately provided" in turn


def test_build_user_turn_includes_py_conclusion_when_given(request_: TestStepRequest):
    turn = build_user_turn(request_)
    assert "PY conclusion: Satisfied. No exceptions noted." in turn


def test_cache_reads_are_cost_weighted_not_face_value(evidence_items, sample_manifest, request_):
    # A real run got budget-aborted mid-investigation because cache READS
    # (billed at ~1/10th of fresh input) were counted at face value -- 12
    # well-cached turns "spent" 300K raw tokens while the actual bill was a
    # small fraction of that. 100K cache-read tokens must count as ~10K
    # weighted units, so three such turns (30K weighted) stay far under a
    # 300K budget instead of tripping it at raw count.
    client = FakeClient(
        responses=[
            response(
                [tool_use("t1", "search_cy_support", {"query": "accrual", "top_k": 5})],
                usage=fake_usage(input_tokens=500, cache_read_input_tokens=100_000),
            ),
            response(
                [tool_use("t2", "search_cy_support", {"query": "accrual GL", "top_k": 5})],
                usage=fake_usage(input_tokens=500, cache_read_input_tokens=100_000),
            ),
            response(
                [
                    tool_use(
                        "t3",
                        "submit_conclusion",
                        _submit_conclusion_input(conclusion="insufficient_evidence", evidence_citations=[]),
                    )
                ],
                usage=fake_usage(input_tokens=500, cache_read_input_tokens=100_000),
            ),
        ]
    )
    # Raw-summed, these 3 turns would be ~301,500 and trip a 300K budget
    # before the conclusion lands; cost-weighted they're ~31,500.
    conclusion, _ = run_test_step(
        request_, evidence_items, sample_manifest, client, max_total_tokens=300_000
    )
    assert conclusion.conclusion == "insufficient_evidence"


def test_output_tokens_weighted_heavier_than_input(evidence_items, sample_manifest, request_):
    # Output bills ~5x input -- 30K output tokens/turn must count as ~150K
    # weighted, tripping a 200K budget on turn 2 even though the raw sum
    # (60K) would look comfortably under it.
    client = FakeClient(
        responses=[
            response([text("stalling")], stop_reason="end_turn", usage=fake_usage(output_tokens=30_000))
            for _ in range(15)
        ]
    )
    with pytest.raises(IncompleteRunError) as exc_info:
        run_test_step(request_, evidence_items, sample_manifest, client, max_total_tokens=200_000)
    assert exc_info.value.reason == "token_budget_exceeded"
    assert exc_info.value.turns_used == 2


def test_build_user_turn_includes_cy_evidence_inventory(request_: TestStepRequest, evidence_items):
    turn = build_user_turn(request_, cy_evidence=evidence_items)
    assert "CY evidence inventory" in turn
    assert "[ev_001]" in turn
    assert "[ev_002]" in turn
    assert "fishing for it" in turn.lower()
    # The map is not license to cite unretrieved items.
    assert "retrieve an item with search_cy_support before citing" in turn


def test_build_user_turn_without_inventory_unchanged(request_: TestStepRequest):
    assert "CY evidence inventory" not in build_user_turn(request_)


def test_run_test_step_puts_inventory_in_first_turn(evidence_items, sample_manifest, request_):
    client = FakeClient(
        responses=[
            response(
                [
                    tool_use(
                        "t1",
                        "submit_conclusion",
                        _submit_conclusion_input(conclusion="insufficient_evidence", evidence_citations=[]),
                    )
                ]
            )
        ]
    )
    run_test_step(request_, evidence_items, sample_manifest, client)
    first = client.calls[0]["messages"][0]["content"]
    text = first if isinstance(first, str) else " ".join(
        b.get("text", "") for b in first if isinstance(b, dict)
    )
    assert "CY evidence inventory" in text
    assert "[ev_001]" in text


def test_truncated_tool_call_is_explained_not_left_as_field_required(
    evidence_items, sample_manifest, request_
):
    # A real run died here: submit_conclusion's JSON was cut off at the
    # output limit, so the schema's LAST fields (confidence,
    # confidence_rationale, additional_support_requests, exceptions) never
    # arrived and validation reported "Field required". Reading that as
    # forgotten fields, the model retried with a LONGER narrative and lost
    # one more field each time -- three dead submits, ~60K weighted tokens.
    truncated_input = {
        "test_step_id": "TS-4.2",
        "control_objective_ref": "CO-4",
        "conclusion": "satisfied",
        "narrative": "A very long narrative that ran past the output limit...",
        "evidence_citations": [],
        "procedures_performed": ["Inspected support."],
        # confidence / confidence_rationale / exceptions /
        # additional_support_requests never made it out of the model.
    }
    client = FakeClient(
        responses=[
            response([tool_use("t1", "submit_conclusion", truncated_input)], stop_reason="max_tokens"),
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

    assert audit_log[0].is_error is True
    # The retry turn must carry an explicit truncation explanation.
    retry_messages = client.calls[1]["messages"]
    assert any("hit the output length limit" in _message_text(m["content"]) for m in retry_messages)
    assert conclusion.conclusion == "insufficient_evidence"


def test_no_tool_call_after_truncation_also_explains_it(evidence_items, sample_manifest, request_):
    client = FakeClient(
        responses=[
            response([text("a long partial answer")], stop_reason="max_tokens"),
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
    run_test_step(request_, evidence_items, sample_manifest, client)
    assert any(
        "hit the output length limit" in _message_text(m["content"]) for m in client.calls[1]["messages"]
    )


def test_output_ceiling_is_large_enough_for_a_real_conclusion(evidence_items, sample_manifest, request_):
    # 4096 could not hold a 2-sample conclusion (multi-paragraph narrative
    # plus nine cited quotes), which is what caused the truncation above.
    from agent.loop import MAX_OUTPUT_TOKENS

    client = FakeClient(
        responses=[
            response(
                [
                    tool_use(
                        "t1",
                        "submit_conclusion",
                        _submit_conclusion_input(conclusion="insufficient_evidence", evidence_citations=[]),
                    )
                ]
            )
        ]
    )
    run_test_step(request_, evidence_items, sample_manifest, client)
    assert client.calls[0]["max_tokens"] == MAX_OUTPUT_TOKENS
    assert MAX_OUTPUT_TOKENS >= 16_000


def test_sample_roster_names_each_selected_item(request_: TestStepRequest):
    # A real 2-sample run burned its whole budget: the model was told only
    # "2 item(s) selected" and had to discover both the sample_ids and what
    # each item was by trial and error.
    request_.sample_size = 2
    request_.population_size = 30
    request_.samples = [
        SampleItem(
            sample_id="1",
            test_step_id="TS-4.2",
            identifying_details="invoice number: 2859; payee: Premiere Onboard LLC; amount: 30000",
        ),
        SampleItem(
            sample_id="2",
            test_step_id="TS-4.2",
            identifying_details="invoice number: 35713082; payee: Lockton Companies; amount: 11193",
        ),
    ]
    turn = build_user_turn(request_)

    assert "sample_id '1'" in turn and "sample_id '2'" in turn
    assert "Premiere Onboard LLC" in turn
    assert "Lockton Companies" in turn
    # The ids must be presented as what check_sample_coverage expects, so
    # the model doesn't have to discover them from a failed call.
    assert "found_evidence_ids" in turn
    # Their support files were named "sample 2 support" -- matching on
    # filenames rather than field values is the trap.
    assert "NOT on a filename" in turn


def test_sample_roster_keeps_identifying_fields_over_filler(request_: TestStepRequest):
    # A real E1 sample row is ~560 chars over 21 columns. A flat 240-char
    # cap truncated it mid-record, keeping "Discount Available: 0" and
    # cutting off the invoice number, invoice date and business unit --
    # exactly the fields the model needs to match evidence to an item.
    key_fields = {
        "Sample #": "1",
        "Match Doc Ty F0413.DCTM": "PK",
        "Check/ Item F0413.DOCM": "881477",
        "Payee Number F0413.PYE": "Premiere Onboard LLC",
        "Total Payment Amount F0413.PAAP": "30000",
        "Document Number F0414.DOC": "11743472",
        "Discount Available F0414.ADSC": "0",
        "Discount Taken F0414.ADSA": "0",
        "PO Doc. Number F0414.PO": "",
        "Business Unit F0414.MCU": "Talent Acquisition",
        "Invoice Date F0411.DIVJ": "2025-09-22 00:00:00",
        "Invoice Number F0411.VINV": "2859",
        "Batch Date F0411.DICJ": "2025-10-07 00:00:00",
    }
    request_.sample_size = 1
    request_.samples = [
        SampleItem(
            sample_id="1",
            test_step_id="TS-4.2",
            identifying_details="; ".join(f"{k}: {v}" for k, v in key_fields.items()),
            key_fields=key_fields,
        )
    ]
    turn = build_user_turn(request_)

    assert "Invoice Number F0411.VINV: 2859" in turn
    assert "Batch Date F0411.DICJ" in turn
    assert "Talent Acquisition" in turn
    # Zero/blank fields carry no identifying information and shouldn't eat
    # the character budget.
    assert "Discount Available" not in turn
    assert "PO Doc. Number" not in turn


def test_sample_roster_is_capped_on_a_large_sample(request_: TestStepRequest):
    request_.sample_size = 40
    request_.samples = [
        SampleItem(sample_id=str(i), test_step_id="TS-4.2", identifying_details=f"invoice {i}")
        for i in range(40)
    ]
    turn = build_user_turn(request_)
    assert "and 15 more item(s)" in turn


def test_build_user_turn_omits_sample_line_when_unknown(request_: TestStepRequest):
    # Neither sample_size nor population_size set -- must not fabricate a
    # sample-size line out of nothing.
    turn = build_user_turn(request_)
    assert "Sample:" not in turn


def test_build_user_turn_states_sample_is_complete_when_population_known(request_: TestStepRequest):
    # This is the direct fix for the real run that treated a correct,
    # intentional 1-item sample as suspicious purely for lack of a
    # population figure to check it against.
    request_.sample_size = 1
    request_.population_size = 1
    turn = build_user_turn(request_)
    assert "Sample: 1 item(s) selected for testing from a population of 1." in turn
    assert "do not treat it as" in turn


def test_build_user_turn_flags_uncertainty_when_population_unknown(request_: TestStepRequest):
    request_.sample_size = 1
    turn = build_user_turn(request_)
    assert "Sample: 1 item(s) selected for testing." in turn
    assert "population size was not provided" in turn
    assert "request_additional_support" in turn


def test_py_excerpts_are_capped_not_dumped_in_full(request_: TestStepRequest):
    # A real run showed a large PY testing file (or, before extraction
    # chunked large tables, one huge extracted table) pushing a single test
    # step to ~400K input tokens per API call -- this caps it.
    huge_items = [
        EvidenceItem(
            evidence_id=f"py_ev_{i:03d}",
            source_file="PY_Testing_Big.pdf",
            source_type="pdf_text",
            location=f"PY_Testing_Big.pdf p.{i}",
            extracted_text="x" * 5_000,
            extraction_confidence=1.0,
            preview_ref=f"PY_Testing_Big.pdf p.{i}",
        )
        for i in range(1, 20)
    ]
    request_.py_support_excerpts = huge_items

    turn = build_user_turn(request_)

    assert len(turn) < 30_000  # nowhere near 19 * 5,000+ chars if it dumped everything
    assert "more PY excerpt(s) omitted for length" in turn
    assert "py_ev_001" in turn  # at least the first excerpt is still there, not dropped entirely


def test_small_py_excerpts_are_not_truncated(request_: TestStepRequest):
    small_items = [
        EvidenceItem(
            evidence_id="py_ev_001",
            source_file="PY_Testing.pdf",
            source_type="pdf_text",
            location="PY_Testing.pdf p.1",
            extracted_text="Recalculated, agreed to GL, no variance.",
            extraction_confidence=1.0,
            preview_ref="PY_Testing.pdf p.1",
        )
    ]
    request_.py_support_excerpts = small_items

    turn = build_user_turn(request_)

    assert "more PY excerpt(s) omitted" not in turn
    assert "Recalculated, agreed to GL, no variance." in turn


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


def test_cache_breakpoint_on_newest_message_never_accumulates(evidence_items, sample_manifest, request_):
    # A real run showed input tokens per request ballooning -- traced back to
    # no cache_control anywhere on the growing messages[] (only on the
    # system prompt), so every turn re-billed the entire history at full
    # price. This locks in the fix: a breakpoint always lands on the very
    # last message before each call, and never piles up beyond one.
    client = FakeClient(
        responses=[
            response([tool_use("t1", "search_cy_support", {"query": "accrual", "top_k": 5})]),
            response([tool_use("t2", "search_cy_support", {"query": "accrual again", "top_k": 5})]),
            response(
                [tool_use("t3", "check_sample_coverage", {"required_sample_ids": [], "found_evidence_ids": ["S01"]})]
            ),
            response([tool_use("t4", "submit_conclusion", _submit_conclusion_input())]),
        ]
    )

    run_test_step(request_, evidence_items, sample_manifest, client)

    assert len(client.calls) == 4
    for call in client.calls:
        last_content = call["messages"][-1]["content"]
        assert isinstance(last_content, list), "last message should be in block form, not a bare string"
        assert last_content[-1].get("cache_control") == {"type": "ephemeral"}

        total_markers = sum(
            1
            for m in call["messages"]
            if isinstance(m["content"], list)
            for block in m["content"]
            if isinstance(block, dict) and "cache_control" in block
        )
        assert total_markers == 1, "messages[] should never carry more than one active breakpoint"


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
            response(
                [tool_use("t3", "check_sample_coverage", {"required_sample_ids": [], "found_evidence_ids": ["S01"]})]
            ),
            response([tool_use("t4", "submit_conclusion", _submit_conclusion_input())]),
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


def test_satisfied_without_coverage_check_is_rejected(evidence_items, sample_manifest, request_):
    # The system prompt ASKS for check_sample_coverage before concluding; this
    # proves the backend REFUSES a sampled "satisfied" that skipped it -- the
    # model is pushed to check coverage and only then gets accepted.
    client = FakeClient(
        responses=[
            response([tool_use("t1", "search_cy_support", {"query": "accrual", "top_k": 5})]),
            response([tool_use("t2", "submit_conclusion", _submit_conclusion_input())]),  # no coverage call yet
            response(
                [tool_use("t3", "check_sample_coverage", {"required_sample_ids": [], "found_evidence_ids": ["S01"]})]
            ),
            response([tool_use("t4", "submit_conclusion", _submit_conclusion_input())]),
        ]
    )

    conclusion, audit_log = run_test_step(request_, evidence_items, sample_manifest, client)

    rejected = audit_log[1]
    assert rejected.tool_name == "submit_conclusion"
    assert rejected.is_error is True
    assert "check_sample_coverage" in rejected.output["error"]
    assert conclusion.conclusion == "satisfied"


def test_satisfied_with_incomplete_coverage_is_rejected(evidence_items, request_):
    # Two required samples, evidence found for only one -- "satisfied" must
    # bounce with the missing sample named, and a downgraded conclusion
    # (insufficient_evidence) must still go through.
    manifest = SamplePopulationManifest(
        test_step_id="TS-4.2",
        population_description="All month-end accrual reconciliations",
        sample_size=2,
        selection_method="all_items",
        samples=[
            SampleItem(sample_id="S01", test_step_id="TS-4.2", identifying_details="October accrual"),
            SampleItem(sample_id="S02", test_step_id="TS-4.2", identifying_details="November accrual"),
        ],
    )
    client = FakeClient(
        responses=[
            response([tool_use("t1", "search_cy_support", {"query": "accrual", "top_k": 5})]),
            response(
                [tool_use("t2", "check_sample_coverage", {"required_sample_ids": [], "found_evidence_ids": ["S01"]})]
            ),
            response([tool_use("t3", "submit_conclusion", _submit_conclusion_input())]),  # satisfied, 1/2 covered
            response(
                [
                    tool_use(
                        "t4",
                        "submit_conclusion",
                        _submit_conclusion_input(conclusion="insufficient_evidence", evidence_citations=[]),
                    )
                ]
            ),
        ]
    )

    conclusion, audit_log = run_test_step(request_, evidence_items, manifest, client)

    rejected = audit_log[2]
    assert rejected.is_error is True
    assert "S02" in rejected.output["error"]
    assert "1/2" in rejected.output["error"]
    assert conclusion.conclusion == "insufficient_evidence"


def test_wrapup_warning_injected_before_last_iteration(evidence_items, sample_manifest, request_):
    # The $65 failure: the model ground through every turn with no idea a
    # limit existed, then aborted with nothing. It must now be told to wrap
    # up one turn before the cliff.
    client = FakeClient(
        responses=[
            response([tool_use("t1", "search_cy_support", {"query": "accrual", "top_k": 5})]),
            response([tool_use("t2", "search_cy_support", {"query": "accrual again", "top_k": 5})]),
            response(
                [
                    tool_use(
                        "t3",
                        "submit_conclusion",
                        _submit_conclusion_input(conclusion="insufficient_evidence", evidence_citations=[]),
                    )
                ]
            ),
        ]
    )

    conclusion, _ = run_test_step(request_, evidence_items, sample_manifest, client, max_iterations=3)

    assert conclusion.conclusion == "insufficient_evidence"
    # Not warned yet going into turn 2...
    assert not any(
        "close to its turn/token limit" in _message_text(m["content"]) for m in client.calls[1]["messages"]
    )
    # ...warned going into turn 3, the final allowed iteration.
    assert any(
        "close to its turn/token limit" in _message_text(m["content"]) for m in client.calls[2]["messages"]
    )


def test_wrapup_warning_injected_near_token_budget(evidence_items, sample_manifest, request_):
    client = FakeClient(
        responses=[
            response(
                [tool_use("t1", "search_cy_support", {"query": "accrual", "top_k": 5})],
                usage=fake_usage(input_tokens=850),  # 85% of the 1000 budget
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

    conclusion, _ = run_test_step(
        request_, evidence_items, sample_manifest, client, max_total_tokens=1_000
    )

    assert conclusion.conclusion == "insufficient_evidence"
    assert any(
        "close to its turn/token limit" in _message_text(m["content"]) for m in client.calls[1]["messages"]
    )


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
    # inserted before the model's second attempt. Content may be a bare
    # string or (after the cache-breakpoint pass converts the last message)
    # a list of text blocks -- check both shapes.
    second_call_messages = client.calls[1]["messages"]
    assert any("Every turn must end in a tool call" in _message_text(m["content"]) for m in second_call_messages)


def test_max_iterations_exhausted_raises(evidence_items, sample_manifest, request_):
    client = FakeClient(
        responses=[response([text("stalling")], stop_reason="end_turn") for _ in range(5)]
    )
    with pytest.raises(IncompleteRunError, match="did not reach submit_conclusion") as exc_info:
        run_test_step(request_, evidence_items, sample_manifest, client, max_iterations=3)

    err = exc_info.value
    assert err.reason == "max_iterations"
    assert err.turns_used == 3
    # Every stalling turn still nudged the model and got logged -- nothing
    # from those (paid-for) turns should be discarded just because the step
    # never reached submit_conclusion.
    assert len(err.audit_log) == 0  # no tool calls were ever made, just nudges -- audit_log is legitimately empty here


def test_none_cache_usage_fields_do_not_crash_token_tracking(
    evidence_items, sample_manifest, request_
):
    # On the real SDK, usage.cache_creation_input_tokens /
    # cache_read_input_tokens are Optional[int]: the attribute EXISTS but is
    # None when unpopulated, so getattr(usage, ..., 0) returns None and a
    # bare sum raises TypeError on the very first real turn -- crashing the
    # run through the generic error path with the audit log lost, which is
    # the exact failure class the circuit breaker exists to prevent.
    none_usage = fake_usage(input_tokens=100, output_tokens=50)
    none_usage.cache_creation_input_tokens = None
    none_usage.cache_read_input_tokens = None

    client = FakeClient(
        responses=[
            response(
                [tool_use("t1", "submit_conclusion", {
                    "test_step_id": "TS-4.2",
                    "control_objective_ref": "CO-4",
                    "conclusion": "insufficient_evidence",
                    "narrative": "No evidence was found.",
                    "evidence_citations": [],
                    "procedures_performed": ["Searched CY support."],
                    "ipe_completeness_accuracy_status": "not_applicable",
                    "ipe_completeness_accuracy_evidence": [],
                    "exceptions": [],
                    "additional_support_requests": [],
                    "confidence": "low",
                    "confidence_rationale": "Nothing relevant was found.",
                })],
                usage=none_usage,
            )
        ]
    )
    conclusion, _ = run_test_step(request_, evidence_items, sample_manifest, client)
    assert conclusion.conclusion == "insufficient_evidence"


def test_refusal_preserves_audit_log(evidence_items, sample_manifest, request_):
    # A refusal after real tool calls must carry them out, same as the
    # budget/iteration aborts -- not discard them via a bare RuntimeError.
    client = FakeClient(
        responses=[
            response([tool_use("t1", "search_cy_support", {"query": "accrual", "top_k": 5})]),
            response([text("declining")], stop_reason="refusal"),
        ]
    )
    with pytest.raises(IncompleteRunError, match="declined") as exc_info:
        run_test_step(request_, evidence_items, sample_manifest, client)

    err = exc_info.value
    assert err.reason == "model_refusal"
    assert len(err.audit_log) == 1
    assert err.audit_log[0].tool_name == "search_cy_support"


def test_on_turn_callback_fires_every_turn(evidence_items, sample_manifest, request_):
    # This is what the Streamlit app hooks to show live progress -- it must
    # fire during the run, not just be inferable after the fact from the
    # returned audit_log.
    client = FakeClient(
        responses=[
            response([tool_use("t1", "search_cy_support", {"query": "accrual", "top_k": 5})]),
            response([tool_use("t2", "submit_conclusion", {
                "test_step_id": "TS-4.2",
                "control_objective_ref": "CO-4",
                "conclusion": "insufficient_evidence",
                "narrative": "No evidence was found.",
                "evidence_citations": [],
                "procedures_performed": ["Searched CY support."],
                "ipe_completeness_accuracy_status": "not_applicable",
                "ipe_completeness_accuracy_evidence": [],
                "exceptions": [],
                "additional_support_requests": [],
                "confidence": "low",
                "confidence_rationale": "Nothing relevant was found.",
            })]),
        ]
    )
    calls = []
    run_test_step(
        request_,
        evidence_items,
        sample_manifest,
        client,
        on_turn=lambda turn, tokens, log: calls.append((turn, tokens, len(log))),
    )
    assert calls == [(1, 0, 1), (2, 0, 2)]


def test_token_budget_exceeded_raises_before_max_iterations(evidence_items, sample_manifest, request_):
    # Each turn "spends" 200K tokens -- with a 300K budget, the circuit
    # breaker should trip after the 2nd turn, well before max_iterations=15
    # would ever be reached. This is the direct fix for the real run that
    # burned $65 grinding through all 15 iterations with zero output.
    client = FakeClient(
        responses=[
            response([text("stalling")], stop_reason="end_turn", usage=fake_usage(input_tokens=200_000))
            for _ in range(15)
        ]
    )
    with pytest.raises(IncompleteRunError, match="token") as exc_info:
        run_test_step(request_, evidence_items, sample_manifest, client, max_total_tokens=300_000)

    err = exc_info.value
    assert err.reason == "token_budget_exceeded"
    assert err.turns_used == 2
    assert err.tokens_used == 400_000
    # Only 2 of the 15 available fake responses were ever consumed -- proof
    # the loop actually stopped early instead of running to max_iterations.
    assert len(client.calls) == 2


def test_token_budget_does_not_abort_the_turn_that_actually_concludes(
    evidence_items, sample_manifest, request_
):
    # A turn that reaches submit_conclusion should always be allowed to
    # finish and return, even if its own usage pushes the running total past
    # the budget -- the call already happened and was already paid for;
    # throwing away a successful conclusion at the last second would be
    # strictly worse than letting it land.
    client = FakeClient(
        responses=[
            response(
                [
                    tool_use(
                        "call_1",
                        "submit_conclusion",
                        {
                            "test_step_id": "TS-4.2",
                            "control_objective_ref": "CO-4",
                            "conclusion": "insufficient_evidence",
                            "narrative": "No evidence was found.",
                            "evidence_citations": [],
                            "procedures_performed": ["Searched CY support."],
                            "ipe_completeness_accuracy_status": "not_applicable",
                            "ipe_completeness_accuracy_evidence": [],
                            "exceptions": [],
                            "additional_support_requests": [],
                            "confidence": "low",
                            "confidence_rationale": "Nothing relevant was found.",
                        },
                    )
                ],
                usage=fake_usage(input_tokens=500_000),
            )
        ]
    )
    conclusion, audit_log = run_test_step(
        request_, evidence_items, sample_manifest, client, max_total_tokens=300_000
    )
    assert conclusion.conclusion == "insufficient_evidence"
    assert len(audit_log) == 1
