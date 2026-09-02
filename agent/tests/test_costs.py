"""Tests for the token-burn ledger.

The point of this module is to answer "where did the money go", so these
tests are mostly about the breakdown being *correct and attributable* --
that OCR is separable from testing, that the four token kinds are priced
apart rather than averaged, and that the prompt-composition rows sum to
what was really billed instead of to a chars/4 guess.
"""

from __future__ import annotations

import pytest

from agent.costs import TokenLedger, prices_for
from agent.loop import run_test_step, TestStepRequest
from agent.schemas import EvidenceItem, SampleItem, SamplePopulationManifest
from agent.tests.conftest import FakeClient, fake_response as response, fake_usage, tool_use


# --------------------------------------------------------------------------
# Pricing
# --------------------------------------------------------------------------


def test_opus_prices_match_the_published_rates():
    family, fresh, cache_write, cache_read, output = prices_for("claude-opus-5")
    assert family == "opus"
    assert (fresh, output) == (5.00, 25.00)
    assert cache_write == pytest.approx(6.25)  # 1.25x input
    assert cache_read == pytest.approx(0.50)  # 0.1x input


def test_sonnet_4_6_is_not_matched_as_plain_sonnet():
    # Substring matching is ordered most-specific-first; "sonnet" would
    # otherwise swallow "sonnet-4-6" and price it $1/MTok too cheap.
    assert prices_for("claude-sonnet-4-6")[1] == 3.00
    assert prices_for("claude-sonnet-5")[1] == 2.00


def test_unknown_deployment_name_falls_back_rather_than_crashing():
    # Foundry deployments are named by whoever created them, so an
    # unrecognised string is a normal input, not an error case.
    family, fresh, _, _, _ = prices_for("my-custom-deployment")
    assert family == "opus" and fresh == 5.00


# --------------------------------------------------------------------------
# Recording
# --------------------------------------------------------------------------


def test_none_cache_fields_do_not_crash_the_ledger():
    # The real SDK returns None (not a missing attribute) for unpopulated
    # cache fields, which once crashed the spending cap on the first live
    # turn. A cost report must never be the thing that kills a finished run.
    ledger = TokenLedger("claude-opus-5")
    ledger.record(fake_usage(input_tokens=100, output_tokens=50), group="TS-1", label="x", turn=1)
    assert ledger.records[0].cache_read_tokens == 0
    assert ledger.dollars == pytest.approx((100 * 5.0 + 50 * 25.0) / 1e6)


def test_missing_usage_object_records_a_zero_row():
    ledger = TokenLedger("claude-opus-5")
    ledger.record(None, group="OCR", label="scan.pdf p.1")
    assert ledger.records[0].raw_tokens == 0
    assert ledger.dollars == 0.0


def test_the_four_token_kinds_are_priced_apart():
    # The whole reason this exists: one blended number cannot show that a
    # million cache reads cost less than 20K of output.
    ledger = TokenLedger("claude-opus-5")
    ledger.record(
        fake_usage(
            input_tokens=1_000_000,
            output_tokens=1_000_000,
            cache_creation_input_tokens=1_000_000,
            cache_read_input_tokens=1_000_000,
        ),
        group="TS-1",
        label="x",
        turn=1,
    )
    kinds = {name: usd for name, _, usd, _ in ledger.kind_rows()}
    assert kinds["Output"] == pytest.approx(25.00)
    assert kinds["Cache writes"] == pytest.approx(6.25)
    assert kinds["Fresh input (never cached)"] == pytest.approx(5.00)
    assert kinds["Cache reads"] == pytest.approx(0.50)
    # Sorted most expensive first, so the thing to fix is the top row.
    assert ledger.kind_rows()[0][0] == "Output"


def test_ocr_is_separable_from_testing():
    # OCR runs before the tool loop and outside the per-step cap, so "was
    # it OCR or was it testing" is the first question a surprising bill
    # raises -- and a single blended total can never answer it.
    ledger = TokenLedger("claude-opus-5")
    ledger.record(fake_usage(input_tokens=2_000), group="OCR", label="inv.pdf p.1")
    ledger.record(fake_usage(input_tokens=2_000), group="OCR", label="inv.pdf p.2")
    ledger.record(fake_usage(input_tokens=6_000), group="TS-1", label="search_cy_support", turn=1)

    rows = {g.group: g for g in ledger.group_rows()}
    assert rows["OCR"].calls == 2
    assert rows["OCR"].raw_tokens == 4_000
    assert rows["TS-1"].raw_tokens == 6_000
    # Run order, not alphabetical -- OCR genuinely happens first.
    assert [g.group for g in ledger.group_rows()] == ["OCR", "TS-1"]


def test_call_rows_filter_to_one_group():
    ledger = TokenLedger("claude-opus-5")
    ledger.record(fake_usage(input_tokens=1), group="OCR", label="a")
    ledger.record(fake_usage(input_tokens=1), group="TS-1", label="b", turn=1)
    assert [r.label for r, _ in ledger.call_rows("TS-1")] == ["b"]
    assert len(ledger.call_rows()) == 2


# --------------------------------------------------------------------------
# Prompt composition
# --------------------------------------------------------------------------


def test_prompt_mix_sums_to_the_measured_turn_1_input():
    # Apportioning a MEASURED total is the whole trick: a chars/4 estimate
    # would be wrong in aggregate, and a breakdown whose parts don't add up
    # to the real bill is not evidence of anything.
    ledger = TokenLedger("claude-opus-5")
    ledger.note_prompt_mix("TS-1", {"PY support excerpts": 8_000, "System prompt": 2_000})
    ledger.record(
        fake_usage(input_tokens=1_000, cache_creation_input_tokens=4_000, output_tokens=900),
        group="TS-1",
        label="search_cy_support",
        turn=1,
    )
    rows = ledger.prompt_mix_rows("TS-1")
    assert [r[0] for r in rows] == ["PY support excerpts", "System prompt"]  # largest first
    # 5,000 tokens were sent on turn 1; output is not part of the prompt.
    assert sum(r[2] for r in rows) == 5_000
    assert rows[0][2] == 4_000  # 80% of the characters


def test_prompt_mix_is_empty_when_nothing_was_recorded():
    assert TokenLedger("claude-opus-5").prompt_mix_rows("TS-1") == []


def test_prompt_mix_survives_a_step_that_never_reached_an_api_call():
    # note_prompt_mix runs before the first request precisely so a step
    # that fails immediately still says what it was about to send.
    ledger = TokenLedger("claude-opus-5")
    ledger.note_prompt_mix("TS-1", {"PY support excerpts": 8_000})
    assert ledger.prompt_mix_rows("TS-1") == [("PY support excerpts", 8_000, 0, 0.0)]


# --------------------------------------------------------------------------
# Wiring: the ledger must reflect a real loop run
# --------------------------------------------------------------------------


def _request() -> TestStepRequest:
    return TestStepRequest(
        test_step_id="TS-1",
        control_id="C-1",
        control_objective_ref="CO-1",
        control_objective_text="Invoices are approved before payment.",
        test_step_text="Verify approval precedes payment.",
        py_conclusion_text="No exceptions.",
        py_support_excerpts=[],
        sample_size=1,
        samples=[SampleItem(sample_id="1", test_step_id="TS-1", identifying_details="INV-1")],
    )


def _conclusion_input() -> dict:
    return {
        "test_step_id": "TS-1",
        "control_objective_ref": "CO-1",
        "conclusion": "satisfied",
        "narrative": "Approved 10/1, paid 11/10.",
        "evidence_citations": [
            {
                "evidence_id": "ev_0001",
                "source_file": "inv.pdf",
                "location": "p.1",
                "quote_or_summary": "Approved",
                "relevance": "approval",
                "sample_id": "1",
            }
        ],
        "procedures_performed": ["inspection"],
        "ipe_completeness_accuracy_status": "not_applicable",
        "ipe_completeness_accuracy_evidence": [],
        "exceptions": [],
        "additional_support_requests": [],
        "confidence": "high",
        "confidence_rationale": "Native text.",
    }


def test_a_real_run_itemises_each_turn_by_what_it_did():
    evidence = [
        EvidenceItem(
            evidence_id="ev_0001",
            source_file="inv.pdf",
            source_type="pdf_text",
            location="inv.pdf p.1",
            extracted_text="Invoice INV-1 approved",
            extraction_confidence=1.0,
            preview_ref="p1",
        )
    ]
    client = FakeClient(
        [
            response(
                [tool_use("t1", "search_cy_support", {"query": "invoice approved"})],
                usage=fake_usage(input_tokens=500, cache_creation_input_tokens=9_000, output_tokens=200),
            ),
            response(
                [tool_use("t2", "check_sample_coverage", {"required_sample_ids": ["1"], "found_evidence_ids": ["1"]})],
                usage=fake_usage(input_tokens=200, cache_read_input_tokens=9_000, output_tokens=150),
            ),
            response(
                [tool_use("t3", "submit_conclusion", _conclusion_input())],
                usage=fake_usage(input_tokens=300, cache_read_input_tokens=9_500, output_tokens=1_200),
            ),
        ]
    )
    ledger = TokenLedger("claude-opus-5")
    run_test_step(
        _request(),
        evidence,
        SamplePopulationManifest(
            test_step_id="TS-1",
            sample_size=1,
            samples=[SampleItem(sample_id="1", test_step_id="TS-1", identifying_details="INV-1")],
        ),
        client,
        ledger=ledger,
    )

    # Every turn recorded, labelled by the tool it called -- "turn 3" alone
    # would not tell a reviewer that the expensive turn was the write-up.
    assert [(r.turn, r.label) for r in ledger.records] == [
        (1, "search_cy_support"),
        (2, "check_sample_coverage"),
        (3, "submit_conclusion"),
    ]
    # Caching visible as its own column rather than blended away. This is
    # the shape of a HEALTHY run and the reason the breakdown exists: cache
    # reads are by far the largest token count, and still cost less than
    # the cache write that is half their size and the output that is a
    # twelfth of it. A raw token dashboard shows the 18,500 at face value
    # and makes the run look several times more expensive than it was; a
    # single blended number cannot show which way it went at all.
    kinds = {name: (tok, usd) for name, tok, usd, _ in ledger.kind_rows()}
    assert kinds["Cache reads"][0] == 18_500
    assert max(kinds.values(), key=lambda v: v[0])[0] == 18_500
    assert kinds["Cache reads"][1] < kinds["Cache writes"][1]
    assert kinds["Cache reads"][1] < kinds["Output"][1]
    assert ledger.prompt_mix_rows("TS-1")  # composition captured for turn 1


def test_the_ledger_agrees_with_the_spending_cap_it_reports_on():
    # The cap counts cost-weighted tokens; the ledger counts dollars. They
    # are different units, but they must not disagree about which phase was
    # expensive -- if they did, the breakdown would be explaining a bill
    # that isn't the one the cap was watching.
    from agent.loop import _usage_tokens

    ledger = TokenLedger("claude-opus-5")
    usages = [
        fake_usage(input_tokens=10_000, output_tokens=500),
        fake_usage(input_tokens=200, cache_read_input_tokens=10_000, output_tokens=2_000),
    ]
    for i, u in enumerate(usages, start=1):
        ledger.record(u, group="TS-1", label="x", turn=i)

    weighted = sum(_usage_tokens(u) for u in usages)
    # $5/MTok fresh input is the unit the weights are expressed in, so
    # dollars and weighted units are the same quantity at a fixed scale.
    assert ledger.dollars == pytest.approx(weighted * 5.0 / 1e6, rel=1e-9)


def test_a_step_that_aborts_still_reports_what_it_spent():
    # Same preserve-paid-work rule the audit log follows: an aborted run
    # burned real money and must still be able to say on what.
    from agent.loop import IncompleteRunError

    client = FakeClient(
        [
            response([tool_use("t1", "search_cy_support", {"query": "a"})], usage=fake_usage(input_tokens=800))
            for _ in range(4)
        ]
    )
    ledger = TokenLedger("claude-opus-5")
    with pytest.raises(IncompleteRunError):
        run_test_step(
            _request(), [], None, client, max_iterations=3, max_total_tokens=10_000, ledger=ledger
        )
    assert len(ledger.records) == 3
    assert ledger.dollars > 0
