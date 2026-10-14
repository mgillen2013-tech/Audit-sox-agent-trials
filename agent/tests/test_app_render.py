"""Smoke tests that actually RENDER the Streamlit app.

Every prior round of review found its real defects in an output path
nothing exercised -- the PDF builder had silently drifted several features
behind the xlsx one for exactly this reason. The app is the other such
path: it is the only consumer of some of the ledger's API, and a renamed
method or a changed row shape would break it with every unit test still
green.

AppTest runs the real cy_testing_app.py in-process, so these go through the
same code a browser would.
"""

from __future__ import annotations

from pathlib import Path

import pytest

_APP = Path(__file__).resolve().parents[2] / "cy_testing_app.py"

from agent.costs import TokenLedger
from agent.tests.conftest import fake_usage

st_testing = pytest.importorskip("streamlit.testing.v1")


def _app():
    # default_timeout is generous: the app imports openpyxl, pdfplumber and
    # the anthropic SDK at module scope, which is slow but not hung.
    return st_testing.AppTest.from_file(str(_APP), default_timeout=120)


def _populated_ledger() -> TokenLedger:
    ledger = TokenLedger("claude-opus-5")
    ledger.record(fake_usage(input_tokens=1_800, output_tokens=900), group="OCR", label="inv.pdf p.1")
    ledger.note_prompt_mix(
        "TS-1", {"System prompt": 9_800, "PY support excerpts": 20_000, "Tool schemas": 6_400}
    )
    ledger.record(
        fake_usage(input_tokens=1_200, cache_creation_input_tokens=11_500, output_tokens=420),
        group="TS-1",
        label="search_cy_support",
        turn=1,
    )
    ledger.record(
        fake_usage(input_tokens=400, cache_read_input_tokens=15_200, output_tokens=3_100),
        group="TS-1",
        label="submit_conclusion",
        turn=2,
    )
    return ledger


def test_app_renders_with_no_run_yet():
    at = _app().run()
    assert not at.exception


def test_token_burn_renders_from_a_finished_run():
    at = _app()
    at.session_state["run_output"] = {
        "results": {},
        "wp_bytes": None,
        "wp_name": None,
        "ledger": _populated_ledger(),
    }
    at.run()
    assert not at.exception, at.exception

    labels = {m.label: m.value for m in at.metric}
    assert labels["API calls"] == "3"
    assert labels["Tokens billed"] == "34,520"
    assert labels["Estimated cost"].startswith("$")

    # Phase table, token-kind table, one turn-by-turn table per phase, and
    # the prompt-composition table for the step that has one.
    assert len(at.dataframe) == 5


def test_token_burn_says_nothing_was_billed_rather_than_showing_empty_tables():
    # A run where every step failed before its first API call. Empty tables
    # would read as a broken report rather than as a free run.
    at = _app()
    at.session_state["run_output"] = {
        "results": {},
        "wp_bytes": None,
        "wp_name": None,
        "ledger": TokenLedger("claude-opus-5"),
    }
    at.run()
    assert not at.exception
    assert any("No API calls were billed" in i.value for i in at.info)
    assert not at.dataframe


def test_an_old_session_without_a_ledger_still_renders():
    # st.session_state survives a code reload in a running Streamlit server,
    # so a session started before this field existed must not crash the app.
    at = _app()
    at.session_state["run_output"] = {"results": {}, "wp_bytes": None, "wp_name": None}
    at.run()
    assert not at.exception


def _real_conclusion():
    """A conclusion shaped exactly like one a real run produces."""
    from agent.schemas import ConclusionOutput, ModelMetadata, SampleCoverage

    return ConclusionOutput(
        test_step_id="TS-1",
        control_objective_ref="CO-1",
        conclusion="insufficient_evidence",
        narrative="Sample 2 evidenced approval before payment; no approval support was provided for sample 1.",
        evidence_citations=[
            {
                "evidence_id": "ev_0003",
                "source_file": "invoice.pdf",
                "location": "invoice.pdf p.1",
                "quote_or_summary": "Lockton invoice 35713082 for $11,193.00",
                "relevance": "the invoice",
                "sample_id": "2",
            }
        ],
        procedures_performed=["inspection"],
        sample_results=[
            {"sample_id": "1", "conclusion": "insufficient_evidence", "note": "no approval provided"},
            {"sample_id": "2", "conclusion": "satisfied"},
        ],
        attribute_results=[
            {
                "attribute": "Approved before payment",
                "sample_id": "2",
                "result": "satisfied",
                "value_observed": "Approved 02/09/2026, paid 02/11/2026",
                "evidence_ids": ["ev_0003"],
            }
        ],
        exceptions=[],
        additional_support_requests=["Selection 1 approval and payment support"],
        confidence="high",
        confidence_rationale="Both packages reviewed.",
        sample_coverage=SampleCoverage(
            total_required=2, total_found=2, missing=[], coverage_pct=100.0, complete=True
        ),
        model_metadata=ModelMetadata(
            model="claude-opus-5", prompt_version="v1", timestamp="2026-10-14T00:00:00Z", tool_call_count=8
        ),
    )


def test_a_real_conclusion_renders_without_crashing():
    # This is the test that was missing. The app tests only ever passed an
    # EMPTY results dict, so _render_conclusion -- the function that
    # actually touches a conclusion's fields -- was never executed by any
    # test at all. A field removed from the schema therefore sailed through
    # a green suite and crashed a real run AFTER the agent had finished,
    # losing a paid run's workpaper.
    at = _app()
    at.session_state["run_output"] = {
        "results": {"TS-1": {"conclusion": _real_conclusion(), "audit_log": []}},
        "wp_bytes": None,
        "wp_name": None,
        "ledger": _populated_ledger(),
    }
    at.run()
    assert not at.exception, at.exception


def test_every_field_the_app_renders_still_exists_on_the_conclusion():
    # Cheaper and blunter than the render test, and it names the field that
    # went missing instead of pointing at a line number.
    import re

    from agent.schemas import ConclusionOutput

    source = _APP.read_text()
    referenced = set(re.findall(r"conclusion\.([a-z_]+)", source))
    # Fields OR methods -- the app legitimately calls model_dump().
    unknown = {
        f for f in referenced
        if f not in ConclusionOutput.model_fields and not hasattr(ConclusionOutput, f)
    }
    assert not unknown, f"cy_testing_app.py reads conclusion fields that no longer exist: {unknown}"


def test_a_fallback_workpaper_is_announced_in_red_with_its_reason():
    # A real run fell back to the standalone layout and said so only in a
    # grey caption under a collapsed expander. The reviewer reasonably read
    # the different-looking file as a regression.
    at = _app()
    at.session_state["run_output"] = {
        "results": {},
        "wp_bytes": b"xlsx bytes",
        "wp_name": "156_CY_Testing_DRAFT.xlsx",
        "ledger": None,
        "wp_mirrored": False,
        "wp_fallback_reason": "it needs tabs named 'Sample' and 'Population'",
        "wp_warnings": ["could not mirror the prior-year layout: it needs tabs named 'Sample'"],
    }
    at.run()
    assert not at.exception, at.exception
    errors = " ".join(e.value for e in at.error)
    assert "NOT in last year's format" in errors
    assert "it needs tabs named 'Sample' and 'Population'" in errors
    # The notes are open, not hidden behind a click.
    assert any(e.proto.expanded for e in at.expander if "NOT in this workpaper" in e.label)


def test_a_mirrored_workpaper_shows_no_fallback_error():
    at = _app()
    at.session_state["run_output"] = {
        "results": {},
        "wp_bytes": b"xlsx bytes",
        "wp_name": "156_CY_Testing_wp_DRAFT.xlsx",
        "ledger": None,
        "wp_mirrored": True,
        "wp_fallback_reason": "",
        "wp_warnings": [],
    }
    at.run()
    assert not at.exception, at.exception
    assert not any("NOT in last year's format" in e.value for e in at.error)
