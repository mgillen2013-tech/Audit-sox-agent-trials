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
