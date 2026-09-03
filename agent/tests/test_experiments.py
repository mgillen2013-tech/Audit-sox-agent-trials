"""The two options under test: templated narrative boxes, and a reviewer pass.

Neither is a decision yet. These tests pin down what each one actually
does so the comparison is on the artefacts, not on recollection of them.
"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from agent.narrative_styles import build_narrative
from agent.reviewer import SampleReview, review_sample


def _lead(text: str) -> str:
    return text


_LEGEND = [("A", "Invoice number agrees", "35713082"), ("B", "Amount agrees", "11,193.00")]


# --------------------------------------------------------------------------
# Option A: templated narrative
# --------------------------------------------------------------------------


def _text(paras) -> str:
    return " ".join(r.text for p in paras for r in p.runs)


def test_template_style_is_identical_between_samples_except_the_values():
    # The whole case for this option: a reviewer reading five tabs in a row
    # only has to check the values, because everything else is constant.
    a = _text(build_narrative("template", step_label="TS-1", model_narrative="one story",
                              legend=[("A", "Invoice agrees", "111")], satisfied=True,
                              lead_sentences=_lead))
    b = _text(build_narrative("template", step_label="TS-1", model_narrative="a different story",
                              legend=[("A", "Invoice agrees", "222")], satisfied=True,
                              lead_sentences=_lead))
    assert a.replace("111", "X") == b.replace("222", "X")


def test_model_style_lets_the_prose_differ_between_samples():
    a = _text(build_narrative("model", step_label="TS-1", model_narrative="one story",
                              legend=_LEGEND, satisfied=True, lead_sentences=_lead))
    b = _text(build_narrative("model", step_label="TS-1", model_narrative="a different story",
                              legend=_LEGEND, satisfied=True, lead_sentences=_lead))
    assert a != b


def test_template_style_keeps_the_models_prose_out_of_the_box():
    # It moves to the Summary rather than being deleted -- an unusual
    # sample still needs somewhere to say so.
    got = _text(build_narrative("template", step_label="TS-1",
                                model_narrative="A long explanation of the circumstances.",
                                legend=_LEGEND, satisfied=True, lead_sentences=_lead))
    assert "long explanation" not in got


def test_an_exception_is_never_closed_with_the_satisfied_wording():
    got = _text(build_narrative("template", step_label="TS-1", model_narrative="",
                                legend=_LEGEND, satisfied=False, lead_sentences=_lead))
    assert "No exceptions noted" not in got
    assert "Exception noted" in got


def test_both_styles_carry_the_same_legend():
    # The legend is what ties the box to the red marks on the exhibits, so
    # it must not depend on which experiment is running.
    for style in ("model", "template"):
        got = _text(build_narrative(style, step_label="", model_narrative="x",
                                    legend=_LEGEND, satisfied=True, lead_sentences=_lead))
        assert "A - " in got and "B - " in got
        assert "35713082" in got and "11,193.00" in got


# --------------------------------------------------------------------------
# Option B: the reviewer pass
# --------------------------------------------------------------------------


def _client(text: str):
    return SimpleNamespace(
        create_message=lambda **kw: SimpleNamespace(
            content=[SimpleNamespace(type="text", text=text)]
        )
    )


def _evidence(body: str):
    return [SimpleNamespace(evidence_id="ev_1", source_file="invoice.pdf",
                            location="p.1", extracted_text=body)]


def test_the_reviewer_reports_an_unsupported_line():
    got = review_sample(
        _client('{"lines":[{"letter":"A","verdict":"supported"},'
                '{"letter":"B","verdict":"unsupported","note":"amount not on the page"}]}'),
        "claude-opus-5", sample_id="1", legend=_LEGEND,
        evidence_items=_evidence("Invoice 35713082"), evidence_ids={"ev_1"},
    )
    assert not got.all_supported
    assert "B unsupported" in got.summary


def test_a_clean_review_says_so_countably():
    got = review_sample(
        _client('{"lines":[{"letter":"A","verdict":"supported"},{"letter":"B","verdict":"supported"}]}'),
        "claude-opus-5", sample_id="1", legend=_LEGEND,
        evidence_items=_evidence("Invoice 35713082 for 11,193.00"), evidence_ids={"ev_1"},
    )
    assert got.all_supported
    assert "2 tickmark(s) supported" in got.summary


def test_a_reviewer_failure_never_breaks_the_build():
    # The workpaper it checks is already built and valid. Losing it because
    # an OPTIONAL second opinion failed would be absurd.
    def boom(**kw):
        raise RuntimeError("model unavailable")

    got = review_sample(
        SimpleNamespace(create_message=boom), "claude-opus-5", sample_id="1",
        legend=_LEGEND, evidence_items=_evidence("text"), evidence_ids={"ev_1"},
    )
    assert got.error and "Not reviewed" in got.summary
    assert not got.all_supported


def test_unparseable_output_is_reported_not_guessed():
    got = review_sample(
        _client("I think it looks fine to me!"), "claude-opus-5", sample_id="1",
        legend=_LEGEND, evidence_items=_evidence("text"), evidence_ids={"ev_1"},
    )
    # No lines parsed -> nothing claimed. Silence beats inventing a verdict.
    assert not got.all_supported


def test_evidence_with_no_readable_text_is_not_reviewed_rather_than_passed():
    # An image-only page that was never OCR'd must not read as "supported"
    # just because there was nothing to contradict the claim.
    got = review_sample(
        _client('{"lines":[{"letter":"A","verdict":"supported"}]}'),
        "claude-opus-5", sample_id="1", legend=_LEGEND,
        evidence_items=_evidence(""), evidence_ids={"ev_1"},
    )
    assert not got.all_supported
    assert "readable text" in got.summary
