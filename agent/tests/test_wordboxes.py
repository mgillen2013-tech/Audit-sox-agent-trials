"""Tickmark box-location tests.

The OCR engine itself is stubbed here rather than run: these tests are about
the matching logic (which is where a box lands on the WRONG field), and real
OCR is slow and unnecessary to prove that. The values in the stub are taken
verbatim from a real scanned invoice.
"""

from __future__ import annotations

import pytest

from agent import wordboxes
from agent.wordboxes import _anchors, find_text_boxes

# A real invoice page: three $10,000.00 line items summing to a $30,000.00
# subtotal and total, plus the identifying fields a tickmark points at.
_REAL_INVOICE_LINES = [
    ((75, 105, 195, 132), "INVOICE", 1.0),
    ((53, 436, 165, 458), "Invoice no.: 2859", 0.99),
    ((55, 479, 211, 496), "Invoice date: 09/22/2025", 0.98),
    ((55, 499, 190, 515), "Due date: 11/06/2025", 0.98),
    ((700, 610, 780, 635), "$10,000.00", 0.99),
    ((700, 640, 780, 665), "$10,000.00", 0.99),
    ((700, 670, 780, 695), "$10,000.00", 0.99),
    ((767, 794, 862, 818), "$30,000.00", 0.99),
    ((758, 1093, 859, 1115), "$30,000.00", 0.99),
    ((414, 907, 464, 928), "0510", 0.95),
    ((421, 955, 512, 975), "21022452", 0.97),
    ((300, 1200, 600, 1230), "Approved by Diane Milosevic", 0.94),
    ((300, 1240, 600, 1270), "low confidence smudge", 0.20),
]


@pytest.fixture
def stub_ocr(monkeypatch):
    def _install(lines):
        monkeypatch.setattr(wordboxes, "ocr_line_boxes", lambda img: lines)

    return _install


# --------------------------------------------------------------------------
# Anchor extraction -- what the matcher considers "distinctive"
# --------------------------------------------------------------------------


def test_amounts_are_anchored_whole_not_in_fragments():
    # THE regression: "$30,000.00" naively tokenizes to 30 / 000 / 00, and
    # "000" is a substring of every round amount on the page. A real run
    # boxed the $10,000.00 line items because of exactly this.
    anchors = _anchors("Invoice payment amount $30,000.00 for Premiere Onboard LLC")
    assert anchors == ("3000000",)
    assert "000" not in anchors
    assert "00" not in anchors


def test_bare_years_are_not_anchors():
    # A year appears in every date on a document -- it identifies nothing.
    assert _anchors("Invoice date 09/22/2025, due 11/06/2025") == ()


def test_identifiers_and_invoice_numbers_are_anchors():
    anchors = _anchors("Vendor Number 21022452 / Cost Center 0510 / invoice 2859")
    assert set(anchors) == {"21022452", "0510", "2859"}
    # Most specific first, so the best box is tried before weaker ones.
    assert anchors[0] == "21022452"


def test_short_numbers_are_not_anchors():
    assert _anchors("line 12 of 30") == ()


# --------------------------------------------------------------------------
# Box selection
# --------------------------------------------------------------------------


def test_box_lands_on_the_cited_total_not_similar_line_items(stub_ocr):
    stub_ocr(_REAL_INVOICE_LINES)
    boxes = find_text_boxes(object(), "Invoice payment amount $30,000.00")

    assert boxes  # something was located
    ten_thousand_boxes = [b for b in boxes if b[1] in (610, 640, 670)]
    assert not ten_thousand_boxes, "boxed a $10,000.00 line item instead of the $30,000.00 cited"
    assert (767, 794, 862, 818) in boxes


def test_ambiguous_anchor_is_dropped_rather_than_guessed(stub_ocr):
    # An anchor matching many lines describes the page, not a value on it --
    # on audit evidence, no box beats a box on the wrong field.
    stub_ocr([((0, i * 20, 50, i * 20 + 15), "PV 11743472", 0.99) for i in range(6)])
    assert find_text_boxes(object(), "voucher 11743472") == []


def test_low_confidence_detections_are_ignored(stub_ocr):
    stub_ocr(_REAL_INVOICE_LINES)
    assert find_text_boxes(object(), "low confidence smudge 999999") == []


def test_prose_quote_falls_back_to_literal_fragment(stub_ocr):
    stub_ocr(_REAL_INVOICE_LINES)
    boxes = find_text_boxes(object(), "Approved by Diane Milosevic on 10/1")
    assert (300, 1200, 600, 1230) in boxes


def test_no_match_returns_nothing(stub_ocr):
    stub_ocr(_REAL_INVOICE_LINES)
    assert find_text_boxes(object(), "Purchase order 99887766") == []


def test_missing_ocr_engine_degrades_quietly(monkeypatch):
    # rapidocr absent must mean "whole-page exhibit", never a crash.
    monkeypatch.setattr(wordboxes, "_ocr_engine", lambda: None)
    assert wordboxes.ocr_line_boxes(object()) == []
    assert find_text_boxes(object(), "$30,000.00") == []
