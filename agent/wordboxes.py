"""Locate quoted text on a rendered page image, for tickmark boxes.

The problem this solves: a native-text PDF lets pdfplumber search its text
layer and hand back exact coordinates, so a tickmark box lands tight on the
quoted words. A SCANNED page (an E1 screenshot saved to PDF, a photographed
invoice) has no text layer at all -- so those exhibits could only ever be
embedded whole, with the tickmark letter parked in the corner, pointing at
nothing in particular.

Vision OCR (agent/extraction/ocr.py) gave those pages readable TEXT, but a
transcription carries no coordinates, so it didn't help the boxes. This
module supplies the missing half: local OCR (rapidocr-onnxruntime, pip-only,
models bundled in the wheel -- no system binary, works offline) that returns
a box per detected text line.

Deliberate division of labor:
- Claude vision produces the text that goes IN the workpaper. It is the more
  accurate reader (local OCR misread one real invoice line as "Involce no.:
  2859").
- Local OCR is used ONLY to answer "where on this page is that?". It costs
  zero tokens and, unlike a vision call, it MEASURES coordinates rather than
  estimating them. On audit evidence a box sitting 40px off, pointing at the
  wrong field, is worse than no box at all -- so the thing that draws boxes
  should be the thing that measures.

Everything degrades quietly: no rapidocr installed, an unreadable page, or
no confident match all return [], and the caller falls back to the whole-page
exhibit it produced before.
"""

from __future__ import annotations

import re
from functools import lru_cache
from typing import Any

# Boxes are only drawn for reasonably confident detections -- a shaky match
# pointed at the wrong field is worse than an unboxed exhibit.
_MIN_OCR_CONFIDENCE = 0.5
_MAX_BOXES_PER_MARK = 3

# Anchor tokens: the distinctive values an audit tickmark actually points at
# (an amount, an invoice number, a vendor id, a cost center). These survive
# OCR far better than prose, and matching on them is what makes a box land on
# "$30,000.00" rather than somewhere in a paragraph that happens to share
# common words.
_AMOUNT_RE = re.compile(r"\d[\d,]*\.\d{2}")
_LONG_NUM_RE = re.compile(r"\b\d{4,}\b")
_ALNUM_ID_RE = re.compile(r"\b(?=[A-Za-z]*\d)[A-Za-z0-9\-]{5,}\b")
_YEAR_RE = re.compile(r"^(19|20)\d{2}$")

# An anchor is only useful if it identifies ONE place on the page. "000"
# (a fragment of "$30,000.00") matched every "$10,000.00" line item on a
# real invoice and boxed the wrong amounts; a bare year matches every date.
# Anchors shorter than this, and anchors that hit more lines than
# _MAX_ANCHOR_HITS, are treated as non-identifying and dropped.
_MIN_ANCHOR_LEN = 4
_MAX_ANCHOR_HITS = 2


def _ocr_engine() -> Any:
    global _ENGINE
    try:
        return _ENGINE
    except NameError:
        pass
    try:
        from rapidocr_onnxruntime import RapidOCR

        _ENGINE = RapidOCR()
    except Exception:  # noqa: BLE001 -- optional dependency; absence is not an error
        _ENGINE = None
    return _ENGINE


def ocr_line_boxes(pil_image: Any) -> list[tuple[tuple[float, float, float, float], str, float]]:
    """[(pixel bbox, text, confidence)] for every text line found on the
    page image. Returns [] when local OCR is unavailable or the page yields
    nothing.
    """
    engine = _ocr_engine()
    if engine is None:
        return []
    try:
        import numpy as np

        result, _ = engine(np.array(pil_image))
    except Exception:  # noqa: BLE001 -- a page that won't OCR just gets no boxes
        return []

    out = []
    for box, text, conf in result or []:
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        out.append(((min(xs), min(ys), max(xs), max(ys)), str(text), float(conf)))
    return out


def _normalize(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower())


@lru_cache(maxsize=256)
def _anchors(quote: str) -> tuple[str, ...]:
    """Distinctive values worth boxing, most specific first.

    Amounts are taken WHOLE ("$30,000.00" -> "3000000"), never as the
    fragments a naive tokenizer would split them into -- "000" is a
    substring of every round amount on the page. Bare years are dropped for
    the same reason: they appear in every date.
    """
    found: list[str] = []

    def add(raw: str) -> None:
        n = _normalize(raw)
        if len(n) >= _MIN_ANCHOR_LEN and n not in found and not _YEAR_RE.match(n):
            found.append(n)

    amounts = _AMOUNT_RE.findall(quote)
    for m in amounts:
        add(m)

    # Blank out the amounts before scanning for plain numbers, so their
    # digit groups can't come back as separate (and far less specific)
    # anchors of their own.
    remainder = _AMOUNT_RE.sub(" ", quote)
    for pattern in (_LONG_NUM_RE, _ALNUM_ID_RE):
        for m in pattern.findall(remainder):
            add(m)

    return tuple(sorted(found, key=len, reverse=True))


def find_text_boxes(
    lines: list[tuple[tuple[float, float, float, float], str, float]], quote: str
) -> list[tuple[float, float, float, float]]:
    """Pixel boxes among already-OCR'd `lines` where `quote`'s distinctive
    values appear.

    Takes lines rather than an image on purpose: OCR is by far the expensive
    step (~1.3s per page against ~0.06s to render it), and a page usually
    carries several tickmarks. Reading the page once and matching each quote
    against the result keeps a 3-citation page at one OCR pass instead of
    three.

    Matching is anchor-based rather than whole-string: the quote comes from
    the model and is often a summary or a stitched-together excerpt, while
    the page shows the raw values. Boxing "$30,000.00" and "2859" where they
    actually sit beats trying to match a sentence that was never printed on
    the page verbatim.
    """
    if not lines:
        return []

    confident = [(box, _normalize(text)) for box, text, conf in lines if conf >= _MIN_OCR_CONFIDENCE]
    boxes: list[tuple[float, float, float, float]] = []

    # Most specific anchor first, and an anchor that lights up several lines
    # is discarded rather than used: it is describing the page, not a value
    # on it. This is what keeps a "$30,000.00" tickmark off the "$10,000.00"
    # line items that share its digits.
    for anchor in _anchors(quote):
        hits = [box for box, text in confident if anchor in text]
        if not hits or len(hits) > _MAX_ANCHOR_HITS:
            continue
        for box in hits:
            if box not in boxes:
                boxes.append(box)
            if len(boxes) >= _MAX_BOXES_PER_MARK:
                return boxes

    if boxes:
        return boxes

    # No distinctive values in the quote (a prose citation, e.g. an approval
    # sentence) -- fall back to locating a long literal fragment of it.
    fragment = _normalize(quote)[:20]
    if len(fragment) < 12:
        return []
    prose_hits = [box for box, text in confident if fragment in text]
    if len(prose_hits) > _MAX_ANCHOR_HITS:
        return []
    return prose_hits[:_MAX_BOXES_PER_MARK]
