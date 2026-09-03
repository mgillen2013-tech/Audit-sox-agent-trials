"""Drop-in replacement for the package's original ocr_utils.py.

Same four names, same signatures, same semantics -- but backed by
pdfplumber + rapidocr-onnxruntime instead of pdf2image + pytesseract.

WHY THIS FILE EXISTS
--------------------
The original needed two things pip cannot install: the `tesseract-ocr`
binary and `poppler-utils`. On the machine this actually runs on --
Windows, code delivered as a downloaded zip, set up with a single
`pip install -r requirements.txt` -- both pip packages install happily and
then fail at RUNTIME, which is the worst possible shape for a dependency
failure: it looks like the code is broken rather than the environment.

pdfplumber renders PDF pages in pure Python, and rapidocr-onnxruntime
ships its own models in the wheel and runs offline with no system binary
and no PATH setup. Both are already dependencies of this project's
extraction layer, so this swap also stops the same job being done by two
different stacks in one repo.

WHAT CHANGES BEHAVIOURALLY
--------------------------
Tesseract emits one box per whitespace-delimited token; rapidocr emits one
box per detected TEXT LINE, with a quadrilateral rather than a rectangle.
Line boxes are useless for "draw a tight box around this one value", so
split_line_into_words() apportions a line's width across its tokens by
character count. That is an approximation -- proportional fonts make the
per-character width uneven -- but the box is a highlight, not a
measurement, and a few pixels either way is invisible at 150 DPI. The
original README makes the same trade for column widths and says so.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from io import BytesIO
from typing import Any

from PIL import Image, ImageChops

# Matches the original's constant, and agent/extraction/ocr.py's _OCR_DPI --
# the two OCR paths in this repo must rasterise at the same resolution or a
# box computed by one is misplaced when applied by the other.
RENDER_DPI = 150

_EMU_PER_INCH = 914_400


def px_to_emu(px: float, dpi: float) -> int:
    """Pixels to EMU at a stated DPI. Only valid for IMAGE coordinates --
    the cell grid uses a fixed 9525 EMU/px (96 DPI) constant instead, and
    conflating the two is the misalignment bug the methodology doc warns
    about.
    """
    return int(round(px * _EMU_PER_INCH / dpi))


def render_pdf_page(pdf_path: str, page: int = 0, dpi: int = RENDER_DPI) -> Image.Image:
    """One PDF page as a PIL image, via pdfplumber (pure Python).

    `page` is 0-indexed, matching the original's pdf2image convention --
    NOT pdfplumber's own 1-indexed page labels. Keeping the original
    convention matters because EvidenceImage.pdf_page values written for
    the old code must keep meaning the same page.
    """
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        # .original, not .annotated: pdfplumber's annotated render draws
        # its own debug rectangles over detected words, which would be
        # baked into the evidence image pasted in the workpaper.
        return pdf.pages[page].to_image(resolution=dpi).original.convert("RGB")


def autocrop(im: Image.Image, pad: int = 10, bg: tuple[int, int, int] = (255, 255, 255)) -> Image.Image:
    """Trim uniform-background margins, keeping `pad` px so a callout box
    near the edge isn't clipped. Unchanged from the original.
    """
    rgb = im.convert("RGB")
    diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, bg))
    bbox = diff.getbbox()
    if not bbox:
        return rgb
    left, top, right, bottom = bbox
    return rgb.crop(
        (
            max(0, left - pad),
            max(0, top - pad),
            min(rgb.width, right + pad),
            min(rgb.height, bottom + pad),
        )
    )


def load_and_prepare(
    source_path: str, pdf_page: int = 0, dpi: int = RENDER_DPI
) -> tuple[Image.Image, float]:
    """(cropped image, dpi actually used). Handles PDFs and plain images.

    The returned DPI must be used for every px->EMU conversion on this
    image. Do not assume 96 or 150: a pasted screenshot's own metadata can
    say anything, and getting it wrong scales every callout on the page.
    """
    if source_path.lower().endswith(".pdf"):
        return autocrop(render_pdf_page(source_path, page=pdf_page, dpi=dpi)), float(dpi)
    im = Image.open(source_path)
    native = im.info.get("dpi", (96, 96))
    return autocrop(im), float(native[0] or 96)


@dataclass
class WordBox:
    text: str
    left: int
    top: int
    width: int
    height: int

    @property
    def right(self) -> int:
        return self.left + self.width

    @property
    def bottom(self) -> int:
        return self.top + self.height


_ocr_engine: Any = None


def _engine() -> Any:
    """One lazily-built engine, reused. Construction loads the ONNX models
    (~1s) -- rebuilding it per image made a multi-exhibit page take several
    seconds for no reason.
    """
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR

        _ocr_engine = RapidOCR()
    return _ocr_engine


# Tokens are runs of non-space, non-colon characters, keeping a trailing
# colon on the label it belongs to. The colon matters: rapidocr frequently
# returns a whole labelled field as ONE string with the space eaten --
# "Amount:30,000.00", "Approvedby:DianeMilosevic". Splitting on whitespace
# alone leaves that as a single token, so a callout anchored on the amount
# gets a box spanning the label too -- measured at 222px covering
# "Amount:30,000.00" where the value alone is ~110px. Splitting at the
# colon puts the red box around the value, which is the entire point of
# the callout. Everything else stays intact: "30,000.00" and "11/10/2025"
# keep their commas and slashes.
_TOKEN_RE = re.compile(r"[^\s:]+:?")


def split_line_into_words(text: str, left: int, top: int, width: int, height: int) -> list[WordBox]:
    """Apportion a detected line's box across its tokens.

    rapidocr returns one box per text LINE; the callers want words. Each
    token's pixel span is derived from its CHARACTER span within the line,
    which assumes even character widths -- invoices and JDE extracts are
    proportional-font, so a token's edge can be off by a few pixels. That
    is acceptable for a highlight box, and is the same order of error the
    methodology doc already accepts for column-width math. It would NOT be
    acceptable if anything downstream measured these boxes.
    """
    if not text:
        return []
    matches = list(_TOKEN_RE.finditer(text))
    if not matches:
        return []

    per_char = width / len(text)
    boxes: list[WordBox] = []
    for m in matches:
        x0 = left + per_char * m.start()
        x1 = left + per_char * m.end()
        boxes.append(WordBox(m.group(), int(round(x0)), top, int(round(x1 - x0)), height))
    return boxes


def ocr_words(im: Image.Image) -> list[WordBox]:
    """Every word on the image with a pixel box, in no guaranteed order.

    Returns [] rather than raising when OCR finds nothing -- an unreadable
    image should surface as "anchor not found" from find_anchor_box, which
    names the anchor and the near-misses, not as an opaque failure here.
    """
    buf = BytesIO()
    im.convert("RGB").save(buf, "PNG")
    result, _ = _engine()(buf.getvalue())
    if not result:
        return []

    words: list[WordBox] = []
    for entry in result:
        # rapidocr rows are [quad_points, text, confidence]; the quad is
        # four (x, y) corners, not a rectangle, so take its bounding box.
        quad, text = entry[0], entry[1]
        if not text or not text.strip():
            continue
        xs = [float(p[0]) for p in quad]
        ys = [float(p[1]) for p in quad]
        left, top = int(round(min(xs))), int(round(min(ys)))
        width, height = int(round(max(xs) - min(xs))), int(round(max(ys) - min(ys)))
        words.extend(split_line_into_words(text, left, top, width, height))
    return words


def find_anchor_box(
    words: list[WordBox],
    anchor_text: str,
    *,
    occurrence: int = 0,
    extra_words: int = 0,
    pad: int = 6,
) -> tuple[int, int, int, int]:
    """Locate anchor_text among the word boxes and return a padded pixel
    box (x, y, w, h). Substring match, so a partial token still matches.

    Raises ValueError when the anchor isn't found -- loudly, rather than
    silently boxing the wrong place. A missing anchor nearly always means
    the OCR text differs slightly from what was expected (a misread digit,
    a ligature, different whitespace), so the error names the near-misses
    to make that diagnosable without re-running anything.

    Behaviour is unchanged from the original implementation.
    """
    needle = anchor_text.strip()
    if " " in needle:
        raise ValueError(
            f"anchor_text {anchor_text!r} contains a space, but word boxes are "
            f"single tokens. Anchor on the first word and use extra_words=N to "
            f"include the following N tokens instead."
        )

    matches = [w for w in words if needle in w.text]
    if not matches:
        raise ValueError(
            f"anchor_text {anchor_text!r} not found in OCR output. Nearby "
            f"candidates: {[w.text for w in words if needle[:3] and needle[:3] in w.text][:10]}"
        )
    if occurrence >= len(matches):
        raise ValueError(
            f"anchor_text {anchor_text!r} found {len(matches)} time(s), but "
            f"occurrence={occurrence} was requested."
        )

    matches.sort(key=lambda w: (w.top, w.left))
    start = matches[occurrence]
    box = [start.left, start.top, start.right, start.bottom]

    if extra_words:
        ordered = sorted((w for w in words if w is not start), key=lambda w: (w.top, w.left))
        start_idx = None
        for i, w in enumerate(ordered):
            if w.left == start.left and w.top == start.top and w.text == start.text:
                start_idx = i
                break
        following = ordered[start_idx + 1 : start_idx + 1 + extra_words] if start_idx is not None else []
        for w in following:
            box[0] = min(box[0], w.left)
            box[1] = min(box[1], w.top)
            box[2] = max(box[2], w.right)
            box[3] = max(box[3], w.bottom)

    x1, y1, x2, y2 = box
    return (x1 - pad, y1 - pad, (x2 - x1) + 2 * pad, (y2 - y1) + 2 * pad)
