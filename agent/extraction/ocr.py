"""Vision OCR for image-only PDF pages, using the same Claude client the
tool loop already has.

Why vision rather than Tesseract: the target user runs this on Windows,
where Tesseract means a separate binary install and PATH setup, and the
documents that land in this bucket (scanned invoices, E1 screenshots,
payment remittances) are exactly the layout-heavy, table-ish material
Tesseract handles worst. The Foundry connection is already configured, so
this is zero extra install and materially more accurate.

This is the ONLY place extraction talks to a model, and it stays honest
about that: OCR'd text is marked with extraction_confidence 0.8 (not the
1.0 of natively-extracted text) because it IS a transcription, not a
direct read of structured data, and a reviewer should be able to tell
which citations rest on it.

Cost is bounded and small: it runs once per image-only page, before the
tool loop starts -- not per turn -- and a page is roughly 1-2K input
tokens. Pages that already have a text layer never reach this module.
Every failure mode leaves the original placeholder EvidenceItem intact
(extraction_confidence 0.0, extracted_text None), so a failed OCR
degrades to exactly today's behavior rather than breaking a run.
"""

from __future__ import annotations

import base64
import re
from io import BytesIO
from pathlib import Path
from typing import Any

from agent.schemas import EvidenceItem

_PAGE_RE = re.compile(r" p\.(\d+) ")
_OCR_DPI = 150
_MAX_EDGE_PX = 1_600  # keeps one page well under the API's per-image limits

OCR_CONFIDENCE = 0.8

# Ceiling on how many scanned pages one run will transcribe. OCR happens
# before the tool loop, so run_test_step's spending cap cannot see or stop
# it -- without this, a control whose support happens to be a 300-page
# scanned bundle would spend unbounded money before testing even began.
# Pages past the limit keep their placeholder (flagged unreadable), which
# is the honest outcome and the one the agent already knows how to handle
# via request_additional_support. At ~1-2K tokens a page this is the
# largest cost incurred BEFORE any testing starts, so the ceiling sits
# below what a whole healthy run costs -- 40 pages could add ~80K on its
# own, more than the run it precedes.
MAX_OCR_PAGES = 15

_PROMPT = """\
Transcribe this document page verbatim for use as audit evidence.

Rules:
- Transcribe exactly what is written. Never infer, correct, complete, or
  guess at a value. This is evidence in a financial control test -- an
  invented number is worse than a missing one.
- Preserve the reading order and structure. Render tabular areas as rows
  with " | " between cells.
- If a value is cut off, smudged, or genuinely unreadable, write
  [illegible] in its place rather than a best guess.
- Output only the transcription. No preamble, no summary, no commentary.
"""


def _page_number(location: str) -> int:
    m = _PAGE_RE.search(location)
    return int(m.group(1)) if m else 1


_XL_IMAGE_RE = re.compile(r"^(.*)!image(\d+)$")


def _downscale(pil: Any) -> bytes:
    if max(pil.width, pil.height) > _MAX_EDGE_PX:
        ratio = _MAX_EDGE_PX / max(pil.width, pil.height)
        pil = pil.resize((int(pil.width * ratio), int(pil.height * ratio)))
    buf = BytesIO()
    pil.convert("RGB").save(buf, "PNG")
    return buf.getvalue()


def _render_pdf_page_png(pdf_path: Path, page_num: int) -> bytes:
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_num - 1]
        return _downscale(page.to_image(resolution=_OCR_DPI).annotated)


def _render_excel_image_png(xlsx_path: Path, sheet_name: str, image_idx: int) -> bytes:
    """Bytes of one image pasted onto a worksheet (1-based index)."""
    import openpyxl
    from PIL import Image as PILImage

    wb = openpyxl.load_workbook(xlsx_path)
    image = wb[sheet_name]._images[image_idx - 1]
    return _downscale(PILImage.open(BytesIO(image._data())))


def _render_source_png(item: EvidenceItem, source_dir: Path) -> bytes:
    """Dispatch on where the image actually lives: a page of a scanned PDF,
    or a screenshot pasted onto a worksheet.
    """
    path = source_dir / item.source_file
    m = _XL_IMAGE_RE.match(item.location)
    if m and path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return _render_excel_image_png(path, m.group(1), int(m.group(2)))
    return _render_pdf_page_png(path, _page_number(item.location))


def _transcribe(png: bytes, client: Any, model: str) -> tuple[str, int]:
    message = {
        "role": "user",
        "content": [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": base64.standard_b64encode(png).decode("ascii"),
                },
            },
            {"type": "text", "text": _PROMPT},
        ],
    }
    kwargs = {"model": model, "max_tokens": 4096, "messages": [message]}
    # Same dual-shape client handling as agent/loop.py's _call_model: the
    # real SDK exposes client.messages.create, test doubles a bare
    # create_message.
    response = (
        client.messages.create(**kwargs) if hasattr(client, "messages") else client.create_message(**kwargs)
    )
    text = "\n".join(
        block.text for block in response.content if getattr(block, "type", None) == "text"
    ).strip()
    # Same cost-weighted units the tool loop's spending cap uses, so OCR
    # spend is reported on the same scale as everything else rather than in
    # a second, incomparable currency.
    from agent.loop import _usage_tokens

    return text, _usage_tokens(getattr(response, "usage", None))


def ocr_image_items(
    items: list[EvidenceItem],
    source_dir: Path,
    client: Any,
    model: str,
    on_page: Any = None,
    max_pages: int = MAX_OCR_PAGES,
) -> tuple[list[EvidenceItem], int, int]:
    """Fills in extracted_text for every image_ocr item by transcribing its
    page. Returns (new item list, count transcribed, cost-weighted tokens
    spent).

    The token total is returned rather than swallowed because this runs
    OUTSIDE run_test_step's spending cap -- it is per-control, before the
    tool loop starts, so the cap can neither see nor stop it. A control
    with many scanned pages could otherwise spend real money with nothing
    on screen accounting for it.

    Items are replaced, not mutated, and non-image items pass through
    untouched, so the caller's evidence_ids and ordering are preserved
    exactly -- that matters because agent/workpaper.py re-extracts to map
    evidence_ids back to source regions for exhibits.

    on_page: optional callback(source_file, page_num, ok, tokens_so_far)
    for progress and running-cost reporting while pages are transcribed.
    """
    out: list[EvidenceItem] = []
    transcribed = 0
    tokens_used = 0

    for item in items:
        if item.source_type != "image_ocr" or item.extracted_text:
            out.append(item)
            continue

        if transcribed >= max_pages:
            # Past the ceiling: leave the placeholder rather than keep
            # spending. The page is still visible to the agent as
            # unreadable evidence, not silently dropped.
            out.append(item)
            continue

        page_num = _page_number(item.location)
        try:
            png = _render_source_png(item, source_dir)
            text, tokens = _transcribe(png, client, model)
            tokens_used += tokens
        except Exception:  # noqa: BLE001 -- a failed OCR degrades to the placeholder, never breaks the run
            if on_page is not None:
                on_page(item.source_file, page_num, False, tokens_used)
            out.append(item)
            continue

        if not text:
            if on_page is not None:
                on_page(item.source_file, page_num, False, tokens_used)
            out.append(item)
            continue

        kind = (
            "an image pasted into the workbook"
            if _XL_IMAGE_RE.match(item.location)
            else "a scanned page"
        )
        out.append(
            item.model_copy(
                update={
                    "extracted_text": f"[OCR transcription of {kind}]\n{text}",
                    "extraction_confidence": OCR_CONFIDENCE,
                }
            )
        )
        transcribed += 1
        if on_page is not None:
            on_page(item.source_file, page_num, True, tokens_used)

    return out, transcribed, tokens_used
