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


def _render_page_png(pdf_path: Path, page_num: int) -> bytes:
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_num - 1]
        pil = page.to_image(resolution=_OCR_DPI).annotated.convert("RGB")

    if max(pil.width, pil.height) > _MAX_EDGE_PX:
        ratio = _MAX_EDGE_PX / max(pil.width, pil.height)
        pil = pil.resize((int(pil.width * ratio), int(pil.height * ratio)))

    buf = BytesIO()
    pil.save(buf, "PNG")
    return buf.getvalue()


def _transcribe(png: bytes, client: Any, model: str) -> str:
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
    return "\n".join(
        block.text for block in response.content if getattr(block, "type", None) == "text"
    ).strip()


def ocr_image_items(
    items: list[EvidenceItem],
    source_dir: Path,
    client: Any,
    model: str,
    on_page: Any = None,
) -> tuple[list[EvidenceItem], int]:
    """Fills in extracted_text for every image_ocr item by transcribing its
    page. Returns (new item list, count actually transcribed).

    Items are replaced, not mutated, and non-image items pass through
    untouched, so the caller's evidence_ids and ordering are preserved
    exactly -- that matters because agent/workpaper.py re-extracts to map
    evidence_ids back to source regions for exhibits.

    on_page: optional callback(source_file, page_num, ok: bool) for
    progress reporting while pages are being transcribed.
    """
    out: list[EvidenceItem] = []
    transcribed = 0

    for item in items:
        if item.source_type != "image_ocr" or item.extracted_text:
            out.append(item)
            continue

        page_num = _page_number(item.location)
        try:
            png = _render_page_png(source_dir / item.source_file, page_num)
            text = _transcribe(png, client, model)
        except Exception:  # noqa: BLE001 -- a failed OCR degrades to the placeholder, never breaks the run
            if on_page is not None:
                on_page(item.source_file, page_num, False)
            out.append(item)
            continue

        if not text:
            if on_page is not None:
                on_page(item.source_file, page_num, False)
            out.append(item)
            continue

        out.append(
            item.model_copy(
                update={
                    "extracted_text": f"[OCR transcription of a scanned page]\n{text}",
                    "extraction_confidence": OCR_CONFIDENCE,
                }
            )
        )
        transcribed += 1
        if on_page is not None:
            on_page(item.source_file, page_num, True)

    return out, transcribed
