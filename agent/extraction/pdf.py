"""PDF extraction: pdfplumber -> EvidenceItem list.

Native-text PDFs (the common case for CY support) are extracted directly:
one ``pdf_text`` EvidenceItem per page and one ``pdf_table`` EvidenceItem per
detected table, each carrying a page number and bounding box so citations
and the review UI can point at the exact region.

Scanned pages -- where pdfplumber finds no extractable text but the page
clearly has content -- are NOT silently dropped and are NOT run through OCR
here (no OCR engine is wired up yet; see the design doc's extraction
section). Instead each such page becomes an ``image_ocr`` EvidenceItem with
extraction_confidence=0.0 and extracted_text=None. This is deliberate: a
low-confidence placeholder is what lets the agent legitimately request
additional support on an illegible source, instead of that page just
vanishing from the evidence set with no trace it was ever there.
"""

from __future__ import annotations

import itertools
from pathlib import Path
from typing import Iterator

import pdfplumber

from agent.schemas import EvidenceItem

_MIN_CHARS_FOR_NATIVE_TEXT = 20  # below this, treat the page as likely-scanned


def extract_pdf(path: str | Path) -> list[EvidenceItem]:
    path = Path(path)
    filename = path.name
    counter = itertools.count(1)

    items: list[EvidenceItem] = []
    with pdfplumber.open(path) as pdf:
        xfa = _xfa_form_data(pdf)
        for page_num, page in enumerate(pdf.pages, start=1):
            text = (page.extract_text() or "").strip()
            # A fillable form keeps what was typed into it in its fields,
            # not in the page text: without this the agent saw "Vendor:
            # Amount: Approved:" and none of the answers.
            fields = _form_fields(page)
            if fields:
                text = (text + "\n\nFilled-in form fields:\n" + "\n".join(fields)).strip()

            if len(text) < _MIN_CHARS_FOR_NATIVE_TEXT:
                items.append(_scanned_placeholder(filename, page, page_num, counter))
                continue

            items.append(
                EvidenceItem(
                    evidence_id=f"ev_{next(counter):04d}",
                    source_file=filename,
                    source_type="pdf_text",
                    location=f"{filename} p.{page_num} (bbox {_bbox_str(page.bbox)})",
                    extracted_text=text,
                    extracted_table=None,
                    extraction_confidence=1.0,
                    preview_ref=f"{filename} p.{page_num}",
                )
            )
            items.extend(_extract_tables(page, filename, page_num, counter))

    if xfa:
        # A dynamic (XFA) form: its pages are often only a "please wait"
        # placeholder outside Adobe Reader, but the answers are stored in
        # the file as data, and are read from there.
        items.append(
            EvidenceItem(
                evidence_id=f"ev_{next(counter):04d}",
                source_file=filename,
                source_type="pdf_text",
                location=f"{filename} (form data)",
                extracted_text="Filled-in form data (dynamic form):\n" + "\n".join(xfa),
                extracted_table=None,
                extraction_confidence=1.0,
                preview_ref=f"{filename} p.1",
            )
        )

    return items


def _text(value) -> str:
    """A PDF field value as plain text: strings may be bytes in either
    PDF encoding, names (checkbox states) arrive as /Yes or /Off."""
    if value is None:
        return ""
    if isinstance(value, list):
        return ", ".join(t for t in (_text(v) for v in value) if t)
    if isinstance(value, bytes):
        if value[:2] in (b"\xfe\xff", b"\xff\xfe"):
            return value.decode("utf-16", errors="replace").strip()
        return value.decode("latin-1", errors="replace").strip()
    name = getattr(value, "name", None)
    if name is not None:
        return "unchecked" if name == "Off" else ("checked" if name in ("Yes", "On") else str(name))
    return str(value).strip()


def _field_label(page, x0: float, top: float, bottom: float) -> str:
    """The printed words just left of a field on the same line, else the
    line just above it -- what a person reads as the field's label."""
    words = page.extract_words()
    mid = (top + bottom) / 2
    left = [w for w in words if w["x1"] <= x0 + 1 and w["top"] <= mid <= w["bottom"]]
    if left:
        left.sort(key=lambda w: w["x0"])
        return " ".join(w["text"] for w in left)[-60:].strip()
    above = [w for w in words if w["bottom"] <= top + 1 and top - w["bottom"] < 20
             and w["x0"] < x0 + 150 and w["x1"] > x0 - 10]
    if above:
        line = max(w["bottom"] for w in above)
        return " ".join(w["text"] for w in sorted(above, key=lambda w: w["x0"]) if abs(w["bottom"] - line) < 2)[-60:]
    return ""


def _form_fields(page) -> list[str]:
    """'Label: value' for every filled-in field on the page. Empty fields
    are left out; a field with no printed label is shown by its name."""
    lines = []
    for annot in page.annots or []:
        data = annot.get("data") or {}
        if str(getattr(data.get("Subtype"), "name", data.get("Subtype"))) != "Widget":
            continue
        parent = data.get("Parent")
        parent = parent.resolve() if hasattr(parent, "resolve") else (parent or {})
        value = data.get("V", parent.get("V") if isinstance(parent, dict) else None)
        if value is None and data.get("AS") is not None:
            value = data.get("AS")
        shown = _text(value)
        if not shown:
            continue
        name = _text(data.get("T") or (parent.get("T") if isinstance(parent, dict) else None))
        label = _field_label(page, annot["x0"], annot["top"], annot["bottom"]) or name or "field"
        lines.append(f"{label} {shown}" if label.endswith(":") else f"{label}: {shown}")
    return lines


def _xfa_form_data(pdf) -> list[str]:
    """'name: value' for each filled-in value of a dynamic (XFA) form, read
    from the form's data packet. [] for any other PDF."""
    import xml.etree.ElementTree as ET

    from pdfminer.pdftypes import resolve1

    try:
        acro = resolve1(pdf.doc.catalog.get("AcroForm"))
        xfa = resolve1(acro.get("XFA")) if isinstance(acro, dict) else None
        if xfa is None:
            return []
        if isinstance(xfa, list):
            parts = [resolve1(xfa[i + 1]) for i in range(0, len(xfa) - 1, 2)
                     if _text(resolve1(xfa[i])) == "datasets"]
        else:
            parts = [xfa]
        lines: list[str] = []
        for part in parts:
            raw = part.get_data() if hasattr(part, "get_data") else bytes(part)
            root = ET.fromstring(raw)
            for el in root.iter():
                if len(el) == 0 and el.text and el.text.strip():
                    tag = el.tag.split("}")[-1]
                    lines.append(f"{tag}: {el.text.strip()}")
        return lines[:400]
    except Exception:  # noqa: BLE001 -- a malformed packet just means no form data
        return []


def _extract_tables(page, filename: str, page_num: int, counter: Iterator[int]) -> list[EvidenceItem]:
    items: list[EvidenceItem] = []
    for table in page.find_tables():
        rows = table.extract()
        if not rows:
            continue
        normalized = [["" if cell is None else str(cell) for cell in row] for row in rows]
        items.append(
            EvidenceItem(
                evidence_id=f"ev_{next(counter):04d}",
                source_file=filename,
                source_type="pdf_table",
                location=f"{filename} p.{page_num} (bbox {_bbox_str(table.bbox)})",
                extracted_text=None,
                extracted_table=normalized,
                extraction_confidence=1.0,
                preview_ref=f"{filename} p.{page_num}",
            )
        )
    return items


def _scanned_placeholder(filename: str, page, page_num: int, counter: Iterator[int]) -> EvidenceItem:
    return EvidenceItem(
        evidence_id=f"ev_{next(counter):04d}",
        source_file=filename,
        source_type="image_ocr",
        location=f"{filename} p.{page_num} (bbox {_bbox_str(page.bbox)})",
        extracted_text=None,
        extracted_table=None,
        extraction_confidence=0.0,
        preview_ref=f"{filename} p.{page_num}",
    )


def _bbox_str(bbox: tuple[float, float, float, float]) -> str:
    return ",".join(str(round(v)) for v in bbox)
