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
        for page_num, page in enumerate(pdf.pages, start=1):
            text = (page.extract_text() or "").strip()

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

    return items


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
