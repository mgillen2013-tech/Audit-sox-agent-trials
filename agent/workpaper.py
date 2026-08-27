"""CY workpaper file generation -- the draft deliverable a reviewer opens.

Takes the structured results a control run already produced (ConclusionOutput
per test step, or the preserved failure info for a step that aborted) and
writes one workpaper file per control, matching the PY workpaper's file type:
PY was a PDF -> generate a PDF; PY was Excel -> generate .xlsx. This is a
CLEAN generated document with standard workpaper sections, not a cell-level
edit of the PY file -- per the design decision that a predictable layout
beats a fragile in-place edit of an arbitrary workbook.

Deliberately zero LLM calls: everything here is already in the conclusion
JSON. Generating the file costs nothing and is fully deterministic, so it
can run after every control run without touching the spending cap.

Every page/sheet is stamped DRAFT -- this output is a draft for human
review, never a finalized workpaper (same non-negotiable as the system
prompt).
"""

from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from agent.schemas import ConclusionOutput, EvidenceCitation, EvidenceItem

_DRAFT_BANNER = "DRAFT -- AI-prepared, pending human reviewer approval. Not a finalized workpaper."

_CONCLUSION_LABELS = {
    "satisfied": "Satisfied",
    "not_satisfied": "Not satisfied",
    "insufficient_evidence": "Insufficient evidence",
}


# ── Evidence exhibits (tickmark annotation) ────────────────────────────────
# Mirrors how a human workpaper points at evidence: the citation table gets a
# tickmark letter per citation (A, B, C...), and the cited source is embedded
# as an exhibit -- a rendered PDF page with red boxes drawn where the quoted
# text was found (tight boxes via text search; falls back to the extracted
# region's bbox; a scanned/screenshot page with no text layer becomes a
# full-page exhibit labeled with its letter). Excel-sourced citations get a
# text excerpt of the cited range instead of an image. All deterministic --
# zero LLM calls -- and every step degrades gracefully: if an exhibit can't
# be rendered, the citation row still stands, just without a picture.

_PDF_RENDER_DPI = 110
_MAX_EXHIBIT_WIDTH_PX = 640
_MAX_EXCERPT_ROWS = 12

_BBOX_RE = re.compile(r"p\.(\d+) \(bbox ([\d.\-]+),([\d.\-]+),([\d.\-]+),([\d.\-]+)\)")


def _tickmark_letters(n: int) -> list[str]:
    return [chr(65 + i) if i < 26 else str(i + 1) for i in range(n)]


def _clean_quote_for_search(quote: str) -> str | None:
    # The model's quote_or_summary may be a paraphrase or stitch pieces with
    # ellipses -- search for the longest literal-looking fragment.
    pieces = re.split(r"\.\.\.|…|\n", quote)
    best = max((re.sub(r"\s+", " ", p).strip() for p in pieces), key=len, default="")
    if len(best) < 12:
        return None
    return best[:60]


def _search_rects(page, quote: str) -> list[tuple[float, float, float, float]]:
    chunk = _clean_quote_for_search(quote)
    if not chunk:
        return []
    try:
        matches = page.search(chunk, regex=False, case=False)
    except Exception:  # noqa: BLE001 -- search is best-effort tightening only
        return []
    return [(m["x0"], m["top"], m["x1"], m["bottom"]) for m in matches[:3]]


def _item_rect(page, item: EvidenceItem) -> list[tuple[float, float, float, float]]:
    m = _BBOX_RE.search(item.location)
    if not m:
        return []
    x0, top, x1, bottom = (float(v) for v in m.groups()[1:])
    # A bbox covering (almost) the whole page isn't a pointer, it's the page
    # -- drawing a box around everything marks nothing.
    if (x1 - x0) * (bottom - top) >= 0.9 * page.width * page.height:
        return []
    return [(x0, top, x1, bottom)]


def _draw_letter(draw, pos: tuple[float, float], letter: str) -> None:
    from PIL import ImageFont

    try:
        font = ImageFont.load_default(size=20)
    except TypeError:  # older Pillow: no size kwarg
        font = ImageFont.load_default()
    x, y = pos
    draw.rectangle([x, y, x + 24, y + 24], fill="red")
    draw.text((x + 6, y + 2), letter, fill="white", font=font)


def _render_pdf_exhibit(
    pdf_path: Path, page_num: int, marks: list[tuple[str, EvidenceItem, str]]
) -> tuple[BytesIO, tuple[int, int]]:
    """marks: (tickmark letter, evidence item, quote to locate). Returns the
    annotated page as PNG bytes plus its pixel size.
    """
    import pdfplumber
    from PIL import ImageDraw

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_num - 1]
        pim = page.to_image(resolution=_PDF_RENDER_DPI)

        letter_positions: list[tuple[str, tuple[float, float] | None]] = []
        for letter, item, quote in marks:
            rects = _search_rects(page, quote) or _item_rect(page, item)
            for r in rects:
                pim.draw_rect(r, fill=None, stroke="red", stroke_width=3)
            letter_positions.append((letter, (rects[0][0], rects[0][1]) if rects else None))

        pil = pim.annotated.convert("RGB")
        draw = ImageDraw.Draw(pil)
        scale = _PDF_RENDER_DPI / 72.0
        unanchored = 0
        for letter, pos in letter_positions:
            if pos is not None:
                px = (max(pos[0] * scale - 28, 0), max(pos[1] * scale - 4, 0))
            else:
                # No locatable region (e.g. a screenshot page with no text
                # layer): the whole page is the exhibit -- stack the letters
                # in the top-left corner.
                px = (6 + unanchored * 30, 6)
                unanchored += 1
            _draw_letter(draw, px, letter)

    if pil.width > _MAX_EXHIBIT_WIDTH_PX:
        ratio = _MAX_EXHIBIT_WIDTH_PX / pil.width
        pil = pil.resize((_MAX_EXHIBIT_WIDTH_PX, int(pil.height * ratio)))
    buf = BytesIO()
    pil.save(buf, "PNG")
    buf.seek(0)
    return buf, (pil.width, pil.height)


def _excerpt_lines(item: EvidenceItem) -> list[str]:
    if item.extracted_table:
        rows = item.extracted_table[:_MAX_EXCERPT_ROWS]
        lines = [" | ".join(row) for row in rows]
        if len(item.extracted_table) > _MAX_EXCERPT_ROWS:
            lines.append(f"... ({len(item.extracted_table) - _MAX_EXCERPT_ROWS} more row(s) in the source)")
        return lines
    if item.extracted_text:
        return item.extracted_text.splitlines()[:_MAX_EXCERPT_ROWS]
    return []


def _build_step_exhibits(
    citations: list[EvidenceCitation],
    evidence_map: dict[str, EvidenceItem],
    support_dir: Path,
) -> tuple[list[str], list[tuple[str, BytesIO, tuple[int, int]]], list[tuple[str, list[str]]]]:
    """Returns (tickmark letter per citation, PDF exhibit images as
    (caption, png bytes, (w,h)), Excel/text excerpts as (caption, lines)).
    """
    letters = _tickmark_letters(len(citations))

    pdf_groups: dict[tuple[str, int], list[tuple[str, EvidenceItem, str]]] = {}
    excerpts: list[tuple[str, list[str]]] = []
    for letter, cit in zip(letters, citations):
        item = evidence_map.get(cit.evidence_id)
        if item is None:
            continue
        m = _BBOX_RE.search(item.location)
        if item.source_type in ("pdf_text", "pdf_table", "image_ocr") and m:
            key = (item.source_file, int(m.group(1)))
            pdf_groups.setdefault(key, []).append((letter, item, cit.quote_or_summary))
        else:
            lines = _excerpt_lines(item)
            if lines:
                excerpts.append((f"Exhibit {letter} — {item.location}", lines))

    images: list[tuple[str, BytesIO, tuple[int, int]]] = []
    for (source_file, page_num), marks in pdf_groups.items():
        try:
            buf, size = _render_pdf_exhibit(support_dir / source_file, page_num, marks)
        except Exception:  # noqa: BLE001 -- an unrenderable page must not sink the workpaper
            continue
        mark_list = ", ".join(letter for letter, _, _ in marks)
        images.append((f"Exhibit ({mark_list}) — {source_file} p.{page_num}", buf, size))

    return letters, images, excerpts


def workpaper_path_for(py_testing_filename: str, control_id: str, out_dir: Path) -> Path:
    """PY was a PDF -> .pdf out; anything else (the Excel family) -> .xlsx."""
    ext = ".pdf" if Path(py_testing_filename).suffix.lower() == ".pdf" else ".xlsx"
    safe_control = re.sub(r"[^\w.-]+", "_", control_id) or "control"
    return out_dir / f"{safe_control}_CY_Testing_DRAFT{ext}"


def build_workpaper(
    spec: dict[str, Any],
    results: dict[str, dict],
    py_testing_filename: str,
    out_dir: Path,
    support_dir: Path | None = None,
) -> Path:
    """Writes the control's CY workpaper and returns the written path.

    spec: the same control spec dict run_control uses (control_id,
    control_objective_ref, control_objective_text, test_steps).
    results: {test_step_id: {"conclusion", "audit_log"} | {"error", ...}} --
    exactly what iter_control_results yielded. Failed steps are documented
    in the workpaper as incomplete rather than silently omitted: a reviewer
    needs to see that a step has no conclusion yet, not a file that looks
    finished with a step missing.

    support_dir: the directory still holding spec["cy_support_files"] (the
    run's staging dir). When given, citations get tickmark letters and the
    cited sources are embedded as annotated exhibits. Extraction is
    deterministic, so re-extracting here reproduces the exact evidence_ids
    the run used -- same files, same order, same ids. When omitted (or if
    re-extraction fails) the workpaper is text-only, never broken.
    """
    out_path = workpaper_path_for(py_testing_filename, spec["control_id"], out_dir)
    step_texts = {s["test_step_id"]: s.get("test_step_text", "") for s in spec.get("test_steps", [])}

    evidence_map: dict[str, EvidenceItem] = {}
    if support_dir is not None:
        try:
            from agent.extraction import extract_many

            items = extract_many([support_dir / f for f in spec.get("cy_support_files", [])])
            evidence_map = {item.evidence_id: item for item in items}
        except Exception:  # noqa: BLE001 -- exhibits are an enhancement, not a precondition
            evidence_map = {}

    if out_path.suffix == ".pdf":
        _build_pdf(spec, results, step_texts, out_path, evidence_map, support_dir)
    else:
        _build_xlsx(spec, results, step_texts, out_path, evidence_map, support_dir)
    return out_path


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------

_HEADER_FILL = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")


def _sheet_title(test_step_id: str) -> str:
    # Excel forbids : \ / ? * [ ] in sheet names and caps them at 31 chars.
    return re.sub(r"[:\\/?*\[\]]", "_", test_step_id)[:31] or "step"


def _exhibit_sheet_title(test_step_id: str) -> str:
    # Same sanitizing, but leave room for the " - Exhibits" suffix within
    # Excel's 31-char sheet-name limit.
    base = re.sub(r"[:\\/?*\[\]]", "_", test_step_id)[:20] or "step"
    return f"{base} - Exhibits"


_SECTION_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
_SECTION_FONT = Font(bold=True, color="FFFFFF", size=11)
_CONCLUSION_COLORS = {
    "satisfied": "006100",
    "not_satisfied": "9C0006",
    "insufficient_evidence": "9C5700",
}


def _build_xlsx(
    spec: dict[str, Any],
    results: dict[str, dict],
    step_texts: dict[str, str],
    out_path: Path,
    evidence_map: dict[str, EvidenceItem],
    support_dir: Path | None,
) -> None:
    wb = openpyxl.Workbook()
    # Exhibit PNG buffers must stay alive until wb.save() -- openpyxl reads
    # image data at save time, not at add_image time.
    _live_buffers: list[BytesIO] = []

    ws = wb.active
    ws.title = "Summary"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 90

    next_row = 1

    def kv(label: str, value: str, bold_value: bool = False) -> None:
        # Explicit counter, not ws.max_row heuristics: a row whose only
        # content is in column B (like the DRAFT banner, with an empty
        # label) must still advance the cursor.
        nonlocal next_row
        ws.cell(next_row, 1, label).font = Font(bold=True)
        c = ws.cell(next_row, 2, value)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if bold_value:
            c.font = Font(bold=True)
        next_row += 1

    kv("", _DRAFT_BANNER, bold_value=True)
    kv("Control ID", spec["control_id"])
    kv("Control objective ref", spec.get("control_objective_ref", ""))
    kv("Control objective", spec.get("control_objective_text", ""))
    kv("", "")

    header_row = next_row
    summary_headers = [
        "Test step",
        "Conclusion",
        "Confidence",
        "Sample coverage",
        "IPE status",
        "Exceptions",
        "Open requests",
        "Detail sheet",
    ]
    for col, name in enumerate(summary_headers, 1):
        cell = ws.cell(header_row, col, name)
        cell.font = Font(bold=True)
        cell.fill = _HEADER_FILL
    for col, width in zip("ABCDEFGH", (22, 22, 12, 16, 16, 12, 14, 22)):
        ws.column_dimensions[col].width = width

    for test_step_id, result in results.items():
        r = ws.max_row + 1
        ws.cell(r, 1, test_step_id)
        if "error" in result:
            ws.cell(r, 2, "INCOMPLETE -- run did not finish").font = Font(bold=True, color="AA0000")
        else:
            conclusion: ConclusionOutput = result["conclusion"]
            verdict = ws.cell(r, 2, _CONCLUSION_LABELS.get(conclusion.conclusion, conclusion.conclusion))
            color = _CONCLUSION_COLORS.get(conclusion.conclusion)
            if color:
                verdict.font = Font(bold=True, color=color)
            ws.cell(r, 3, conclusion.confidence)
            if conclusion.sample_coverage:
                sc = conclusion.sample_coverage
                ws.cell(r, 4, f"{sc.total_found}/{sc.total_required}")
            ws.cell(r, 5, conclusion.ipe_completeness_accuracy_status)
            ws.cell(r, 6, len(conclusion.exceptions))
            ws.cell(r, 7, len(conclusion.additional_support_requests))
        ws.cell(r, 8, _sheet_title(test_step_id))

    for test_step_id, result in results.items():
        step_ws = wb.create_sheet(_sheet_title(test_step_id))
        letters, images, excerpts = _write_step_sheet(
            step_ws, test_step_id, step_texts.get(test_step_id, ""), result, evidence_map, support_dir
        )
        # Exhibits live on their own sheet: inline, a couple of full-page
        # renders pushed the conclusion, IPE status, and open requests
        # ~110 rows down the sheet, so a reviewer had to scroll past the
        # pictures to reach the answers. The step sheet stays readable
        # top-to-bottom; this is the appendix.
        if images or excerpts:
            ex_ws = wb.create_sheet(_exhibit_sheet_title(test_step_id))
            _write_exhibit_sheet(ex_ws, test_step_id, images, excerpts, _live_buffers)

    wb.save(out_path)


def _write_exhibit_sheet(
    ws,
    test_step_id: str,
    images: list[tuple[str, BytesIO, tuple[int, int]]],
    excerpts: list[tuple[str, list[str]]],
    live_buffers: list[BytesIO],
) -> None:
    ws.column_dimensions["A"].width = 30
    for col in "BCDEF":
        ws.column_dimensions[col].width = 26

    from openpyxl.drawing.image import Image as XLImage

    row = 1
    ws.cell(row, 1, f"Evidence exhibits — test step {test_step_id}").font = Font(bold=True, size=12)
    row += 1
    ws.cell(row, 1, "Tickmark letters match the 'Evidence cited' table on the test step sheet.").font = Font(
        italic=True
    )
    row += 2

    for caption, buf, (w, h) in images:
        ws.cell(row, 1, caption).font = Font(bold=True)
        row += 1
        img = XLImage(buf)
        img.width, img.height = w, h
        ws.add_image(img, f"A{row}")
        live_buffers.append(buf)
        row += int(h / 19) + 3  # default row height ~19px at 96dpi

    for caption, lines in excerpts:
        ws.cell(row, 1, caption).font = Font(bold=True)
        row += 1
        for line in lines:
            ws.cell(row, 1, line).alignment = Alignment(vertical="top")
            row += 1
        row += 1


def _write_step_sheet(
    ws,
    test_step_id: str,
    test_step_text: str,
    result: dict,
    evidence_map: dict[str, EvidenceItem],
    support_dir: Path | None,
) -> tuple[list[str], list[tuple[str, BytesIO, tuple[int, int]]], list[tuple[str, list[str]]]]:
    """Writes one test step's sheet, ordered so a reviewer reads answers
    first (conclusion, coverage, IPE, exceptions, open requests) before the
    supporting detail. Returns the exhibit material for the caller to place
    on a separate sheet -- inline images used to push the conclusion ~110
    rows down the page.
    """
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 100
    for col in "CDEF":
        ws.column_dimensions[col].width = 34

    row = 1

    def section(title: str) -> None:
        nonlocal row
        for col in range(1, 7):
            ws.cell(row, col).fill = _SECTION_FILL
        c = ws.cell(row, 1, title)
        c.font = _SECTION_FONT
        row += 2

    def put(label: str, value: str, *, bold: bool = False, color: str | None = None) -> None:
        nonlocal row
        ws.cell(row, 1, label).font = Font(bold=True)
        c = ws.cell(row, 2, value)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if bold or color:
            c.font = Font(bold=bold, color=color) if color else Font(bold=True)
        row += 1

    def put_block(text: str) -> None:
        nonlocal row
        c = ws.cell(row, 2, text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        row += 2

    def put_list(values: list[str], *, empty: str | None = None) -> None:
        nonlocal row
        if not values:
            if empty:
                ws.cell(row, 2, empty).font = Font(italic=True)
                row += 2
            return
        for v in values:
            c = ws.cell(row, 2, f"• {v}")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            row += 1
        row += 1

    ws.cell(row, 1, _DRAFT_BANNER).font = Font(bold=True, color="9C0006")
    row += 2
    put("Test step", test_step_id)
    put("Test step text", test_step_text)
    row += 1

    if "error" in result:
        section("RESULT — INCOMPLETE")
        put("Status", "INCOMPLETE -- run did not finish", bold=True, color="9C0006")
        put("Error", str(result["error"]))
        if "reason" in result:
            put("Abort reason", str(result["reason"]))
            put("Cost-weighted tokens used", f"{result.get('tokens_used', 0):,}")
        audit_log = result.get("audit_log") or []
        if audit_log:
            put("Tool calls before abort", str(len(audit_log)))
            row += 1
            section("SEARCHES ATTEMPTED BEFORE ABORT")
            put_list(
                [
                    str(e.input.get("query", ""))
                    for e in audit_log
                    if e.tool_name == "search_cy_support" and e.input.get("query")
                ]
            )
        return [], [], []

    conclusion: ConclusionOutput = result["conclusion"]

    letters: list[str] = []
    exhibit_images: list[tuple[str, BytesIO, tuple[int, int]]] = []
    excerpts: list[tuple[str, list[str]]] = []
    if conclusion.evidence_citations and evidence_map and support_dir is not None:
        letters, exhibit_images, excerpts = _build_step_exhibits(
            conclusion.evidence_citations, evidence_map, support_dir
        )

    # ── Answers first ──────────────────────────────────────────────────
    section("CONCLUSION")
    put(
        "Conclusion",
        _CONCLUSION_LABELS.get(conclusion.conclusion, conclusion.conclusion),
        bold=True,
        color=_CONCLUSION_COLORS.get(conclusion.conclusion),
    )
    put("Confidence", conclusion.confidence)
    put("Confidence rationale", conclusion.confidence_rationale)
    if conclusion.sample_coverage:
        sc = conclusion.sample_coverage
        put(
            "Sample coverage",
            f"{sc.total_found} of {sc.total_required} ({sc.coverage_pct}%)"
            + (f"; missing: {', '.join(sc.missing)}" if sc.missing else ""),
        )
    put("IPE status", conclusion.ipe_completeness_accuracy_status)
    row += 1

    section("EXCEPTIONS")
    put_list(conclusion.exceptions, empty="None noted.")

    section("ADDITIONAL SUPPORT REQUESTED")
    put_list(conclusion.additional_support_requests, empty="None — testing is complete on the evidence provided.")

    # ── Supporting detail ──────────────────────────────────────────────
    section("DOCUMENTATION")
    put_block(conclusion.narrative)

    section("PROCEDURES PERFORMED")
    put_list(conclusion.procedures_performed)

    if conclusion.ipe_completeness_accuracy_evidence:
        section("IPE COMPLETENESS & ACCURACY EVIDENCE")
        put_list(conclusion.ipe_completeness_accuracy_evidence)

    if conclusion.evidence_citations:
        section("EVIDENCE CITED")
        headers = ["Tickmark", "Evidence ID", "Source file", "Location", "Quote / summary", "Relevance"]
        for col, name in enumerate(headers, 1):
            cell = ws.cell(row, col, name)
            cell.font = Font(bold=True)
            cell.fill = _HEADER_FILL
        row += 1
        for i, cit in enumerate(conclusion.evidence_citations):
            letter = letters[i] if i < len(letters) else ""
            for col, value in enumerate(
                [letter, cit.evidence_id, cit.source_file, cit.location, cit.quote_or_summary, cit.relevance], 1
            ):
                c = ws.cell(row, col, value)
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if col == 1:
                    c.font = Font(bold=True, color="AA0000")
            row += 1
        row += 1
        if exhibit_images or excerpts:
            ws.cell(row, 1, f"Exhibits for these citations: see the '{_exhibit_sheet_title(test_step_id)}' sheet.").font = Font(
                italic=True
            )
            row += 2

    section("PREPARED BY")
    md = conclusion.model_metadata
    put("Prepared by", f"CY testing agent ({md.model}, prompt {md.prompt_version})")
    put("Timestamp", md.timestamp)
    put("Tool calls", str(md.tool_call_count))

    return letters, exhibit_images, excerpts


# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------


def _build_pdf(
    spec: dict[str, Any],
    results: dict[str, dict],
    step_texts: dict[str, str],
    out_path: Path,
    evidence_map: dict[str, EvidenceItem],
    support_dir: Path | None,
) -> None:
    # Imported here, not module-level: reportlab is only needed when PY was a
    # PDF, and keeping the import local means an Excel-only user without it
    # installed can still generate xlsx workpapers.
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    styles = getSampleStyleSheet()
    body = styles["BodyText"]
    h1, h2 = styles["Heading1"], styles["Heading2"]
    draft = Paragraph(f"<b>{_DRAFT_BANNER}</b>", body)

    story: list[Any] = [draft, Spacer(1, 8)]
    story.append(Paragraph(f"Control {spec['control_id']} -- CY Testing", h1))
    story.append(
        Paragraph(
            f"<b>Objective ({spec.get('control_objective_ref', '')}):</b> "
            f"{spec.get('control_objective_text', '')}",
            body,
        )
    )
    story.append(Spacer(1, 12))

    def bullet_list(label: str, values: list[str]) -> None:
        if not values:
            return
        story.append(Paragraph(f"<b>{label}:</b>", body))
        for v in values:
            story.append(Paragraph(f"• {v}", body))
        story.append(Spacer(1, 6))

    for test_step_id, result in results.items():
        story.append(Paragraph(f"Test step {test_step_id}", h2))
        step_text = step_texts.get(test_step_id, "")
        if step_text:
            story.append(Paragraph(f"<b>Test step:</b> {step_text}", body))

        if "error" in result:
            story.append(Paragraph("<b>Status: INCOMPLETE -- run did not finish</b>", body))
            story.append(Paragraph(f"Error: {result['error']}", body))
            if "reason" in result:
                story.append(
                    Paragraph(
                        f"Abort reason: {result['reason']} ({result.get('tokens_used', 0):,} tokens used)", body
                    )
                )
            story.append(Spacer(1, 12))
            continue

        conclusion: ConclusionOutput = result["conclusion"]
        story.append(
            Paragraph(
                f"<b>Conclusion:</b> {_CONCLUSION_LABELS.get(conclusion.conclusion, conclusion.conclusion)} "
                f"(confidence: {conclusion.confidence})",
                body,
            )
        )
        story.append(Paragraph(f"<b>Documentation:</b> {conclusion.narrative}", body))
        story.append(Spacer(1, 6))
        bullet_list("Procedures performed", conclusion.procedures_performed)

        letters: list[str] = []
        exhibit_images: list[tuple[str, BytesIO, tuple[int, int]]] = []
        excerpts: list[tuple[str, list[str]]] = []
        if conclusion.evidence_citations and evidence_map and support_dir is not None:
            letters, exhibit_images, excerpts = _build_step_exhibits(
                conclusion.evidence_citations, evidence_map, support_dir
            )

        if conclusion.evidence_citations:
            rows = [["Tickmark", "Evidence ID", "Source file", "Location", "Quote / summary"]] + [
                [
                    Paragraph(f"<b><font color='red'>{letters[i] if i < len(letters) else ''}</font></b>", body),
                    Paragraph(cit.evidence_id, body),
                    Paragraph(cit.source_file, body),
                    Paragraph(cit.location, body),
                    Paragraph(cit.quote_or_summary, body),
                ]
                for i, cit in enumerate(conclusion.evidence_citations)
            ]
            table = Table(rows, colWidths=[0.7 * inch, 0.8 * inch, 1.4 * inch, 1.6 * inch, 2.5 * inch])
            table.setStyle(
                TableStyle(
                    [
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ]
                )
            )
            story.append(table)
            story.append(Spacer(1, 6))

        if exhibit_images or excerpts:
            from reportlab.platypus import Image as RLImage

            story.append(Paragraph("<b>Evidence exhibits:</b>", body))
            for caption, buf, (w, h) in exhibit_images:
                story.append(Paragraph(f"<i>{caption}</i>", body))
                display_w = min(6.5 * inch, w)
                story.append(RLImage(buf, width=display_w, height=h * (display_w / w)))
                story.append(Spacer(1, 8))
            for caption, lines in excerpts:
                story.append(Paragraph(f"<i>{caption}</i>", body))
                for line in lines:
                    story.append(Paragraph(line, body))
                story.append(Spacer(1, 8))

        if conclusion.sample_coverage:
            sc = conclusion.sample_coverage
            story.append(
                Paragraph(
                    f"<b>Sample coverage:</b> {sc.total_found} of {sc.total_required} ({sc.coverage_pct}%)"
                    + (f"; missing: {', '.join(sc.missing)}" if sc.missing else ""),
                    body,
                )
            )
        story.append(Paragraph(f"<b>IPE status:</b> {conclusion.ipe_completeness_accuracy_status}", body))
        bullet_list("IPE C&A evidence", conclusion.ipe_completeness_accuracy_evidence)
        bullet_list("Exceptions", conclusion.exceptions)
        bullet_list("Additional support requested", conclusion.additional_support_requests)

        md = conclusion.model_metadata
        story.append(
            Paragraph(
                f"<i>Prepared by CY testing agent ({md.model}, prompt {md.prompt_version}) -- "
                f"{md.timestamp}; {md.tool_call_count} tool calls.</i>",
                body,
            )
        )
        story.append(Spacer(1, 16))

    SimpleDocTemplate(str(out_path), pagesize=letter).build(story)
