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
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from agent.schemas import ConclusionOutput

_DRAFT_BANNER = "DRAFT -- AI-prepared, pending human reviewer approval. Not a finalized workpaper."

_CONCLUSION_LABELS = {
    "satisfied": "Satisfied",
    "not_satisfied": "Not satisfied",
    "insufficient_evidence": "Insufficient evidence",
}


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
) -> Path:
    """Writes the control's CY workpaper and returns the written path.

    spec: the same control spec dict run_control uses (control_id,
    control_objective_ref, control_objective_text, test_steps).
    results: {test_step_id: {"conclusion", "audit_log"} | {"error", ...}} --
    exactly what iter_control_results yielded. Failed steps are documented
    in the workpaper as incomplete rather than silently omitted: a reviewer
    needs to see that a step has no conclusion yet, not a file that looks
    finished with a step missing.
    """
    out_path = workpaper_path_for(py_testing_filename, spec["control_id"], out_dir)
    step_texts = {s["test_step_id"]: s.get("test_step_text", "") for s in spec.get("test_steps", [])}
    if out_path.suffix == ".pdf":
        _build_pdf(spec, results, step_texts, out_path)
    else:
        _build_xlsx(spec, results, step_texts, out_path)
    return out_path


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------

_HEADER_FILL = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")


def _sheet_title(test_step_id: str) -> str:
    # Excel forbids : \ / ? * [ ] in sheet names and caps them at 31 chars.
    return re.sub(r"[:\\/?*\[\]]", "_", test_step_id)[:31] or "step"


def _build_xlsx(
    spec: dict[str, Any], results: dict[str, dict], step_texts: dict[str, str], out_path: Path
) -> None:
    wb = openpyxl.Workbook()

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
    for col, name in enumerate(["Test step", "Conclusion", "Confidence", "Sample coverage", "Detail sheet"], 1):
        cell = ws.cell(header_row, col, name)
        cell.font = Font(bold=True)
        cell.fill = _HEADER_FILL
    for col, width in zip("ABCDE", (28, 24, 12, 18, 16)):
        ws.column_dimensions[col].width = width

    for test_step_id, result in results.items():
        r = ws.max_row + 1
        ws.cell(r, 1, test_step_id)
        if "error" in result:
            ws.cell(r, 2, "INCOMPLETE -- run did not finish").font = Font(bold=True, color="AA0000")
        else:
            conclusion: ConclusionOutput = result["conclusion"]
            ws.cell(r, 2, _CONCLUSION_LABELS.get(conclusion.conclusion, conclusion.conclusion))
            ws.cell(r, 3, conclusion.confidence)
            if conclusion.sample_coverage:
                sc = conclusion.sample_coverage
                ws.cell(r, 4, f"{sc.total_found}/{sc.total_required}")
        ws.cell(r, 5, _sheet_title(test_step_id))

    for test_step_id, result in results.items():
        step_ws = wb.create_sheet(_sheet_title(test_step_id))
        _write_step_sheet(step_ws, test_step_id, step_texts.get(test_step_id, ""), result)

    wb.save(out_path)


def _write_step_sheet(ws, test_step_id: str, test_step_text: str, result: dict) -> None:
    ws.column_dimensions["A"].width = 26
    for col in "BCDE":
        ws.column_dimensions[col].width = 40

    row = 1

    def put(label: str, value: str, *, bold: bool = False) -> None:
        nonlocal row
        ws.cell(row, 1, label).font = Font(bold=True)
        c = ws.cell(row, 2, value)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if bold:
            c.font = Font(bold=True)
        row += 1

    def put_list(label: str, values: list[str]) -> None:
        nonlocal row
        if not values:
            return
        ws.cell(row, 1, label).font = Font(bold=True)
        for v in values:
            c = ws.cell(row, 2, f"- {v}")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            row += 1
        row += 1

    put("", _DRAFT_BANNER, bold=True)
    put("Test step", test_step_id)
    put("Test step text", test_step_text)
    row += 1

    if "error" in result:
        put("Status", "INCOMPLETE -- run did not finish", bold=True)
        put("Error", str(result["error"]))
        if "reason" in result:
            put("Abort reason", str(result["reason"]))
            put("Tokens used", f"{result.get('tokens_used', 0):,}")
        audit_log = result.get("audit_log") or []
        if audit_log:
            put("Tool calls before abort", str(len(audit_log)))
            put_list(
                "Searches attempted",
                [
                    str(e.input.get("query", ""))
                    for e in audit_log
                    if e.tool_name == "search_cy_support" and e.input.get("query")
                ],
            )
        return

    conclusion: ConclusionOutput = result["conclusion"]
    put("Conclusion", _CONCLUSION_LABELS.get(conclusion.conclusion, conclusion.conclusion), bold=True)
    put("Confidence", f"{conclusion.confidence} -- {conclusion.confidence_rationale}")
    row += 1
    put("Documentation", conclusion.narrative)
    row += 1
    put_list("Procedures performed", conclusion.procedures_performed)

    if conclusion.evidence_citations:
        ws.cell(row, 1, "Evidence cited").font = Font(bold=True)
        row += 1
        for col, name in enumerate(["Evidence ID", "Source file", "Location", "Quote / summary", "Relevance"], 1):
            cell = ws.cell(row, col, name)
            cell.font = Font(bold=True)
            cell.fill = _HEADER_FILL
        row += 1
        for cit in conclusion.evidence_citations:
            for col, value in enumerate(
                [cit.evidence_id, cit.source_file, cit.location, cit.quote_or_summary, cit.relevance], 1
            ):
                ws.cell(row, col, value).alignment = Alignment(wrap_text=True, vertical="top")
            row += 1
        row += 1

    if conclusion.sample_coverage:
        sc = conclusion.sample_coverage
        put(
            "Sample coverage",
            f"{sc.total_found} of {sc.total_required} ({sc.coverage_pct}%)"
            + (f"; missing: {', '.join(sc.missing)}" if sc.missing else ""),
        )
    put("IPE status", conclusion.ipe_completeness_accuracy_status)
    put_list("IPE C&A evidence", conclusion.ipe_completeness_accuracy_evidence)
    put_list("Exceptions", conclusion.exceptions)
    put_list("Additional support requested", conclusion.additional_support_requests)

    row += 1
    md = conclusion.model_metadata
    put("Prepared by", f"CY testing agent ({md.model}, prompt {md.prompt_version}) -- {md.timestamp}")
    put("Tool calls", str(md.tool_call_count))


# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------


def _build_pdf(
    spec: dict[str, Any], results: dict[str, dict], step_texts: dict[str, str], out_path: Path
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

        if conclusion.evidence_citations:
            rows = [["Evidence ID", "Source file", "Location", "Quote / summary"]] + [
                [
                    Paragraph(cit.evidence_id, body),
                    Paragraph(cit.source_file, body),
                    Paragraph(cit.location, body),
                    Paragraph(cit.quote_or_summary, body),
                ]
                for cit in conclusion.evidence_citations
            ]
            table = Table(rows, colWidths=[0.8 * inch, 1.5 * inch, 1.7 * inch, 3.0 * inch])
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
