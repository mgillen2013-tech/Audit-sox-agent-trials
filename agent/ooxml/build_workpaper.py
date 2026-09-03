"""
Top-level entry point. build_workpaper(request) turns a validated
WorkpaperRequest into a finished .xlsx, doing all the OOXML/EMU/OCR work
described in README.md.

IMPORTANT ASSUMPTION: this assumes the *template* passed in is a workbook
that already has the four-tab shape described in README.md, with cell
styles s="1".."14" etc. already defined for: text cells, date cells, number
cells, and the IA-Calculation box's header/line/total rows (see
STYLE_INDEX_* constants below). Those style indices are template-specific --
if you start from a different base template, open it, find the equivalent
style ids (Sample!B19 etc.) and update the constants, or (cleaner) generate
new styles.xml entries programmatically. This package doesn't do that
automatically because "what should the IA Calculation box look like" is a
one-time template decision, not a per-sample one.
"""
from __future__ import annotations
import os
from PIL import Image

from agent.ooxml.models import WorkpaperRequest, SampleItem, PopulationRow, RawDataRow
from agent.ooxml import ooxml_utils as ox
from agent.ooxml import drawing_builder as db

# Style indices lifted from T0_SS_PTP_AP_156_01__PY_Testing_wp.xlsx's
# styles.xml. Verify these against `Sample.xml` in your own template before
# reusing on a different base file.
STYLE_TEXT = 1        # generic text cell (matches header font)
STYLE_DATE = 2
STYLE_NUMBER = 3
STYLE_NUMBER_ALT = 5  # used for the Invoice Payment Amount column in the
                       # original template; cosmetic difference from STYLE_NUMBER
STYLE_IA_DIVIDER = 7        # thin rule row above the IA Calculation box
STYLE_IA_HEADER_LABEL = 9   # "IA Calculation" cell
STYLE_IA_HEADER_BLANK = 10
STYLE_IA_LINE_LABEL = 11
STYLE_IA_LINE_AMOUNT = 12
STYLE_IA_TOTAL_LABEL = 13
STYLE_IA_TOTAL_AMOUNT = 14


def build_sample_sheet_xml(sheet_xml_path: str, sample_items: list[SampleItem],
                            sst: ox.SharedStrings) -> None:
    """Overwrite the Sample sheet's data rows (raw data rows + IA Calculation
    block(s)) in place. Leaves row 1 (headers) and the <drawing> reference
    untouched."""
    rows: list[str] = []

    # --- raw data rows, one per sample, starting at row 2 ---
    for i, sample in enumerate(sample_items):
        row_num = i + 2
        cells = []
        for col_letter, value in sorted(
            sample.raw_data.values.items(),
            key=lambda kv: ox.col_letter_to_index(kv[0]),
        ):
            style = STYLE_DATE if hasattr(value, "year") else (
                STYLE_TEXT if isinstance(value, str) else None
            )
            cells.append(ox.cell_xml(f"{col_letter}{row_num}", value,
                                      style=style, sst=sst))
        rows.append(ox.row_xml(row_num, cells))

    next_row = len(sample_items) + 2

    # --- IA Calculation block(s) ---
    for sample in sample_items:
        if sample.ia_calculation is None:
            continue
        divider_row = next_row
        rows.append(
            f'<row r="{divider_row}" spans="2:6" ht="15.75" thickBot="1" '
            f'x14ac:dyDescent="0.3"><c r="B{divider_row}" s="{STYLE_IA_DIVIDER}"/></row>'
        )
        header_row = divider_row + 1
        rows.append(ox.row_xml(
            header_row,
            [ox.cell_xml(f"B{header_row}", "IA Calculation",
                         style=STYLE_IA_HEADER_LABEL, sst=sst),
             f'<c r="C{header_row}" s="{STYLE_IA_HEADER_BLANK}"/>'],
            spans="2:6",
        ))
        line_row = header_row + 1
        for line in sample.ia_calculation.lines:
            rows.append(ox.row_xml(
                line_row,
                [ox.cell_xml(f"B{line_row}", line.label,
                             style=STYLE_IA_LINE_LABEL, sst=sst),
                 ox.cell_xml(f"C{line_row}", line.amount,
                             style=STYLE_IA_LINE_AMOUNT)],
                spans="2:6",
            ))
            line_row += 1
        total_row = line_row
        first_line_row = header_row + 1
        last_line_row = line_row - 1
        formula = f"SUM(C{first_line_row}:C{last_line_row})"
        rows.append(
            f'<row r="{total_row}" spans="2:6" ht="15.75" thickBot="1" '
            f'x14ac:dyDescent="0.3">'
            f'{ox.cell_xml(f"B{total_row}", sample.ia_calculation.total_label, style=STYLE_IA_TOTAL_LABEL, sst=sst)}'
            f'{ox.cell_xml(f"C{total_row}", sample.ia_calculation.total, style=STYLE_IA_TOTAL_AMOUNT, formula=formula)}'
            f'</row>'
        )
        next_row = total_row + 2  # leave one blank row of breathing room

    new_dimension = f'<dimension ref="A1:T{max(next_row, 25)}"/>'
    ox.replace_sheet_rows(
        sheet_xml_path,
        start_marker='<row r="2"', end_marker="</sheetData>",
        new_rows_xml="".join(rows),
        new_dimension=new_dimension, old_dimension=None,
    )
    # dimension replacement needs the *actual* old string; do it separately
    # since we don't know it without reading first:
    data = open(sheet_xml_path, encoding="utf-8").read()
    import re
    data = re.sub(r'<dimension ref="[^"]*"/>', new_dimension, data, count=1)
    open(sheet_xml_path, "w", encoding="utf-8").write(data)


def build_population_sheet_xml(sheet_xml_path: str, population_rows: list[PopulationRow],
                                sst: ox.SharedStrings, *, first_data_row: int = 6) -> None:
    """Replace the Population tab's data rows.

    An EMPTY list leaves the sheet untouched. It used to clear it, which
    meant a caller that simply had no population to hand -- the common case
    while the rest of the pipeline was being wired up -- silently shipped a
    workpaper whose Population tab had been emptied out. Deleting the
    population from an audit workpaper is worse than any formatting defect
    on this sheet: it removes the basis for the sample. "I have nothing to
    write here" must never mean "erase what is there".
    """
    if not population_rows:
        return

    rows = []
    for i, prow in enumerate(population_rows):
        row_num = first_data_row + i
        cells = []
        for col_letter, value in sorted(
            prow.values.items(), key=lambda kv: ox.col_letter_to_index(kv[0])
        ):
            if value is None:
                continue  # template omits blank cells entirely in Population
            style = STYLE_DATE if hasattr(value, "year") else (
                STYLE_TEXT if isinstance(value, str) else
                (STYLE_NUMBER if col_letter in ("K", "L", "M") else None)
            )
            cells.append(ox.cell_xml(f"{col_letter}{row_num}", value,
                                      style=style, sst=sst))
        rows.append(ox.row_xml(row_num, cells))

    last_row = first_data_row + len(population_rows) - 1
    ox.replace_sheet_rows(
        sheet_xml_path,
        start_marker=f'<row r="{first_data_row}"', end_marker="</sheetData>",
        new_rows_xml="".join(rows),
    )
    import re
    data = open(sheet_xml_path, encoding="utf-8").read()
    data = re.sub(r'<dimension ref="[^"]*"/>',
                   f'<dimension ref="A5:T{last_row}"/>', data, count=1)
    open(sheet_xml_path, "w", encoding="utf-8").write(data)


def _write_drawing_and_media(work_dir: str, sheet_num: int, result: db.DrawingResult) -> None:
    drawing_path = os.path.join(work_dir, "xl", "drawings", f"drawing{sheet_num}.xml")
    open(drawing_path, "w", encoding="utf-8").write(result.xml)

    rels_dir = os.path.join(work_dir, "xl", "drawings", "_rels")
    os.makedirs(rels_dir, exist_ok=True)
    rels_path = os.path.join(rels_dir, f"drawing{sheet_num}.xml.rels")
    open(rels_path, "w", encoding="utf-8").write(result.rels_xml())

    media_dir = os.path.join(work_dir, "xl", "media")
    os.makedirs(media_dir, exist_ok=True)
    for _rid, fn, im in result.media:
        im.save(os.path.join(media_dir, fn))


def build_workpaper(request: WorkpaperRequest,
                     *, sample_sheet: str = "Sample", population_sheet: str = "Population",
                     parameters_sheet: str = "Parameters",
                     on_unplaced=None, summary_rows=None) -> str:
    """Writes the workpaper and returns its path.

    on_unplaced, if given, is called once with the list of callouts whose
    anchor could not be located on its evidence image, as
    (image name, letter, anchor_text, reason) tuples.

    Passing it is close to mandatory for any caller that cares whether the
    file is right. An unlocatable callout is DROPPED rather than failing the
    build, and its legend entry is still written -- so without this signal
    the workpaper ships claiming a tickmark it does not have, and the
    reviewer goes hunting for a red box that was never drawn. That is worse
    than either extreme the drop was chosen between.
    """
    work_dir = request.output_path + ".workdir"
    ox.extract_template(request.template_path, work_dir)

    # Refuse a template this builder cannot write correctly, before writing
    # anything. Across a population of controls a wrong-but-produced
    # workpaper costs far more than a refusal -- see check_template.
    sheets = ox.check_template(
        work_dir,
        required_sheets=[sample_sheet, population_sheet],
        max_style_index=STYLE_IA_TOTAL_AMOUNT,
    )
    sample_sheet_num = sheets[sample_sheet]
    population_sheet_num = sheets[population_sheet]
    parameters_sheet_num = sheets.get(parameters_sheet)

    sst = ox.SharedStrings(work_dir)

    sample_sheet_path = os.path.join(work_dir, "xl", "worksheets", f"sheet{sample_sheet_num}.xml")
    pop_sheet_path = os.path.join(work_dir, "xl", "worksheets", f"sheet{population_sheet_num}.xml")

    build_sample_sheet_xml(sample_sheet_path, request.sample_items, sst)
    build_population_sheet_xml(pop_sheet_path, request.population_rows, sst)
    sst.write()

    column_layout = ox.parse_column_widths(sample_sheet_path)
    sample_drawing = db.build_sample_drawing(request.sample_items, column_layout)
    _write_drawing_and_media(work_dir, sample_sheet_num, sample_drawing)

    if request.parameters is not None:
        if parameters_sheet_num is None:
            raise ox.TemplateMismatch(
                f"parameters were supplied but the template has no {parameters_sheet!r} tab"
            )
        params_drawing = db.build_parameters_drawing(request.parameters)
        _write_drawing_and_media(work_dir, parameters_sheet_num, params_drawing)

    if summary_rows:
        # First tab: a reviewer opening the file cold needs the answer
        # before the evidence. Added after the other sheets are written so
        # it cannot disturb their sheetN numbering.
        from agent.ooxml.summary_sheet import add_summary_sheet

        add_summary_sheet(work_dir, summary_rows, sst)
        sst.write()

    if on_unplaced is not None:
        on_unplaced(sample_drawing.unplaced_callouts)

    ox.remove_calc_chain(work_dir)
    ox.repackage(work_dir, request.output_path)
    return request.output_path


# ---------------------------------------------------------------------------
# Worked example: regenerates the Premiere Onboard sample from the previous
# conversation. Run `python build_workpaper.py` to execute it end to end.
# ---------------------------------------------------------------------------

def _premiere_onboard_example() -> WorkpaperRequest:
    import datetime as dt
    from models import (
        RawDataRow, NarrativeParagraph, NarrativeRun, RawDataTieOut,
        EvidenceImage, EvidenceTieOut, IACalculation, IACalculationLine,
        ParametersTab, ParametersCallout,
    )

    raw_data = RawDataRow(values={
        "A": "PK", "B": 881477, "C": 21022452, "D": "Premiere Onboard LLC",
        "E": dt.date(2025, 11, 10), "F": None, "G": 30000, "H": "PV",
        "I": 11743472, "J": "001", "K": 30000, "L": 0, "M": 0, "N": None,
        "O": "        0510", "P": "Talent Acquisition", "Q": None,
        "R": dt.date(2025, 9, 22), "S": "2859", "T": dt.date(2025, 10, 7),
    })

    narrative = [
        NarrativeParagraph(runs=[
            NarrativeRun(text="Test Step 1: ", bold=True, color="RED"),
            NarrativeRun(text=(
                "IA verified that the invoice was properly approved prior to "
                "payment. Diane Milosevic, Senior Manager, Talent Acquisition, "
                "approved the Premiere Onboard invoice (#2859) via email on "
                "10/01/25, which is before the date of payment on 11/10/25."
            )),
        ]),
        NarrativeParagraph(runs=[]),
        NarrativeParagraph(runs=[NarrativeRun(text=(
            "In order to ensure the invoice and approval provided matched the "
            "sample selected and the payment screenshot, IA matched the "
            "following attributes:"
        ))]),
        NarrativeParagraph(runs=[NarrativeRun(text="A - Supplier/Payee Name/Number")]),
        NarrativeParagraph(runs=[NarrativeRun(text="B - Invoice/Payment Amount")]),
        NarrativeParagraph(runs=[NarrativeRun(text="C - Payment Date")]),
        NarrativeParagraph(runs=[NarrativeRun(text="D - Approver")]),
        NarrativeParagraph(runs=[NarrativeRun(text="E - Approval Date")]),
        NarrativeParagraph(runs=[]),
        NarrativeParagraph(runs=[NarrativeRun(
            text="Test step 1 satisfied. No exceptions noted.", bold=True, color="RED")]),
    ]

    sample = SampleItem(
        control_id="SS.PTP.AP.156",
        raw_data=raw_data,
        narrative=narrative,
        raw_data_tie_outs=[
            RawDataTieOut(letter="A", columns=["C", "D"]),
            RawDataTieOut(letter="B", columns=["G"]),
            RawDataTieOut(letter="C", columns=["E"]),
        ],
        evidence_images=[
            EvidenceImage(
                name="Payment Screenshot",
                source_path="/mnt/user-data/uploads/Selection_1_Payment.pdf",
                display_width_emu=6_500_000,
                tie_outs=[
                    EvidenceTieOut(letter="A", anchor_text="21022452", occurrence=0),
                    # NOTE: this JDE screenshot's font makes Tesseract misread
                    # the leading "3" as "2" ("30,000.00" -> "20,000.00").
                    # Anchor on the trailing digits, which OCR reads reliably,
                    # rather than the leading digit, which is the most
                    # error-prone character in a low-res screenshot. Always
                    # spot-check OCR output on real evidence before trusting
                    # an anchor string -- see README "OCR is not exact."
                    EvidenceTieOut(letter="B", anchor_text="0,000.00", occurrence=0),
                    EvidenceTieOut(letter="C", anchor_text="11/10/2025"),
                ],
            ),
            EvidenceImage(
                name="Invoice",
                source_path="/mnt/user-data/uploads/SS_156_selection_1_invoice_2859.pdf",
                display_width_emu=7_800_000,
                tie_outs=[
                    EvidenceTieOut(letter="A", anchor_text="21022452"),
                    EvidenceTieOut(letter="B", anchor_text="30,000.00", occurrence=0),
                    # The signature is a cursive/script font ("Diane
                    # Milosevic") which Tesseract reliably mis-OCRs as
                    # gibberish (e.g. "Whlocwwer"). "Approval:" immediately
                    # preceding it is printed text and OCRs cleanly, so
                    # anchor there and pull in the next word (the garbled
                    # signature) rather than trying to match the name itself.
                    EvidenceTieOut(letter="D", anchor_text="Approval:", extra_words=1),
                ],
            ),
            EvidenceImage(
                name="Approval Email",
                source_path="/mnt/user-data/uploads/SS_156_selection_1_-_invoice_2859_approval.pdf",
                display_width_emu=6_100_000,
                tie_outs=[
                    EvidenceTieOut(letter="A", anchor_text="21022452"),
                    # find_anchor_box matches single OCR word-tokens, so a
                    # two-word phrase must be built from anchor + extra_words
                    # rather than searched for as one string. "Diane" appears
                    # 5x on this page in reading order: (0) the From: header,
                    # (1) "Thank you, / Diane" sign-off (no surname on that
                    # line), (2) the "Diane Milosevic" signature block we
                    # actually want, (3)-(4) inside the quoted/nested emails
                    # further down. occurrence=2 + extra_words=1 boxes
                    # "Diane Milosevic" together.
                    EvidenceTieOut(letter="D", anchor_text="Diane", occurrence=2, extra_words=1),
                    EvidenceTieOut(letter="E", anchor_text="October"),
                ],
            ),
        ],
        ia_calculation=IACalculation(lines=[
            IACalculationLine(label="Kristen Attardo - BD - Phoenix, AZ - Start Date 9/22/25", amount=10000),
            IACalculationLine(label="Matt Magaraci - BD - Hamilton, NJ - Start Date 9/22/25", amount=10000),
            IACalculationLine(label="Juan Lapa - BD - Scottsdale, AZ - Start Date 9/22/25", amount=10000),
        ]),
        conclusion_satisfied=True,
    )

    # population_rows omitted here for brevity -- in the real run these come
    # from SS_156_Sample_listing_upload.xlsx's Population tab, loaded with
    # openpyxl (read-only) and converted to PopulationRow objects.
    population_rows: list[PopulationRow] = []

    params = ParametersTab(
        screenshot_path="/mnt/user-data/uploads/params_screenshot.png",
        screenshot_dpi=144.0,
        display_width_emu=None,
        narrative=[
            NarrativeParagraph(runs=[NarrativeRun(text=(
                "For completeness and accuracy procedures, IA verified the "
                "population was run for the correct period 10/1/25-3/31/26. "
                "IA also validated that the document types excluded are "
                "related to PO invoices. IA noticed that the population "
                "provided included 6 document types related to non-PO "
                "invoices, as follows:"
            ))]),
            NarrativeParagraph(runs=[NarrativeRun(
                text="PV is Payment Voucher. This is the main document type for non-PO invoices.")]),
            NarrativeParagraph(runs=[NarrativeRun(
                text="PC is related to Co-Start lease payments. Moving forward, this should be excluded from the population.")]),
            NarrativeParagraph(runs=[NarrativeRun(
                text="PS, PU, T7, and PR have been determined to not be material and should also be excluded from the population.")]),
        ],
        callouts=[
            # Both of these small/dense UI screenshots get misread by OCR at
            # the digit level ("57,039" -> "57059", the first "10/01/25" on
            # the page -> "10/01/28") even though the actual pixels are
            # correct -- Tesseract just struggles with this UI's font at this
            # resolution. Anchor on the surrounding label text (which OCRs
            # reliably) instead of the exact digits, and always eyeball the
            # rendered callout against the source screenshot once. See
            # README "OCR is not exact."
            ParametersCallout(anchor_text="Selected", extra_words=2),
            # Only the *second* on-screen "Between 10/01/25 and 3/31/26" (the
            # one under the column headers) OCRs correctly; the first, inside
            # the small parameter input box, gets misread as "10/01/28". So
            # there's exactly one clean match here, occurrence=0.
            ParametersCallout(anchor_text="10/01/25", occurrence=0),
        ],
    )

    return WorkpaperRequest(
        template_path="/mnt/user-data/uploads/T0_SS_PTP_AP_156_01__PY_Testing_wp.xlsx",
        output_path="/mnt/user-data/outputs/T0_SS_PTP_AP_156_01__PY_Testing_wp.xlsx",
        sample_items=[sample],
        population_rows=population_rows,
        parameters=params,
    )


if __name__ == "__main__":
    req = _premiere_onboard_example()
    out = build_workpaper(req)
    print("wrote", out)
