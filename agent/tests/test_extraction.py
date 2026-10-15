"""Extraction tests against synthetic files -- the "prove it against a real
file" check for the biggest unknown in the design (garbage extraction
poisons everything downstream).
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.pdfgen import canvas

from agent.extraction import extract_excel, extract_many, extract_pdf


# --------------------------------------------------------------------------
# Fixtures: build synthetic files rather than checking binaries into the repo
# --------------------------------------------------------------------------


@pytest.fixture
def sample_xlsx(tmp_path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PO Testing"

    header = ["Sample", "PO Number", "Branch", "Amount"]
    rows = [
        ["S01", "PO-48213", "35210", 12450],
        ["S02", "PO-48214", "35211", 8300],
        ["S03", "PO-48215", "35210", 15020],
    ]
    for r, values in enumerate([header, *rows], start=1):
        for c, value in enumerate(values, start=1):
            ws.cell(row=r, column=c, value=value)

    # Exception: preparer highlighted a cell and left a comment on it.
    flagged_cell = ws.cell(row=3, column=4)  # S02 amount
    flagged_cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    flagged_cell.comment = Comment("Approval missing supervisor signature", "Preparer")

    path = tmp_path / "po_sample.xlsx"
    wb.save(path)
    return path


@pytest.fixture
def native_text_pdf(tmp_path: Path) -> Path:
    path = tmp_path / "gl_export.pdf"
    doc = SimpleDocTemplate(str(path), pagesize=letter)
    data = [
        ["Account", "Description", "Amount"],
        ["6100", "Accrued Interest", "482110"],
        ["6200", "AP Trade", "119400"],
    ]
    table = Table(data)
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ]
        )
    )
    doc.build([table])
    return path


@pytest.fixture
def scanned_pdf(tmp_path: Path) -> Path:
    # A page with only a rasterized image and no text layer -- the case
    # extract_pdf should flag as needing OCR rather than silently dropping.
    img_path = tmp_path / "screenshot.png"
    Image.new("RGB", (400, 200), color="white").save(img_path)

    path = tmp_path / "screenshot_as_pdf.pdf"
    c = canvas.Canvas(str(path), pagesize=letter)
    c.drawImage(str(img_path), 50, 500, width=400, height=200)
    c.showPage()
    c.save()
    return path


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------


def test_excel_table_extracted_with_header_and_rows(sample_xlsx: Path):
    items = extract_excel(sample_xlsx)
    tables = [i for i in items if i.source_type == "excel_table"]

    assert len(tables) == 1
    table = tables[0].extracted_table
    assert table[0] == ["Sample", "PO Number", "Branch", "Amount"]
    assert table[1] == ["S01", "PO-48213", "35210", "12450"]
    assert tables[0].location.startswith("PO Testing!A1:")
    assert tables[0].extraction_confidence == 1.0


def test_excel_flagged_cell_captured_as_evidence(sample_xlsx: Path):
    items = extract_excel(sample_xlsx)
    cells = [i for i in items if i.source_type == "excel_cell"]

    assert len(cells) == 1
    cell = cells[0]
    assert cell.location == "PO Testing!D3"
    assert "FFFF00" in cell.extracted_text
    assert "Approval missing supervisor signature" in cell.extracted_text


def test_excel_whole_row_highlight_collapses_to_one_item(tmp_path: Path):
    # Real preparer convention seen in a live run: the ENTIRE row of a
    # sampled population item gets highlighted (every column), not one
    # cell. The old per-cell logic turned that into one EvidenceItem per
    # cell -- individually meaningless ("value: 'PK'; fill_color: ...")
    # and already redundant with the row's own content in the table
    # extraction. This is the direct regression test for that: it flooded
    # a real run's evidence pool (57 fragments from 3 rows) and burned the
    # whole tool-call budget with the model trying to chase them down.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Population"
    header = ["Invoice", "Vendor", "Amount", "Date"]
    ws.append(header)
    ws.append(["INV-1", "Acme", 100, "2026-01-01"])
    ws.append(["INV-2", "Premiere Onboard LLC", 30000, "2026-02-01"])  # the sampled row
    ws.append(["INV-3", "Beta", 200, "2026-03-01"])

    yellow = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
    for col in range(1, 5):
        ws.cell(row=3, column=col).fill = yellow

    path = tmp_path / "population.xlsx"
    wb.save(path)

    items = extract_excel(path)
    cells = [i for i in items if i.source_type == "excel_cell"]

    assert len(cells) == 1
    assert cells[0].location == "Population!A3:D3"
    assert "entire row highlighted" in cells[0].extracted_text
    assert "4 of 4 populated columns" in cells[0].extracted_text
    # Untouched rows produce no cell-level noise at all.
    assert not any("INV-1" in c.extracted_text or "INV-3" in c.extracted_text for c in cells)


def test_excel_sparse_highlight_in_row_stays_per_cell(tmp_path: Path):
    # A single flagged cell in an otherwise-normal row (the real
    # PO-testing exception case) must NOT be swept into a row-level
    # summary -- only a genuinely whole-row highlight collapses.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Invoice", "Vendor", "Amount", "Date"])
    ws.append(["INV-1", "Acme", 100, "2026-01-01"])
    ws.cell(row=2, column=3).fill = PatternFill(
        start_color="FFFF0000", end_color="FFFF0000", fill_type="solid"
    )

    path = tmp_path / "sparse.xlsx"
    wb.save(path)

    items = extract_excel(path)
    cells = [i for i in items if i.source_type == "excel_cell"]
    assert len(cells) == 1
    assert cells[0].location == "Sheet!C2"
    assert "entire row highlighted" not in cells[0].extracted_text
    assert "value: 100" in cells[0].extracted_text


def test_excel_comment_in_highlighted_row_still_reported_individually(tmp_path: Path):
    # A comment is real, specific content a preparer wrote -- it must
    # survive even when the rest of its row collapses into a row summary.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Invoice", "Vendor", "Amount", "Date"])
    ws.append(["INV-1", "Acme", 100, "2026-01-01"])

    yellow = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
    for col in range(1, 5):
        ws.cell(row=2, column=col).fill = yellow
    ws.cell(row=2, column=3).comment = Comment("Confirmed against remittance", "Preparer")

    path = tmp_path / "commented_row.xlsx"
    wb.save(path)

    items = extract_excel(path)
    cells = [i for i in items if i.source_type == "excel_cell"]

    row_summaries = [c for c in cells if "entire row highlighted" in c.extracted_text]
    commented = [c for c in cells if "comment:" in c.extracted_text]
    assert len(row_summaries) == 1
    assert len(commented) == 1
    assert commented[0].location == "Sheet!C2"
    assert "Confirmed against remittance" in commented[0].extracted_text


def test_excel_all_ids_unique(sample_xlsx: Path):
    items = extract_excel(sample_xlsx)
    ids = [i.evidence_id for i in items]
    assert len(ids) == len(set(ids))


def test_excel_large_table_is_chunked_not_one_giant_blob(tmp_path: Path):
    # A real population export can be thousands of unbroken rows -- this is
    # what would otherwise become a single EvidenceItem that then gets
    # rendered in full into every prompt turn with no size cap of its own.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Invoice", "Amount"])
    for i in range(450):
        ws.append([f"INV-{i}", i])
    path = tmp_path / "big_population.xlsx"
    wb.save(path)

    items = extract_excel(path)
    tables = [i for i in items if i.source_type == "excel_table"]

    assert len(tables) == 3  # 450 data rows / 200 per chunk -> 3 chunks
    for t in tables:
        assert len(t.extracted_table) <= 201  # header + up to 200 data rows
        assert t.extracted_table[0] == ["Invoice", "Amount"]  # header repeated in every chunk

    # No row lost or duplicated across chunks.
    all_invoices = [row[0] for t in tables for row in t.extracted_table[1:]]
    assert all_invoices == [f"INV-{i}" for i in range(450)]


# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------


def test_native_pdf_text_and_table_extracted(native_text_pdf: Path):
    items = extract_pdf(native_text_pdf)

    text_items = [i for i in items if i.source_type == "pdf_text"]
    table_items = [i for i in items if i.source_type == "pdf_table"]

    assert len(text_items) == 1
    assert "Accrued Interest" in text_items[0].extracted_text
    assert text_items[0].extraction_confidence == 1.0

    assert len(table_items) == 1
    table = table_items[0].extracted_table
    assert table[0] == ["Account", "Description", "Amount"]
    assert ["6100", "Accrued Interest", "482110"] in table


def test_scanned_pdf_flagged_not_dropped(scanned_pdf: Path):
    items = extract_pdf(scanned_pdf)

    assert len(items) == 1
    item = items[0]
    assert item.source_type == "image_ocr"
    assert item.extraction_confidence == 0.0
    assert item.extracted_text is None
    # The page still shows up with a location -- nothing silently vanished.
    assert "p.1" in item.location


# --------------------------------------------------------------------------
# Multi-file merge
# --------------------------------------------------------------------------


def test_extract_many_renumbers_ids_without_collision(sample_xlsx: Path, native_text_pdf: Path):
    # Each single-file extractor numbers ev_0001.. independently -- naively
    # concatenating extract(a) + extract(b) would collide. extract_many must not.
    solo_excel = extract_excel(sample_xlsx)
    solo_pdf = extract_pdf(native_text_pdf)
    assert solo_excel[0].evidence_id == solo_pdf[0].evidence_id == "ev_0001"  # confirms the collision exists

    merged = extract_many([sample_xlsx, native_text_pdf])

    assert len(merged) == len(solo_excel) + len(solo_pdf)
    ids = [item.evidence_id for item in merged]
    assert len(ids) == len(set(ids))

    # Original file/location/content is preserved -- only evidence_id changed.
    assert merged[0].source_file == "po_sample.xlsx"
    assert merged[len(solo_excel)].source_file == "gl_export.pdf"


# --------------------------------------------------------------------------
# Fillable (form) PDFs
# --------------------------------------------------------------------------


def _form_pdf(path: Path, *, with_heading: bool = True) -> Path:
    c = canvas.Canvas(str(path))
    if with_heading:
        c.drawString(50, 800, "Check Request")
    c.drawString(50, 770, "Vendor:")
    c.acroForm.textfield(name="Text1", x=120, y=760, width=200, height=20, value="Acme Supply Co")
    c.drawString(50, 730, "Amount:")
    c.acroForm.textfield(name="Text2", x=120, y=720, width=200, height=20, value="14,945.60")
    c.drawString(50, 690, "Approved:")
    c.acroForm.checkbox(name="Check1", x=120, y=685, checked=True)
    c.acroForm.textfield(name="Blank", x=120, y=650, width=200, height=20, value="")
    c.save()
    return path


def test_filled_in_form_values_are_read_with_their_printed_labels(tmp_path: Path):
    # A fillable form keeps its answers in fields, not the page text: the
    # agent used to see the labels and none of the values.
    text = extract_pdf(_form_pdf(tmp_path / "form.pdf"))[0].extracted_text
    assert "Vendor: Acme Supply Co" in text
    assert "Amount: 14,945.60" in text
    assert "Approved: checked" in text
    assert "Blank" not in text  # empty fields add nothing


def test_a_form_page_with_little_printed_text_is_not_mistaken_for_a_scan(tmp_path: Path):
    items = extract_pdf(_form_pdf(tmp_path / "form.pdf", with_heading=False))
    assert items[0].source_type == "pdf_text"
    assert "Acme Supply Co" in items[0].extracted_text


def test_form_pages_render_with_their_values_showing(tmp_path: Path):
    # The exhibit image is rendered with forms on; without it every field
    # came out blank in the workpaper.
    import pdfplumber
    from PIL import ImageChops

    from agent.ooxml.imaging import render_pdf_page

    path = _form_pdf(tmp_path / "form.pdf")
    ours = render_pdf_page(str(path), 0, dpi=72)
    with pdfplumber.open(path) as pdf:
        without_forms = pdf.pages[0].to_image(resolution=72).original.convert("RGB")
    field_area = (120, 792 - 780, 320, 792 - 720)  # where the vendor/amount fields sit
    assert ImageChops.difference(ours.crop(field_area), without_forms.crop(field_area)).getbbox()


def _xfa_pdf(path: Path, datasets_xml: bytes) -> Path:
    """A minimal dynamic (XFA) form: one blank page, answers in the data packet."""
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R /AcroForm 4 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>",
        b"<< /Fields [] /XFA [(datasets) 5 0 R] >>",
        b"<< /Length %d >>\nstream\n" % len(datasets_xml) + datasets_xml + b"\nendstream",
    ]
    out = bytearray(b"%PDF-1.7\n")
    offsets = []
    for i, body in enumerate(objects, start=1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % i + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    out += b"".join(b"%010d 00000 n \n" % o for o in offsets)
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objects) + 1, xref)
    path.write_bytes(bytes(out))
    return path


def test_dynamic_xfa_form_answers_are_read_from_the_form_data(tmp_path: Path):
    xml = (b'<xfa:datasets xmlns:xfa="http://www.xfa.org/schema/xfa-data/1.0/"><xfa:data>'
           b"<form1><VendorName>Acme Supply Co</VendorName><Amount>14945.60</Amount>"
           b"<ApprovedBy>M. Chen</ApprovedBy><Empty/></form1></xfa:data></xfa:datasets>")
    items = extract_pdf(_xfa_pdf(tmp_path / "xfa.pdf", xml))
    data = next(i for i in items if "(form data)" in i.location)
    assert "VendorName: Acme Supply Co" in data.extracted_text
    assert "ApprovedBy: M. Chen" in data.extracted_text
    assert "Empty" not in data.extracted_text
