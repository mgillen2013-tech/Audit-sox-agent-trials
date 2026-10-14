"""
Top-level entry point. build_workpaper(request) turns a validated
WorkpaperRequest into a finished .xlsx, doing all the OOXML/EMU/OCR work
described in README.md.

CELL FORMATS ARE COPIED FROM THE TEMPLATE, NOT ASSUMED
------------------------------------------------------
A cell's format is a number -- a position in that workbook's styles.xml.
The first version hard-coded those numbers (STYLE_* below) from one PY
file. That is wrong in two ways at once:

  - every PY workpaper has its own styles.xml, so across 172 controls
    "format #2" means something different in each, and a date can come out
    as 46061 or a percentage;
  - Excel COMPACTS styles.xml every time a file is saved, so even the one
    template they were lifted from stops matching once someone opens and
    saves it. A real run hit exactly that and fell back to the standalone
    layout.

So each cell takes the format of the template cell it replaces (the PY
sample row, the PY population row, the PY IA Calculation box). Where the
template has no such cell, a format is appended to styles.xml for the
purpose -- never looked up by position. See _resolve_styles. The STYLE_*
constants remain only as defaults for direct callers that pass no styles.
"""
from __future__ import annotations
import os
from dataclasses import dataclass, field
from PIL import Image

from agent.ooxml.models import WorkpaperRequest, SampleItem, PopulationRow, RawDataRow
from agent.ooxml import ooxml_utils as ox
from agent.ooxml import drawing_builder as db
from agent.ooxml.summary_sheet import DRAFT_BANNER

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


@dataclass
class CellStyles:
    """The format index for each kind of cell this builder writes."""

    text: int = STYLE_TEXT
    date: int = STYLE_DATE
    number: int = STYLE_NUMBER
    ia_divider: int = STYLE_IA_DIVIDER
    ia_header_label: int = STYLE_IA_HEADER_LABEL
    ia_header_blank: int = STYLE_IA_HEADER_BLANK
    ia_line_label: int = STYLE_IA_LINE_LABEL
    ia_line_amount: int = STYLE_IA_LINE_AMOUNT
    ia_total_label: int = STYLE_IA_TOTAL_LABEL
    ia_total_amount: int = STYLE_IA_TOTAL_AMOUNT
    # Column letter -> the format the TEMPLATE used in that column.
    sample_columns: dict[str, int] = field(default_factory=dict)
    population_columns: dict[str, int] = field(default_factory=dict)

    def for_value(self, value, column: str, by_column: dict[str, int]) -> int | None:
        if column in by_column:
            return by_column[column]
        if hasattr(value, "year"):
            return self.date
        if isinstance(value, str):
            return self.text
        if isinstance(value, float):
            return self.number
        return None


def _append_styles(work_dir: str) -> CellStyles:
    """Append a complete set of formats to styles.xml; return their indices.

    The fallback for anything the template cannot show us. Appended, never
    inserted: every existing cell addresses formats by position, so
    renumbering would restyle the whole workbook.
    """
    import re as _re

    path = os.path.join(work_dir, "xl", "styles.xml")
    with open(path, encoding="utf-8") as fh:
        xml = fh.read()

    def count(tag: str) -> int:
        m = _re.search(rf"<{tag}\b[^>]*count=\"(\d+)\"", xml)
        return int(m.group(1)) if m else len(_re.findall(rf"<{tag[:-1]}\b", xml))

    n_fonts, n_borders, n_xfs = count("fonts"), count("borders"), count("cellXfs")
    bold = n_fonts
    thin_box, rule_below, total_rule = n_borders, n_borders + 1, n_borders + 2
    fonts = '<font><b/><sz val="11"/><name val="Calibri"/></font>'
    borders = (
        '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>'
        '<border><left/><right/><top/><bottom style="medium"/><diagonal/></border>'
        '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="double"/><diagonal/></border>'
    )

    def xf(fmt: int, font: int = 0, border: int = 0) -> str:
        return (
            f'<xf numFmtId="{fmt}" fontId="{font}" fillId="0" borderId="{border}" '
            f'applyNumberFormat="1" applyFont="1" applyBorder="1"/>'
        )

    # Built-in number formats: 0 General, 14 short date, 4 #,##0.00.
    new = [
        xf(0), xf(14), xf(4),                 # text, date, number
        xf(0, border=rule_below),             # ia_divider
        xf(0, bold, thin_box), xf(0, 0, thin_box),       # ia header label / blank
        xf(0, 0, thin_box), xf(4, 0, thin_box),          # ia line label / amount
        xf(0, bold, total_rule), xf(4, bold, total_rule),  # ia total label / amount
    ]
    for tag, addition, added in (
        ("fonts", fonts, 1), ("borders", borders, 3), ("cellXfs", "".join(new), len(new)),
    ):
        xml = _re.sub(rf"(<{tag}\b[^>]*count=\")(\d+)", lambda m: f"{m.group(1)}{int(m.group(2)) + added}", xml, count=1)
        xml = xml.replace(f"</{tag}>", f"{addition}</{tag}>", 1)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(xml)

    i = iter(range(n_xfs, n_xfs + len(new)))
    return CellStyles(*(next(i) for _ in new))


def _template_cells(sheet_xml_path: str, shared: list) -> dict[str, tuple[int | None, str, bool]]:
    """{ref: (format index, text, has formula)} for every cell on a tab."""
    import xml.etree.ElementTree as ET

    ns = f"{{{ox.NS_MAIN}}}"
    out: dict[str, tuple[int | None, str, bool]] = {}
    for c in ET.parse(sheet_xml_path).getroot().iter(f"{ns}c"):
        v = c.find(f"{ns}v")
        raw = v.text if v is not None and v.text else ""
        if c.get("t") == "s" and raw.isdigit() and int(raw) < len(shared):
            raw = shared[int(raw)] or ""
        elif c.get("t") == "inlineStr":
            raw = "".join(t.text or "" for t in c.iter(f"{ns}t"))
        style = c.get("s")
        out[c.get("r", "")] = (int(style) if style else None, raw, c.find(f"{ns}f") is not None)
    return out


def _split_ref(ref: str) -> tuple[str, int]:
    import re as _re

    m = _re.fullmatch(r"([A-Z]+)(\d+)", ref)
    return (m.group(1), int(m.group(2))) if m else ("", 0)


def _resolve_styles(
    work_dir: str, sample_sheet_path: str, population_sheet_path: str,
    shared: list, *, population_first_row: int = 6,
) -> CellStyles:
    """Formats for every cell this builder writes, copied from the template
    wherever the template has the equivalent cell. Must run BEFORE any rows
    are replaced -- it reads the PY's own cells.
    """
    styles = _append_styles(work_dir)

    sample = _template_cells(sample_sheet_path, shared)
    for ref, (style, _text, _f) in sample.items():
        col, row = _split_ref(ref)
        if row == 2 and style is not None:
            styles.sample_columns[col] = style

    population = _template_cells(population_sheet_path, shared)
    for ref, (style, _text, _f) in population.items():
        col, row = _split_ref(ref)
        if row == population_first_row and style is not None:
            styles.population_columns[col] = style

    # The PY's IA Calculation box, if it has one: found by its label, then
    # read cell by cell around it.
    label = next(
        (ref for ref, (_s, text, _f) in sample.items() if text.strip().lower() == "ia calculation"),
        None,
    )
    if label:
        col, row = _split_ref(label)
        nxt = ox.col_index_to_letter(ox.col_letter_to_index(col) + 1)

        def style_at(ref: str) -> int | None:
            return sample.get(ref, (None, "", False))[0]

        for attr, ref in (
            ("ia_header_label", label), ("ia_header_blank", f"{nxt}{row}"),
            ("ia_divider", f"{col}{row - 1}"),
            ("ia_line_label", f"{col}{row + 1}"), ("ia_line_amount", f"{nxt}{row + 1}"),
        ):
            got = style_at(ref)
            if got is not None:
                setattr(styles, attr, got)
        total = next(
            (r for r in range(row + 1, row + 30) if sample.get(f"{nxt}{r}", (None, "", False))[2]),
            None,
        )
        if total:
            for attr, ref in (("ia_total_label", f"{col}{total}"), ("ia_total_amount", f"{nxt}{total}")):
                got = style_at(ref)
                if got is not None:
                    setattr(styles, attr, got)
    return styles


def sample_tab_name(base: str, sample_id: str, index: int, total: int) -> str:
    """Tab name for one sampled item.

    A single sample keeps the template's own tab name, so the ordinary case
    is byte-for-byte the shape of last year's workbook. Several samples get
    numbered tabs. Excel caps a sheet name at 31 characters and forbids
    : \\ / ? * [ ], and a name that violates either makes the whole
    workbook unopenable -- so the identifier is sanitised rather than
    trusted, since it comes from a sample listing nobody validated.
    """
    if total <= 1:
        return base
    # From the SAMPLE ID, never from a data column. The Summary tab builds
    # its hyperlink from the same id, and a first version named tabs after
    # the check number while the links used the id -- so every link pointed
    # at a tab that did not exist. Excel does not complain about that; the
    # click just does nothing.
    raw = f"{base} {sample_id or index + 1}"
    for bad in ':\\/?*[]':
        raw = raw.replace(bad, "-")
    return raw[:31].strip() or f"{base} {index + 1}"


def adopt_prior_output(work_dir: str, sample_sheet: str = "Sample") -> list[str]:
    """Make a previous CY workpaper usable as this run's template.

    A workpaper this builder wrote is the natural template for the next
    cycle -- it IS last year's file -- but it no longer has the shape the
    builder looks for. Its single "Sample" tab became "Sample 1", "Sample
    2", ... (one per sampled item) and a Summary was added in front. A real
    run uploaded one as the prior year, the "Sample" tab was missing, and
    the whole workpaper silently fell back to the standalone layout: no
    tickmarks, no PY format, no explanation in the file.

    So, when -- and only when -- the named sample tab is absent:
      - the numbered sample tabs are recognised by name ("Sample" + a
        space + something containing a digit, which is what this builder
        produces; "Sample Selections" does not qualify and is never touched)
      - the FIRST becomes the layout every new sample tab is cloned from,
        renamed back to the plain name
      - the rest are DELETED. Leaving them would ship last year's sampled
        items, with last year's exhibits, inside this year's workpaper.

    Independently, a Summary tab this builder generated is deleted so a
    fresh one can replace it. It is recognised by its HEADINGS, not its
    name: a hand-built PY workpaper may have a "Summary" of its own, and
    deleting that would destroy prior-year content.

    Returns one plain-English line per change, for the run's notes. Empty
    when the template was an ordinary PY workpaper and nothing was done.
    """
    import re as _re
    from agent.ooxml.summary_sheet import GENERATED_HEADINGS

    notes: list[str] = []
    sheets = ox.resolve_sheets(work_dir)

    for name, num in list(sheets.items()):
        rows = ox.first_rows_text(work_dir, num, rows=2)
        if any(tuple(r[: len(GENERATED_HEADINGS)]) == GENERATED_HEADINGS for r in rows):
            ox.delete_sheet(work_dir, name)
            notes.append(f"removed the previous run's {name!r} tab; a new one replaces it")

    sheets = ox.resolve_sheets(work_dir)
    if sample_sheet in sheets:
        return notes

    numbered = [
        name for name in sheets
        if _re.fullmatch(rf"{_re.escape(sample_sheet)} .*\d.*", name)
    ]
    if not numbered:
        return notes
    base, stale = numbered[0], numbered[1:]
    for name in stale:
        ox.delete_sheet(work_dir, name)
    ox.rename_sheet(work_dir, ox.resolve_sheets(work_dir)[base], sample_sheet)
    notes.append(
        f"the prior-year file is a previous CY workpaper: used its {base!r} tab as the "
        f"sample layout"
        + (f" and removed last cycle's {', '.join(repr(n) for n in stale)}" if stale else "")
    )
    return notes


def build_sample_sheet_xml(sheet_xml_path: str, sample_items: list[SampleItem],
                            sst: ox.SharedStrings, styles: CellStyles | None = None) -> None:
    """Overwrite the Sample sheet's data rows (raw data rows + IA Calculation
    block(s)) in place. Leaves row 1 (headers) and the <drawing> reference
    untouched."""
    st = styles or CellStyles()
    rows: list[str] = []

    # --- raw data rows, one per sample, starting at row 2 ---
    for i, sample in enumerate(sample_items):
        row_num = i + 2
        cells = []
        for col_letter, value in sorted(
            sample.raw_data.values.items(),
            key=lambda kv: ox.col_letter_to_index(kv[0]),
        ):
            style = (
                st.for_value(value, col_letter, st.sample_columns) if styles
                else STYLE_DATE if hasattr(value, "year")
                else STYLE_TEXT if isinstance(value, str) else None
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
            f'<row r="{divider_row}" spans="2:6" ht="15.75" thickBot="1">'
            f'<c r="B{divider_row}" s="{st.ia_divider}"/></row>'
        )
        header_row = divider_row + 1
        rows.append(ox.row_xml(
            header_row,
            [ox.cell_xml(f"B{header_row}", "IA Calculation",
                         style=st.ia_header_label, sst=sst),
             f'<c r="C{header_row}" s="{st.ia_header_blank}"/>'],
            spans="2:6",
        ))
        line_row = header_row + 1
        for line in sample.ia_calculation.lines:
            rows.append(ox.row_xml(
                line_row,
                [ox.cell_xml(f"B{line_row}", line.label,
                             style=st.ia_line_label, sst=sst),
                 ox.cell_xml(f"C{line_row}", line.amount,
                             style=st.ia_line_amount)],
                spans="2:6",
            ))
            line_row += 1
        total_row = line_row
        first_line_row = header_row + 1
        last_line_row = line_row - 1
        formula = f"SUM(C{first_line_row}:C{last_line_row})"
        rows.append(
            f'<row r="{total_row}" spans="2:6" ht="15.75" thickBot="1">'
            f'{ox.cell_xml(f"B{total_row}", sample.ia_calculation.total_label, style=st.ia_total_label, sst=sst)}'
            f'{ox.cell_xml(f"C{total_row}", sample.ia_calculation.total, style=st.ia_total_amount, formula=formula)}'
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
                                sst: ox.SharedStrings, *, first_data_row: int = 6,
                                styles: CellStyles | None = None) -> None:
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
            style = (
                styles.for_value(value, col_letter, styles.population_columns) if styles
                else STYLE_DATE if hasattr(value, "year")
                else STYLE_TEXT if isinstance(value, str)
                else (STYLE_NUMBER if col_letter in ("K", "L", "M") else None)
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


def _drawing_for_sheet(work_dir: str, sheet_num: int) -> str:
    """Name of the drawing part this sheet uses, creating the link if absent.

    Two assumptions used to live here, both invisible when wrong:

    1. that xl/drawings/ exists. On a template whose sheet carries no
       shapes it does not, and writing into it raised FileNotFoundError
       partway through a build.
    2. that drawingN.xml pairs with sheetN.xml. That holds for the template
       this was built from and is a coincidence -- the pairing lives in the
       sheet's own .rels. Guessing produces an ORPHANED drawing part: the
       file is written, the sheet never references it, and the workpaper
       opens with no exhibits and no tickmarks at all. Silent and total.

    So: use the target the sheet already points at, and when it points at
    nothing, create one and register it (part + sheet rel + <drawing>
    element + content type) the way Excel expects.
    """
    import re as _re

    os.makedirs(os.path.join(work_dir, "xl", "drawings", "_rels"), exist_ok=True)
    sheet_path = os.path.join(work_dir, "xl", "worksheets", f"sheet{sheet_num}.xml")
    rels_path = os.path.join(work_dir, "xl", "worksheets", "_rels", f"sheet{sheet_num}.xml.rels")

    with open(sheet_path, encoding="utf-8") as fh:
        sheet = fh.read()
    existing_rid = _re.search(r'<drawing r:id="([^"]+)"', sheet)
    if existing_rid and os.path.exists(rels_path):
        with open(rels_path, encoding="utf-8") as fh:
            rels = fh.read()
        target = _re.search(
            rf'<Relationship[^>]*Id="{_re.escape(existing_rid.group(1))}"[^>]*Target="([^"]*drawing[^"]+)"',
            rels,
        ) or _re.search(
            rf'<Relationship[^>]*Target="([^"]*drawing[^"]+)"[^>]*Id="{_re.escape(existing_rid.group(1))}"',
            rels,
        )
        if target:
            return os.path.basename(target.group(1))

    # drawing<sheet_num>.xml is only a preference. Excel numbers drawings
    # independently of sheets, so that name can already belong to ANOTHER
    # tab -- and overwriting it would put this tab's exhibits on that one.
    drawings_dir = os.path.join(work_dir, "xl", "drawings")
    k = sheet_num
    while os.path.exists(os.path.join(drawings_dir, f"drawing{k}.xml")):
        k += 1
    name = f"drawing{k}.xml"
    rid = "rIdDrawingCY"
    os.makedirs(os.path.dirname(rels_path), exist_ok=True)
    rels = (
        open(rels_path, encoding="utf-8").read()
        if os.path.exists(rels_path)
        else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             "</Relationships>"
    )
    rels = rels.replace(
        "</Relationships>",
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/drawing" Target="../drawings/{name}"/>'
        "</Relationships>",
        1,
    )
    open(rels_path, "w", encoding="utf-8").write(rels)

    if "xmlns:r=" not in sheet:
        sheet = sheet.replace(
            "<worksheet ",
            '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
            1,
        )
    sheet = sheet.replace("</worksheet>", f'<drawing r:id="{rid}"/></worksheet>', 1)
    open(sheet_path, "w", encoding="utf-8").write(sheet)

    ct_path = os.path.join(work_dir, "[Content_Types].xml")
    ct = open(ct_path, encoding="utf-8").read()
    if f"/xl/drawings/{name}" not in ct:
        ct = ct.replace(
            "</Types>",
            f'<Override PartName="/xl/drawings/{name}" ContentType="application/'
            f'vnd.openxmlformats-officedocument.drawing+xml"/></Types>',
            1,
        )
        open(ct_path, "w", encoding="utf-8").write(ct)
    return name


def _write_drawing_and_media(work_dir: str, sheet_num: int, result: db.DrawingResult) -> None:
    # The drawing builder numbers its images from image1.png for every tab
    # it draws, but xl/media is ONE folder for the whole workbook. Written
    # as-is, sample 2's exhibits overwrote sample 1's -- and the template's
    # own images on other tabs -- so every tab showed the last sample's
    # pages. Renamed here to names nothing else is using.
    media_dir = os.path.join(work_dir, "xl", "media")
    os.makedirs(media_dir, exist_ok=True)
    taken = set(os.listdir(media_dir))
    renamed = []
    n = 1
    for rid, _fn, im in result.media:
        while f"image{n}.png" in taken:
            n += 1
        taken.add(f"image{n}.png")
        renamed.append((rid, f"image{n}.png", im))
    result.media = renamed

    name = _drawing_for_sheet(work_dir, sheet_num)
    drawing_path = os.path.join(work_dir, "xl", "drawings", name)
    open(drawing_path, "w", encoding="utf-8").write(result.xml)

    rels_dir = os.path.join(work_dir, "xl", "drawings", "_rels")
    os.makedirs(rels_dir, exist_ok=True)
    open(os.path.join(rels_dir, f"{name}.rels"), "w", encoding="utf-8").write(result.rels_xml())

    for _rid, fn, im in result.media:
        im.save(os.path.join(media_dir, fn))


def build_workpaper(request: WorkpaperRequest,
                     *, sample_sheet: str = "Sample", population_sheet: str = "Population",
                     parameters_sheet: str = "Parameters",
                     on_unplaced=None, summary_rows=None,
                     draft_banner: str | None = DRAFT_BANNER) -> str:
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

    draft_banner is printed in every tab's page header (the Summary also
    carries it on screen). Pass None only for a file that is genuinely not
    a draft -- nothing this package produces qualifies.
    """
    work_dir = request.output_path + ".workdir"
    ox.extract_template(request.template_path, work_dir)
    adopt_prior_output(work_dir, sample_sheet)

    # Refuse a template this builder cannot write correctly, before writing
    # anything. Across a population of controls a wrong-but-produced
    # workpaper costs far more than a refusal -- see check_template.
    sheets = ox.check_template(
        work_dir,
        required_sheets=[sample_sheet, population_sheet],
        # Formats are copied from the template or appended (see
        # _resolve_styles), so no particular count is required -- only a
        # styles.xml that exists.
        max_style_index=0,
    )
    sample_sheet_num = sheets[sample_sheet]
    population_sheet_num = sheets[population_sheet]
    parameters_sheet_num = sheets.get(parameters_sheet)

    sst = ox.SharedStrings(work_dir)

    sample_sheet_path = os.path.join(work_dir, "xl", "worksheets", f"sheet{sample_sheet_num}.xml")
    pop_sheet_path = os.path.join(work_dir, "xl", "worksheets", f"sheet{population_sheet_num}.xml")

    styles = _resolve_styles(work_dir, sample_sheet_path, pop_sheet_path, sst.strings)
    build_population_sheet_xml(pop_sheet_path, request.population_rows, sst, styles=styles)

    # ONE TAB PER SAMPLED ITEM. The prior-year workpaper documents a single
    # selection on its Sample tab -- narrative, raw data row, exhibits,
    # tickmarks -- so five selections means five tabs shaped like that one,
    # not five rows crammed onto it. Stacking them also produced the large
    # empty regions a reviewer flagged, because each sample's exhibits were
    # pushed below the previous sample's.
    #
    # A single sample keeps the tab named exactly as the template had it, so
    # the common case is indistinguishable from last year's file.
    unplaced: list[tuple[str, str, str, str]] = []
    sheets_for_sample: dict[str, str] = {}
    column_layout = ox.parse_column_widths(sample_sheet_path)
    # Where the PY put its narrative and exhibits -- read now, before the
    # tab's drawing is replaced. See agent/ooxml/layout.py.
    from agent.ooxml.layout import read_template_layout

    template_layout = read_template_layout(work_dir, sample_sheet_path, column_layout)

    for i, item in enumerate(request.sample_items):
        label = sample_tab_name(sample_sheet, item.sample_id, i, len(request.sample_items))
        if i == 0:
            num, path = sample_sheet_num, sample_sheet_path
            if label != sample_sheet:
                ox.rename_sheet(work_dir, num, label)
        else:
            num = ox.clone_sheet(work_dir, sample_sheet_num, label)
            path = os.path.join(work_dir, "xl", "worksheets", f"sheet{num}.xml")
        sheets_for_sample[label] = label

        build_sample_sheet_xml(path, [item], sst, styles)
        drawing = db.build_sample_drawing([item], column_layout, layout=template_layout)
        _write_drawing_and_media(work_dir, num, drawing)
        unplaced.extend(drawing.unplaced_callouts)

    sst.write()

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
        on_unplaced(unplaced)

    if draft_banner:
        for num in ox.resolve_sheets(work_dir).values():
            ox.stamp_print_header(
                os.path.join(work_dir, "xl", "worksheets", f"sheet{num}.xml"), draft_banner
            )

    ox.repair_builtin_name_scopes(work_dir)
    ox.remove_orphan_media(work_dir)
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
