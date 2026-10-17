"""A fifth tab: what was tested, what proved it, and did it pass.

The PY template has four tabs and no summary, because a human author holds
the whole control in their head while they build it. A generated workpaper
does not get that benefit -- a reviewer opens it cold, and the first thing
they need is not the evidence, it is the answer.

DESIGN CONSTRAINT, and the reason this file is opinionated: the agent's
own output is too long to paste here. It writes a multi-paragraph
narrative and an evidence citation carrying a full quote for every item.
That is the right amount of detail for the working papers underneath and
the wrong amount for a summary -- a summary a reviewer has to READ is not
a summary. So this tab is built from the SHORTEST fields the agent
produces (attribute names, observed values, verdicts) and never from the
narrative. Anything a reviewer needs in full is one hyperlink away.

Two rows earn their place by name rather than by being generic:
  - one per test step, because that is the unit a control is signed off on
  - one for IPE, because completeness and accuracy of the population is a
    separate assertion from any test step, and burying it inside one makes
    it invisible to a reviewer scanning for it.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from xml.sax.saxutils import escape

from agent.ooxml import ooxml_utils as ox

# Excel's standard conditional-formatting palette. Reusing the colours a
# reviewer already reads as pass/fail everywhere else in Excel costs
# nothing and means the tab needs no legend of its own.
_GREEN_BG, _GREEN_FG = "FFC6EFCE", "FF006100"
_RED_BG, _RED_FG = "FFFFC7CE", "FF9C0006"
_AMBER_BG, _AMBER_FG = "FFFFEB9C", "FF9C6500"
_HEADER_BG = "FFD9D9D9"

# Shown across the top of the Summary and in every tab's printed header.
# The standalone builder (agent/workpaper.py) has carried this since the
# start; the PY-mirroring path did not, so the file that looked MOST
# finished was the only one with no draft marking at all.
DRAFT_BANNER = "DRAFT -- AI-prepared, pending human reviewer approval. Not a finalized workpaper."

# Column F carries the model's prose. It used to sit in the narrative box
# on each sample tab, where it repeated per sample and pushed the legend
# down; on the summary it is written once per test step, next to the rows
# it explains. Column G exists only when the reviewer pass is enabled.
_COLUMNS = [
    ("A", "Test Step", 46),
    ("B", "Attributes", 30),
    ("C", "Attributes Evidenced?", 62),
    ("D", "Result", 12),
    ("E", "Evidence", 16),
    ("F", "Conclusion", 52),
]
_REVIEW_COLUMN = ("G", "Reviewer check", 40)

# Row 1 is the DRAFT banner, row 2 the headings.
_FIRST_DATA_ROW = 3

# How a previous run's Summary is recognised when a past CY workpaper is
# used as the template: by its headings, which no PY tab shares. The tab
# NAME is not enough -- a real PY workpaper may have a "Summary" of its own,
# and that one must never be deleted.
GENERATED_HEADINGS = tuple(title for _c, title, _w in _COLUMNS[:3])


@dataclass
class SummaryRow:
    """One line of the summary. `link_to` is a cell reference on another
    sheet ("Sample!A1"); the Evidence cell becomes an internal hyperlink to
    it, which is what makes the tab navigable rather than merely readable.
    """

    test_step: str
    attributes: list[str] = field(default_factory=list)
    evidenced: list[str] = field(default_factory=list)
    result: str = ""  # "Satisfied" | "Exception" | "Not tested" ...
    evidence_label: str = ""
    link_to: str = ""
    # The model's own prose, written once per test step rather than on
    # every row -- repeating it beside each selection is noise.
    narrative: str = ""
    # Populated only when the reviewer pass ran; empty string means the
    # column is not shown at all, which is different from "reviewed and
    # found nothing".
    review: str = ""

    @property
    def verdict_colour(self) -> str:
        """green / red / amber, from the result text.

        Defaults to AMBER, never green. An unrecognised verdict rendering
        as "pass" is the one failure mode that actively misleads a
        reviewer, so anything this function does not understand is shown
        as needing attention.
        """
        lowered = self.result.strip().lower()
        if lowered.startswith("satisf"):
            return "green"
        if any(w in lowered for w in ("exception", "not satisfied", "fail", "not updated")):
            return "red"
        return "amber"


@dataclass
class _Styles:
    header: int
    text: int
    green: int
    red: int
    amber: int
    link: int
    banner: int


def _add_styles(work_dir: str) -> _Styles:
    """Append the fills, fonts and cell formats the summary needs.

    The template carries no green or red fill -- it has three fills, one of
    them yellow -- so they have to be added rather than looked up. Existing
    indices are never touched: everything is appended and the new indices
    returned, because every other sheet in the workbook addresses its
    styles by position and renumbering would silently restyle the whole
    file.
    """
    path = os.path.join(work_dir, "xl", "styles.xml")
    with open(path, encoding="utf-8") as fh:
        xml = fh.read()

    def _count(tag: str) -> int:
        m = re.search(rf'<{tag} count="(\d+)"', xml)
        return int(m.group(1)) if m else 0

    n_fonts, n_fills, n_borders, n_xfs = (
        _count("fonts"), _count("fills"), _count("borders"), _count("cellXfs")
    )

    new_fonts = (
        f'<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/></font>'
        f'<font><sz val="11"/><color rgb="{_GREEN_FG}"/><name val="Calibri"/></font>'
        f'<font><sz val="11"/><color rgb="{_RED_FG}"/><name val="Calibri"/></font>'
        f'<font><sz val="11"/><color rgb="{_AMBER_FG}"/><name val="Calibri"/></font>'
        f'<font><u/><sz val="11"/><color theme="10"/><name val="Calibri"/></font>'
        f'<font><b/><sz val="12"/><color rgb="{_RED_FG}"/><name val="Calibri"/></font>'
    )
    f_bold, f_green, f_red, f_amber, f_link, f_banner = (n_fonts + i for i in range(6))

    new_fills = "".join(
        f'<fill><patternFill patternType="solid"><fgColor rgb="{bg}"/>'
        f'<bgColor indexed="64"/></patternFill></fill>'
        for bg in (_HEADER_BG, _GREEN_BG, _RED_BG, _AMBER_BG)
    )
    fl_header, fl_green, fl_red, fl_amber = (n_fills + i for i in range(4))

    thin = '<left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/>'
    new_borders = f"<border>{thin}<diagonal/></border>"
    b_thin = n_borders

    # wrapText + top alignment throughout: these cells hold several short
    # lines, and without it Excel shows one line and hides the rest.
    def xf(font: int, fill: int, *, centre: bool = False) -> str:
        align = (
            f'<alignment horizontal="{"center" if centre else "left"}" '
            f'vertical="top" wrapText="1"/>'
        )
        return (
            f'<xf numFmtId="0" fontId="{font}" fillId="{fill}" borderId="{b_thin}" '
            f'applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">{align}</xf>'
        )

    new_xfs = (
        xf(f_bold, fl_header, centre=True)
        + xf(0, 0)
        + xf(f_green, fl_green, centre=True)
        + xf(f_red, fl_red, centre=True)
        + xf(f_amber, fl_amber, centre=True)
        + xf(f_link, 0, centre=True)
        + xf(f_banner, fl_red)
    )
    s_header, s_text, s_green, s_red, s_amber, s_link, s_banner = (n_xfs + i for i in range(7))

    for tag, count, addition in (
        ("fonts", n_fonts + 6, new_fonts),
        ("fills", n_fills + 4, new_fills),
        ("borders", n_borders + 1, new_borders),
        ("cellXfs", n_xfs + 7, new_xfs),
    ):
        xml = re.sub(rf'<{tag} count="\d+"', f'<{tag} count="{count}"', xml, count=1)
        xml = xml.replace(f"</{tag}>", f"{addition}</{tag}>", 1)

    with open(path, "w", encoding="utf-8") as fh:
        fh.write(xml)

    return _Styles(s_header, s_text, s_green, s_red, s_amber, s_link, s_banner)


def _sheet_xml(
    rows: list[SummaryRow],
    sst: ox.SharedStrings,
    styles: _Styles,
    existing_sheets: "set[str] | None" = None,
) -> str:
    columns = list(_COLUMNS)
    if any(r.review for r in rows):
        columns.append(_REVIEW_COLUMN)
    last_col = columns[-1][0]
    cols = "".join(
        f'<col min="{ox.col_letter_to_index(c)}" max="{ox.col_letter_to_index(c)}" '
        f'width="{w}" customWidth="1"/>'
        for c, _t, w in columns
    )

    spans = f"1:{ox.col_letter_to_index(last_col)}"
    body = [
        # Row 1 is the draft banner, merged across the table. It is the
        # first thing on the first tab, so nobody opens this file without
        # seeing it.
        ox.row_xml(
            1,
            [ox.cell_xml("A1", DRAFT_BANNER, style=styles.banner, sst=sst)]
            + [ox.cell_xml(f"{c}1", None, style=styles.banner, sst=sst) for c, _t, _w in columns[1:]],
            spans=spans,
            extra_attrs=' ht="22" customHeight="1"',
        ),
        ox.row_xml(
            2,
            [
                ox.cell_xml(f"{c}2", title, style=styles.header, sst=sst)
                for c, title, _w in columns
            ],
            spans=spans,
            extra_attrs=' ht="20" customHeight="1"',
        ),
    ]

    hyperlinks = []
    fill_for = {"green": styles.green, "red": styles.red, "amber": styles.amber}

    for i, row in enumerate(rows, start=_FIRST_DATA_ROW):
        verdict_style = fill_for[row.verdict_colour]
        target_sheet = row.link_to.split("!")[0].strip("'") if row.link_to else ""
        linkable = bool(row.link_to) and (existing_sheets is None or target_sheet in existing_sheets)
        cells = [
            ox.cell_xml(f"A{i}", row.test_step, style=styles.text, sst=sst),
            ox.cell_xml(f"B{i}", "\n".join(row.attributes), style=styles.text, sst=sst),
            ox.cell_xml(f"C{i}", "\n".join(row.evidenced), style=styles.text, sst=sst),
            ox.cell_xml(f"D{i}", row.result, style=verdict_style, sst=sst),
            ox.cell_xml(
                f"E{i}",
                row.evidence_label or "-",
                style=styles.link if linkable else styles.text,
                sst=sst,
            ),
            ox.cell_xml(f"F{i}", row.narrative, style=styles.text, sst=sst),
        ]
        if len(columns) > len(_COLUMNS):
            cells.append(ox.cell_xml(f"G{i}", row.review, style=styles.text, sst=sst))
        # Row height scaled to the longest stacked column, so a reviewer
        # sees every attribute without resizing anything.
        lines = max(len(row.attributes), len(row.evidenced), 1)
        body.append(
            ox.row_xml(i, cells, spans=f"1:{ox.col_letter_to_index(last_col)}", extra_attrs=f' ht="{max(30, lines * 15)}" customHeight="1"')
        )
        # Drop a link whose target tab is not in this workbook. Checked
        # HERE, at the only layer that knows which sheets exist, so no
        # caller can ship a broken one: the IPE row links to "Parameters",
        # which some templates simply do not have. Excel does not complain
        # about a dangling internal link -- the click just does nothing,
        # which reads to a reviewer as the workpaper being broken.
        if linkable:
            hyperlinks.append(
                f'<hyperlink ref="E{i}" location="{escape(row.link_to)}" '
                f'display="{escape(row.evidence_label)}"/>'
            )

    links_xml = f"<hyperlinks>{''.join(hyperlinks)}</hyperlinks>" if hyperlinks else ""
    merge_xml = f'<mergeCells count="1"><mergeCell ref="A1:{last_col}1"/></mergeCells>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="A1:{last_col}{len(rows) + _FIRST_DATA_ROW - 1}"/>'
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f"<cols>{cols}</cols>"
        f"<sheetData>{''.join(body)}</sheetData>"
        f"{merge_xml}"
        f"{links_xml}"
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        "</worksheet>"
    )


def add_summary_sheet(
    work_dir: str, rows: list[SummaryRow], sst: ox.SharedStrings, *, name: str = "Summary"
) -> None:
    """Write the summary sheet into an unzipped workbook and register it.

    A new sheet is not one file. Excel refuses to open the workbook -- with
    a repair prompt that names nothing useful -- if any of the four
    registrations below is missing, so they are done together here rather
    than left to the caller: the sheet part itself, the workbook
    relationship, the <sheet> entry that gives it a name and tab position,
    and the content-type override.

    Placed FIRST in the tab order deliberately. It is the answer; the
    evidence behind it is what the other tabs are for.
    """
    styles = _add_styles(work_dir)
    # Named for what it holds. A plain `existing` collided with the
    # sheet-FILENAME list further down, which silently overwrote it and
    # made every hyperlink target look absent -- so the tab shipped with
    # no working links at all.
    existing_sheet_names = set(ox.resolve_sheets(work_dir))
    # A PY workpaper may already have a "Summary" of its own (a previous
    # run's is removed before this point -- see adopt_prior_output). Two
    # tabs with one name is a workbook Excel will not open, so ours steps
    # aside rather than the template's being touched.
    if name in existing_sheet_names:
        name = next(
            candidate for candidate in (f"CY {name}", *(f"CY {name} {i}" for i in range(2, 100)))
            if candidate not in existing_sheet_names
        )

    # Next free sheetN.xml, so this never collides with the template's own.
    sheet_dir = os.path.join(work_dir, "xl", "worksheets")
    existing = [f for f in os.listdir(sheet_dir) if re.fullmatch(r"sheet\d+\.xml", f)]
    n = max((int(re.findall(r"\d+", f)[0]) for f in existing), default=0) + 1

    # Only ONE tab may be selected. The template arrives with its own
    # selected tab; leaving that flag on as well opens the workbook with
    # both tabs GROUPED, and anything typed then lands on both of them.
    for fn in existing:
        path = os.path.join(sheet_dir, fn)
        with open(path, encoding="utf-8") as fh:
            xml = fh.read()
        if 'tabSelected="1"' in xml:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(xml.replace(' tabSelected="1"', ""))

    with open(os.path.join(sheet_dir, f"sheet{n}.xml"), "w", encoding="utf-8") as fh:
        fh.write(_sheet_xml(rows, sst, styles, existing_sheet_names))

    rels_path = os.path.join(work_dir, "xl", "_rels", "workbook.xml.rels")
    with open(rels_path, encoding="utf-8") as fh:
        rels = fh.read()
    # A relationship id that cannot collide with the template's rId1..rIdN.
    rid = "rIdSummary"
    while f'Id="{rid}"' in rels:
        rid += "X"
    rels = rels.replace(
        "</Relationships>",
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{n}.xml"/>'
        "</Relationships>",
        1,
    )
    with open(rels_path, "w", encoding="utf-8") as fh:
        fh.write(rels)

    wb_path = os.path.join(work_dir, "xl", "workbook.xml")
    with open(wb_path, encoding="utf-8") as fh:
        wb = fh.read()
    used_ids = [int(x) for x in re.findall(r'sheetId="(\d+)"', wb)] or [0]
    entry = f'<sheet name="{escape(name)}" sheetId="{max(used_ids) + 1}" r:id="{rid}"/>'
    wb = wb.replace("<sheets>", f"<sheets>{entry}", 1)
    # Everything that addressed a tab by position now addresses the one to
    # its left unless shifted -- see shift_tab_positions. Then make the
    # Summary the tab the file opens on.
    wb = ox.shift_tab_positions(wb, 0, +1)
    wb = re.sub(r'\bactiveTab="\d+"', 'activeTab="0"', wb)
    with open(wb_path, "w", encoding="utf-8") as fh:
        fh.write(wb)

    ct_path = os.path.join(work_dir, "[Content_Types].xml")
    with open(ct_path, encoding="utf-8") as fh:
        ct = fh.read()
    ct = ct.replace(
        "</Types>",
        f'<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/'
        f'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        1,
    )
    with open(ct_path, "w", encoding="utf-8") as fh:
        fh.write(ct)


def add_table_sheet(
    work_dir: str,
    name: str,
    columns: list[tuple[str, str, int]],
    rows: list[list[str]],
    sst: ox.SharedStrings,
    *,
    position: int = 1,
    red_rows: set[int] | None = None,
    amber_rows: set[int] | None = None,
) -> str:
    """A plain table tab (headings + rows) inserted at a tab position.

    Registered the same four ways add_summary_sheet registers its tab. Used
    for the Mirror Check: it sits right after the Summary so a reviewer sees
    what to look at before the evidence.
    """
    styles = _add_styles(work_dir)
    existing = set(ox.resolve_sheets(work_dir))
    final = name
    k = 2
    while final in existing:
        final, k = f"{name} {k}", k + 1

    sheet_dir = os.path.join(work_dir, "xl", "worksheets")
    files = [f for f in os.listdir(sheet_dir) if re.fullmatch(r"sheet\d+\.xml", f)]
    n = max((int(re.findall(r"\d+", f)[0]) for f in files), default=0) + 1
    last = columns[-1][0]
    cols = "".join(f'<col min="{ox.col_letter_to_index(c)}" max="{ox.col_letter_to_index(c)}" '
                   f'width="{w}" customWidth="1"/>' for c, _t, w in columns)
    body = [ox.row_xml(1, [ox.cell_xml(f"{c}1", t, style=styles.header, sst=sst) for c, t, _w in columns],
                       spans=f"1:{ox.col_letter_to_index(last)}")]
    for i, row in enumerate(rows, start=2):
        style = (styles.red if red_rows and i - 2 in red_rows
                 else styles.amber if amber_rows and i - 2 in amber_rows else styles.text)
        cells = [ox.cell_xml(f"{c}{i}", (row[j] if j < len(row) else "") or None,
                             style=style if j == 0 else styles.text, sst=sst)
                 for j, (c, _t, _w) in enumerate(columns)]
        body.append(ox.row_xml(i, cells, spans=f"1:{ox.col_letter_to_index(last)}"))
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           f'<dimension ref="A1:{last}{len(rows) + 1}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>'
           f'<sheetFormatPr defaultRowHeight="15"/><cols>{cols}</cols><sheetData>{"".join(body)}</sheetData>'
           '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
    with open(os.path.join(sheet_dir, f"sheet{n}.xml"), "w", encoding="utf-8") as fh:
        fh.write(xml)

    rels_path = os.path.join(work_dir, "xl", "_rels", "workbook.xml.rels")
    with open(rels_path, encoding="utf-8") as fh:
        rels = fh.read()
    rid = f"rIdTable{n}"
    rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                        f'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{n}.xml"/></Relationships>', 1)
    with open(rels_path, "w", encoding="utf-8") as fh:
        fh.write(rels)

    wb_path = os.path.join(work_dir, "xl", "workbook.xml")
    with open(wb_path, encoding="utf-8") as fh:
        wb = fh.read()
    used = [int(x) for x in re.findall(r'sheetId="(\d+)"', wb)] or [0]
    entry = f'<sheet name="{escape(final)}" sheetId="{max(used) + 1}" r:id="{rid}"/>'
    tabs = list(re.finditer(r"<sheet\b[^>]*/>", wb))
    if position < len(tabs):
        at = tabs[position].start()
        wb = wb[:at] + entry + wb[at:]
        wb = ox.shift_tab_positions(wb, position, +1)
    else:
        wb = wb.replace("</sheets>", f"{entry}</sheets>", 1)
    with open(wb_path, "w", encoding="utf-8") as fh:
        fh.write(wb)

    ct_path = os.path.join(work_dir, "[Content_Types].xml")
    with open(ct_path, encoding="utf-8") as fh:
        ct = fh.read()
    ct = ct.replace("</Types>", f'<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/'
                    f'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>', 1)
    with open(ct_path, "w", encoding="utf-8") as fh:
        fh.write(ct)
    return final
