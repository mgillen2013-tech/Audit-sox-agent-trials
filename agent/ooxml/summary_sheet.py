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

_COLUMNS = [
    ("A", "Test Step", 46),
    ("B", "Attributes", 30),
    ("C", "Attributes Evidenced?", 62),
    ("D", "Result", 12),
    ("E", "Evidence", 16),
]


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
        if any(w in lowered for w in ("exception", "not satisfied", "fail")):
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
    )
    f_bold, f_green, f_red, f_amber, f_link = (n_fonts + i for i in range(5))

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
    )
    s_header, s_text, s_green, s_red, s_amber, s_link = (n_xfs + i for i in range(6))

    for tag, count, addition in (
        ("fonts", n_fonts + 5, new_fonts),
        ("fills", n_fills + 4, new_fills),
        ("borders", n_borders + 1, new_borders),
        ("cellXfs", n_xfs + 6, new_xfs),
    ):
        xml = re.sub(rf'<{tag} count="\d+"', f'<{tag} count="{count}"', xml, count=1)
        xml = xml.replace(f"</{tag}>", f"{addition}</{tag}>", 1)

    with open(path, "w", encoding="utf-8") as fh:
        fh.write(xml)

    return _Styles(s_header, s_text, s_green, s_red, s_amber, s_link)


def _sheet_xml(
    rows: list[SummaryRow],
    sst: ox.SharedStrings,
    styles: _Styles,
    existing_sheets: "set[str] | None" = None,
) -> str:
    cols = "".join(
        f'<col min="{ox.col_letter_to_index(c)}" max="{ox.col_letter_to_index(c)}" '
        f'width="{w}" customWidth="1"/>'
        for c, _t, w in _COLUMNS
    )

    body = [
        ox.row_xml(
            1,
            [
                ox.cell_xml(f"{c}1", title, style=styles.header, sst=sst)
                for c, title, _w in _COLUMNS
            ],
            spans="1:5",
            extra_attrs=' ht="20" customHeight="1"',
        )
    ]

    hyperlinks = []
    fill_for = {"green": styles.green, "red": styles.red, "amber": styles.amber}

    for i, row in enumerate(rows, start=2):
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
        ]
        # Row height scaled to the longest stacked column, so a reviewer
        # sees every attribute without resizing anything.
        lines = max(len(row.attributes), len(row.evidenced), 1)
        body.append(
            ox.row_xml(i, cells, spans="1:5", extra_attrs=f' ht="{max(30, lines * 15)}" customHeight="1"')
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
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="A1:E{len(rows) + 1}"/>'
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f"<cols>{cols}</cols>"
        f"<sheetData>{''.join(body)}</sheetData>"
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

    # Next free sheetN.xml, so this never collides with the template's own.
    sheet_dir = os.path.join(work_dir, "xl", "worksheets")
    existing = [f for f in os.listdir(sheet_dir) if re.fullmatch(r"sheet\d+\.xml", f)]
    n = max((int(re.findall(r"\d+", f)[0]) for f in existing), default=0) + 1

    with open(os.path.join(sheet_dir, f"sheet{n}.xml"), "w", encoding="utf-8") as fh:
        fh.write(_sheet_xml(rows, sst, styles, existing_sheet_names))

    rels_path = os.path.join(work_dir, "xl", "_rels", "workbook.xml.rels")
    with open(rels_path, encoding="utf-8") as fh:
        rels = fh.read()
    # A relationship id that cannot collide with the template's rId1..rIdN.
    rid = "rIdSummary"
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
