"""
Deterministic OOXML surgery utilities.

These operate on an unzipped .xlsx directory tree. Nothing here makes
judgment calls; everything is pure data transformation. See README.md for
why this bypasses openpyxl for writing.
"""
from __future__ import annotations
import os
import re
import shutil
import zipfile
import datetime as dt
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, unescape
from dataclasses import dataclass, field

EXCEL_EPOCH = dt.date(1899, 12, 30)
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def excel_serial(d) -> int:
    """Convert a date/datetime to an Excel 1900-system serial number."""
    if isinstance(d, dt.datetime):
        d = d.date()
    return (d - EXCEL_EPOCH).days


def col_letter_to_index(letter: str) -> int:
    """'A' -> 1, 'T' -> 20, etc."""
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch.upper()) - ord("A") + 1)
    return n


def col_index_to_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------------------
# Template extraction / repackaging
# ---------------------------------------------------------------------------

def extract_template(template_path: str, work_dir: str) -> str:
    """Unzip a template .xlsx into work_dir (overwriting). Returns work_dir."""
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir)
    with zipfile.ZipFile(template_path) as z:
        z.extractall(work_dir)
    return work_dir


def repackage(work_dir: str, output_path: str) -> None:
    """Zip work_dir back into a valid .xlsx. Preserves the directory
    structure; does not rely on system `zip` so it's portable."""
    if os.path.exists(output_path):
        os.remove(output_path)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _dirs, files in os.walk(work_dir):
            for fn in files:
                full = os.path.join(root, fn)
                arcname = os.path.relpath(full, work_dir)
                z.write(full, arcname)


def remove_calc_chain(work_dir: str) -> None:
    """Delete calcChain.xml and its references. Safe/recommended whenever a
    formula cell's address changes (e.g. IA Calculation grows from 2 to 3
    lines) — a stale calc chain referencing a moved formula cell is a common
    cause of Excel's "we found a problem" repair prompt. Excel rebuilds the
    chain automatically on open."""
    cc_path = os.path.join(work_dir, "xl", "calcChain.xml")
    if os.path.exists(cc_path):
        os.remove(cc_path)

    ct_path = os.path.join(work_dir, "[Content_Types].xml")
    data = open(ct_path, encoding="utf-8").read()
    data = data.replace(
        '<Override PartName="/xl/calcChain.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.calcChain+xml"/>', ""
    )
    open(ct_path, "w", encoding="utf-8").write(data)

    rels_path = os.path.join(work_dir, "xl", "_rels", "workbook.xml.rels")
    data = open(rels_path, encoding="utf-8").read()
    data = re.sub(
        r'<Relationship Id="[^"]*" Type="[^"]*relationships/calcChain"[^>]*/>',
        "", data,
    )
    open(rels_path, "w", encoding="utf-8").write(data)


# ---------------------------------------------------------------------------
# Shared strings
# ---------------------------------------------------------------------------

class SharedStrings:
    """Loads an existing sharedStrings.xml, lets you look up-or-append
    strings, and writes the result back. Excel text cells always reference
    this table by index (t="s"), never inline text, in a well-formed file
    produced this way."""

    def __init__(self, work_dir: str):
        self.path = os.path.join(work_dir, "xl", "sharedStrings.xml")
        self.strings: list[str | None] = []
        self._index: dict[str, int] = {}
        if os.path.exists(self.path):
            root = ET.fromstring(open(self.path, encoding="utf-8").read())
            ns = {"m": NS_MAIN}
            for si in root:
                t = si.find("m:t", ns)
                self.strings.append(t.text if t is not None else None)
        for i, s in enumerate(self.strings):
            self._index.setdefault(s if s is not None else "", i)

    def get_or_add(self, value: str | None) -> int:
        key = value if value is not None else ""
        if key in self._index:
            return self._index[key]
        idx = len(self.strings)
        self.strings.append(value)
        self._index[key] = idx
        return idx

    def _register_part(self) -> None:
        """Make sure the workbook actually points at sharedStrings.xml.

        Writing the file is not enough. A part that nothing references is
        invisible: the workbook has no relationship to it and
        [Content_Types].xml has no override, so a reader finds an EMPTY
        string table and every t="s" cell resolves out of range. The file
        opens to an IndexError in openpyxl and a repair prompt in Excel.

        This is not hypothetical for a population of templates. A workbook
        written by openpyxl uses INLINE strings and ships no
        sharedStrings.xml at all, so every cell this builder writes into
        such a template referenced a table that was never wired up. Excel
        -authored templates happen to carry one, which is why it went
        unnoticed against the single template this was built from.
        """
        work_dir = os.path.dirname(os.path.dirname(self.path))
        rels_path = os.path.join(work_dir, "xl", "_rels", "workbook.xml.rels")
        if os.path.exists(rels_path):
            rels = open(rels_path, encoding="utf-8").read()
            if "sharedStrings.xml" not in rels:
                rels = rels.replace(
                    "</Relationships>",
                    '<Relationship Id="rIdSharedStringsCY" Type="http://schemas.'
                    'openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
                    'Target="sharedStrings.xml"/></Relationships>',
                    1,
                )
                open(rels_path, "w", encoding="utf-8").write(rels)

        ct_path = os.path.join(work_dir, "[Content_Types].xml")
        if os.path.exists(ct_path):
            ct = open(ct_path, encoding="utf-8").read()
            if "/xl/sharedStrings.xml" not in ct:
                ct = ct.replace(
                    "</Types>",
                    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
                    "</Types>",
                    1,
                )
                open(ct_path, "w", encoding="utf-8").write(ct)

    def write(self) -> None:
        self._register_part()
        items = []
        for s in self.strings:
            if s is None or s == "":
                items.append('<si><t xml:space="preserve"></t></si>')
            else:
                items.append(f'<si><t xml:space="preserve">{escape(s)}</t></si>')
        xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<sst xmlns="{NS_MAIN}" count="{len(self.strings)}" '
            f'uniqueCount="{len(self.strings)}">' + "".join(items) + "</sst>"
        )
        open(self.path, "w", encoding="utf-8").write(xml)


# ---------------------------------------------------------------------------
# Cell XML generation
# ---------------------------------------------------------------------------

def cell_xml(ref: str, value, *, style: int | None = None,
             sst: SharedStrings | None = None,
             formula: str | None = None) -> str:
    """Build a single <c> element.

    - None -> empty cell (self-closing, no <v>)
    - str -> shared string (requires `sst`)
    - datetime.date/datetime -> Excel serial number, numeric
    - int/float -> numeric
    - formula -> adds <f>...</f> before <v> (value is the cached result)
    """
    s_attr = f' s="{style}"' if style is not None else ""
    if value is None and formula is None:
        return f'<c r="{ref}"{s_attr}/>'
    if isinstance(value, str):
        assert sst is not None, "shared strings table required for text cells"
        idx = sst.get_or_add(value)
        return f'<c r="{ref}"{s_attr} t="s"><v>{idx}</v></c>'
    if isinstance(value, (dt.date, dt.datetime)):
        v = excel_serial(value)
    else:
        v = value
    f_xml = f"<f>{escape(formula)}</f>" if formula else ""
    return f'<c r="{ref}"{s_attr}>{f_xml}<v>{v}</v></c>'


def row_xml(row_num: int, cells: list[str], *, spans: str = "1:20",
            extra_attrs: str = "") -> str:
    # No x14ac:dyDescent. It is a cosmetic row-height hint that Excel
    # recomputes anyway, and emitting it requires the sheet's root element
    # to declare xmlns:x14ac. Excel-authored workbooks usually do; others
    # do not -- and writing a prefixed attribute into a sheet that has not
    # bound the prefix produces XML Excel refuses to open. Across 172
    # templates of unknown provenance that is not a risk worth running for
    # a hint.
    return (f'<row r="{row_num}" spans="{spans}"{extra_attrs}>'
            + "".join(cells) + "</row>")


def replace_sheet_rows(sheet_xml_path: str, start_marker: str, end_marker: str,
                        new_rows_xml: str, new_dimension: str | None = None,
                        old_dimension: str | None = None) -> None:
    """Splice new row XML into an existing worksheet XML file, replacing
    everything between the first occurrence of start_marker and end_marker.
    Typical usage: start_marker='<row r="6"', end_marker='</sheetData>'."""
    data = open(sheet_xml_path, encoding="utf-8").read()
    start = data.find(start_marker)
    end = data.find(end_marker)
    assert start != -1 and end != -1, "markers not found in sheet XML"
    new_data = data[:start] + new_rows_xml + data[end:]
    if new_dimension and old_dimension:
        new_data = new_data.replace(old_dimension, new_dimension)
    open(sheet_xml_path, "w", encoding="utf-8").write(new_data)


# ---------------------------------------------------------------------------
# Column-width / cell-grid EMU math (for positioning callouts over cells,
# as opposed to over images -- see README for why the two use different
# px->EMU constants)
# ---------------------------------------------------------------------------

CELL_GRID_EMU_PER_PX = 9525  # fixed, 96 DPI, independent of any image DPI


def excel_col_width_to_px(width_chars: float) -> int:
    """Excel's standard approximation for the default Calibri 11 font
    (MDW=7px). Good to a few px -- fine for a highlight box."""
    return round(width_chars * 7 + 5)


@dataclass
class ColumnLayout:
    widths_px: dict[str, int]  # column letter -> pixel width
    default_px: int = 64       # ~8.43 chars, Excel's default column width

    def cumulative_x_px(self, up_to_col_exclusive: str) -> int:
        """Sum of pixel widths of every column strictly before the given
        column letter (i.e. the x-coordinate where that column starts)."""
        target = col_letter_to_index(up_to_col_exclusive)
        total = 0
        for i in range(1, target):
            letter = col_index_to_letter(i)
            total += self.widths_px.get(letter, self.default_px)
        return total

    def cell_box_px(self, col_start: str, col_end: str, row_top_px: int,
                     row_height_px: int) -> tuple[int, int, int, int]:
        """Pixel bounding box (x, y, w, h) spanning columns col_start..col_end
        inclusive, for a row starting at row_top_px."""
        x1 = self.cumulative_x_px(col_start)
        x2 = self.cumulative_x_px(col_index_to_letter(
            col_letter_to_index(col_end) + 1))
        return x1, row_top_px, x2 - x1, row_height_px


def parse_column_widths(sheet_xml_path: str) -> ColumnLayout:
    """Read <cols><col min max width .../></cols> from a worksheet XML and
    return pixel widths per column letter."""
    data = open(sheet_xml_path, encoding="utf-8").read()
    ns = {"m": NS_MAIN}
    root = ET.fromstring(data)
    cols_el = root.find("m:cols", ns)
    widths: dict[str, int] = {}
    if cols_el is not None:
        for col in cols_el:
            w = float(col.get("width"))
            px = excel_col_width_to_px(w)
            for i in range(int(col.get("min")), int(col.get("max")) + 1):
                if i > 512:  # skip the "rest of the sheet" catch-all ranges
                    break
                widths[col_index_to_letter(i)] = px
    return ColumnLayout(widths_px=widths)


def cell_box_emu(layout: ColumnLayout, col_start: str, col_end: str,
                  row_index_1based: int, row_height_px: int = 20,
                  header_rows_px: int = 0) -> tuple[int, int, int, int]:
    """EMU bounding box for a range of columns on a given data row (assumes
    uniform row_height_px above the target row, e.g. 20px = default 15pt).
    header_rows_px lets you account for a differently-sized header row if the
    template has one."""
    row_top_px = header_rows_px + (row_index_1based - 2) * row_height_px \
        if row_index_1based > 1 else 0
    # simplest correct case (matches this template): row 1 = header, row 2 = data
    row_top_px = row_height_px * (row_index_1based - 1)
    x, y, w, h = layout.cell_box_px(col_start, col_end, row_top_px, row_height_px)
    return (x * CELL_GRID_EMU_PER_PX, y * CELL_GRID_EMU_PER_PX,
            w * CELL_GRID_EMU_PER_PX, h * CELL_GRID_EMU_PER_PX)


# ---------------------------------------------------------------------------
# Template compatibility
# ---------------------------------------------------------------------------

class TemplateMismatch(ValueError):
    """The template does not have the shape this builder writes into."""


def resolve_sheets(work_dir: str) -> dict[str, int]:
    """{sheet name: N} for xl/worksheets/sheetN.xml, from the workbook itself.

    Sheet NAME is the stable thing; sheetN.xml numbering is not. The two are
    related only through workbook.xml -> workbook.xml.rels, and in the real
    PY template they already disagree: "IA Leadsheet" carries sheetId="4"
    while living in sheet1.xml. Hard-coding "the Sample tab is sheet2.xml"
    happens to hold for one control and is a coin flip across a population
    of them -- and when it is wrong it does not fail, it writes the sample
    rows into whatever tab happens to be second.

    Parsed with ElementTree rather than regex, deliberately. Attribute ORDER
    is not significant in XML and real files disagree about it: Excel writes
    <Relationship Id=... Target=...>, openpyxl writes Target= before Id=.
    A regex pinned to one order silently returns nothing for files written
    by the other, and "nothing" here means every sheet looks missing.
    """
    ns_r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    ns_main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    ns_pkg = "{http://schemas.openxmlformats.org/package/2006/relationships}"

    rels_root = ET.parse(os.path.join(work_dir, "xl", "_rels", "workbook.xml.rels")).getroot()
    target_by_rid = {
        rel.get("Id"): rel.get("Target", "")
        for rel in rels_root.findall(f"{ns_pkg}Relationship")
    }

    wb_root = ET.parse(os.path.join(work_dir, "xl", "workbook.xml")).getroot()
    sheets: dict[str, int] = {}
    for sheet in wb_root.iter(f"{ns_main}sheet"):
        name = sheet.get("name")
        # Target is relative ("worksheets/sheet1.xml") from Excel and
        # absolute ("/xl/worksheets/sheet1.xml") from openpyxl; only the
        # trailing sheetN.xml is load-bearing.
        target = target_by_rid.get(sheet.get(f"{ns_r}id"), "")
        num = re.search(r"sheet(\d+)\.xml$", target)
        if name and num:
            sheets[name] = int(num.group(1))
    return sheets


def check_template(work_dir: str, *, required_sheets: list[str], max_style_index: int) -> dict[str, int]:
    """Verify a template can take this builder's output; return its sheet map.

    Written for a population of 172 controls, not for one. Every assumption
    below is one that holds for the template this package was built from and
    may not hold for the next: the tab names, and the style indices lifted
    out of that file's styles.xml.

    Both fail SILENTLY if unchecked, which is the reason this exists. A
    missing tab writes evidence into the wrong sheet; a style index past the
    end of cellXfs makes Excel show a repair prompt that names nothing, or
    quietly renders the cell in some other format. Across 172 workpapers
    either one is worse than a refusal, because a refusal is a two-minute
    fix and a wrong workpaper is a finding.
    """
    import re as _re

    sheets = resolve_sheets(work_dir)
    missing = [name for name in required_sheets if name not in sheets]
    if missing:
        raise TemplateMismatch(
            f"template has no sheet named {missing} -- it has {sorted(sheets)}. "
            f"This builder writes into tabs by name; either the template is for a "
            f"differently-shaped control, or the tab names differ and the caller "
            f"should pass the names this template actually uses."
        )

    with open(os.path.join(work_dir, "xl", "styles.xml"), encoding="utf-8") as fh:
        styles = fh.read()
    m = _re.search(r'<cellXfs count="(\d+)"', styles)
    available = int(m.group(1)) if m else 0
    if available <= max_style_index:
        raise TemplateMismatch(
            f"template's styles.xml defines {available} cell formats, but this "
            f"builder addresses style index {max_style_index}. The STYLE_* "
            f"constants in build_workpaper.py were lifted from a specific "
            f"template's styles.xml and do not transfer to this one."
        )
    return sheets


def clone_sheet(work_dir: str, src_sheet_num: int, new_name: str) -> int:
    """Duplicate a worksheet part under a new tab name; return its number.

    Used to give every sampled item its own tab. The prior-year workpaper
    documents ONE selection on its Sample tab -- narrative, raw data row,
    exhibits, tickmarks -- so the honest way to document five selections is
    five tabs shaped like that one, not five rows crammed onto it. It also
    makes mirroring exact rather than approximate: PY's Sample layout is a
    layout for a single sample, and a single-sample tab can use it as-is.

    The clone deliberately drops the source's <drawing> reference and its
    relationships. Sharing a drawing part between two sheets would put
    sample 1's exhibits on sample 2's tab; each clone gets its own drawing
    written afterwards by _write_drawing_and_media.
    """
    sheets_dir = os.path.join(work_dir, "xl", "worksheets")
    existing = [f for f in os.listdir(sheets_dir) if re.fullmatch(r"sheet\d+\.xml", f)]
    new_num = max((int(re.findall(r"\d+", f)[0]) for f in existing), default=0) + 1

    with open(os.path.join(sheets_dir, f"sheet{src_sheet_num}.xml"), encoding="utf-8") as fh:
        body = fh.read()
    # Strip the inherited drawing link -- a fresh one is added per clone.
    body = re.sub(r"<drawing[^>]*/>", "", body)
    with open(os.path.join(sheets_dir, f"sheet{new_num}.xml"), "w", encoding="utf-8") as fh:
        fh.write(body)

    rid = f"rIdSampleClone{new_num}"
    rels_path = os.path.join(work_dir, "xl", "_rels", "workbook.xml.rels")
    with open(rels_path, encoding="utf-8") as fh:
        rels = fh.read()
    rels = rels.replace(
        "</Relationships>",
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{new_num}.xml"/>'
        "</Relationships>",
        1,
    )
    with open(rels_path, "w", encoding="utf-8") as fh:
        fh.write(rels)

    wb_path = os.path.join(work_dir, "xl", "workbook.xml")
    with open(wb_path, encoding="utf-8") as fh:
        wb = fh.read()
    used = [int(x) for x in re.findall(r'sheetId="(\d+)"', wb)] or [0]
    entry = f'<sheet name="{escape(new_name)}" sheetId="{max(used) + 1}" r:id="{rid}"/>'
    # Insert immediately after the LAST tab whose name starts the same way,
    # so cloned sample tabs stay together in tab order. Appending at the end
    # scattered them after Population and Parameters, which reads as a
    # mistake in a workpaper a reviewer pages through in order.
    prefix = new_name.rsplit(" ", 1)[0] if " " in new_name else new_name
    siblings = [
        m for m in re.finditer(r"<sheet\b[^>]*/>", wb)
        if re.search(rf'name="{re.escape(prefix)}[^"]*"', m.group(0))
    ]
    if siblings:
        at = siblings[-1].end()
        wb = wb[:at] + entry + wb[at:]
    else:
        wb = wb.replace("</sheets>", f"{entry}</sheets>", 1)
    with open(wb_path, "w", encoding="utf-8") as fh:
        fh.write(wb)

    ct_path = os.path.join(work_dir, "[Content_Types].xml")
    with open(ct_path, encoding="utf-8") as fh:
        ct = fh.read()
    ct = ct.replace(
        "</Types>",
        f'<Override PartName="/xl/worksheets/sheet{new_num}.xml" ContentType="application/'
        f'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        1,
    )
    with open(ct_path, "w", encoding="utf-8") as fh:
        fh.write(ct)
    return new_num


def rename_sheet(work_dir: str, sheet_num: int, new_name: str) -> None:
    """Rename the tab that points at sheet<N>.xml."""
    wb_path = os.path.join(work_dir, "xl", "workbook.xml")
    rels_path = os.path.join(work_dir, "xl", "_rels", "workbook.xml.rels")
    with open(rels_path, encoding="utf-8") as fh:
        rels = fh.read()
    rid = None
    for m in re.finditer(r'<Relationship\b[^>]*/>', rels):
        tag = m.group(0)
        if re.search(rf'Target="[^"]*sheet{sheet_num}\.xml"', tag):
            got = re.search(r'Id="([^"]+)"', tag)
            if got:
                rid = got.group(1)
            break
    if not rid:
        return
    with open(wb_path, encoding="utf-8") as fh:
        wb = fh.read()
    wb = re.sub(
        rf'(<sheet )name="[^"]*"([^>]*r:id="{re.escape(rid)}")',
        rf'\1name="{escape(new_name)}"\2',
        wb,
        count=1,
    )
    with open(wb_path, "w", encoding="utf-8") as fh:
        fh.write(wb)
