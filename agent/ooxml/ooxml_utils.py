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
from xml.sax.saxutils import escape
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

    def write(self) -> None:
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
    return (f'<row r="{row_num}" spans="{spans}" '
            f'x14ac:dyDescent="0.25"{extra_attrs}>' + "".join(cells) + "</row>")


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
