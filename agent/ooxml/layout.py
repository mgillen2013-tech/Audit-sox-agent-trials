"""Where exhibits go on a Sample tab: read from the PY, never invented.

See docs/layout_must_be_mirrored_not_computed.md. The PY workpaper's own
drawing says where its narrative box and each exhibit sat and how big
they were. This module reads those positions ("slots") and places this
year's exhibits into them.

The builder this replaces stacked every exhibit at full page size in ONE
column down the left. A real output had its fourth exhibit starting 29
inches down the sheet -- around 140 rows of scrolling -- while the PY for
the same control laid six exhibits out three across and two down.

Three cases, in order of preference:

1. The PY has exhibit slots: fill them in reading order, each exhibit
   shrunk to fit its slot without stretching. If CY has more exhibits than
   PY had slots, further rows continue the PY's own rhythm (its typical
   slot size, its number across).
2. The template's drawing is one THIS BUILDER made -- a previous CY
   workpaper used as the prior year. Its layout is the old single column,
   and mirroring it would preserve exactly the problem being fixed, so it
   is ignored in favour of (3).
3. No slots to mirror: a plain grid, three across, each exhibit fitted to
   a fixed cell. Same shape as the PY layouts, so it reads the same way.

Everything here is pure arithmetic on rectangles, so it is testable
without images, OCR or Excel.
"""

from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from agent.ooxml.ooxml_utils import CELL_GRID_EMU_PER_PX, ColumnLayout, col_index_to_letter

EMU_PER_INCH = 914_400
EMU_PER_PT = 12_700
GAP = 228_600  # 0.25in between exhibits

# The fallback grid: three across, each exhibit fitted inside 5.5in x 7in.
# A letter-size page at 7in tall is about 5.4in wide -- legible at 100%
# zoom, sharp when zoomed (the image keeps its full resolution), and three
# fit across a normal monitor. The PY layouts measured use slots of about
# this size or smaller.
GRID_COLUMNS = 3
GRID_CELL_W = int(5.5 * EMU_PER_INCH)
GRID_CELL_H = int(7.0 * EMU_PER_INCH)
GRID_TOP = 500_000  # below the header and the sampled-item row

# Shape names only this builder produces. Their presence means the drawing
# being read is a previous CY output, not a PY layout.
_GENERATED_MARKERS = ("TextBox narrative_", "Rectangle raw_", "Tickmark raw_")

_NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


@dataclass(frozen=True)
class Slot:
    x: int
    y: int
    cx: int
    cy: int

    @property
    def bottom(self) -> int:
        return self.y + self.cy

    @property
    def right(self) -> int:
        return self.x + self.cx

    def overlaps(self, other: "Slot") -> bool:
        return not (
            self.right <= other.x or other.right <= self.x
            or self.bottom <= other.y or other.bottom <= self.y
        )


@dataclass
class TemplateLayout:
    narrative: Slot | None = None
    images: list[Slot] = field(default_factory=list)
    # True when the drawing is a previous output of this builder; its
    # positions are then deliberately NOT mirrored.
    generated: bool = False

    @property
    def usable(self) -> bool:
        return bool(self.images) and not self.generated


# ---------------------------------------------------------------------------
# Reading the PY drawing
# ---------------------------------------------------------------------------


@dataclass
class _Grid:
    """Converts a cell anchor (col, colOff, row, rowOff) to sheet EMU."""

    columns: ColumnLayout
    default_row_emu: int
    row_emu: dict[int, int]  # 0-based row -> height

    def x(self, col: int, off: int) -> int:
        px = sum(
            self.columns.widths_px.get(col_index_to_letter(i + 1), self.columns.default_px)
            for i in range(col)
        )
        return px * CELL_GRID_EMU_PER_PX + off

    def y(self, row: int, off: int) -> int:
        return sum(self.row_emu.get(r, self.default_row_emu) for r in range(row)) + off


def _row_heights(sheet_xml: str) -> tuple[int, dict[int, int]]:
    default = re.search(r'<sheetFormatPr\b[^>]*\bdefaultRowHeight="([\d.]+)"', sheet_xml)
    default_emu = int(float(default.group(1)) * EMU_PER_PT) if default else 15 * EMU_PER_PT
    custom = {
        int(r) - 1: int(float(ht) * EMU_PER_PT)
        for r, ht in re.findall(r'<row\b[^>]*?\br="(\d+)"[^>]*?\bht="([\d.]+)"', sheet_xml)
    }
    return default_emu, custom


def _drawing_path(work_dir: str, sheet_xml_path: str) -> str | None:
    folder, name = os.path.split(sheet_xml_path)
    rels = os.path.join(folder, "_rels", f"{name}.rels")
    if not os.path.exists(rels):
        return None
    with open(rels, encoding="utf-8") as fh:
        m = re.search(r'Target="[^"]*?drawings/(drawing\d+\.xml)"', fh.read())
    if not m:
        return None
    path = os.path.join(work_dir, "xl", "drawings", m.group(1))
    return path if os.path.exists(path) else None


def _anchor_box(anchor: ET.Element, grid: _Grid) -> Slot | None:
    tag = anchor.tag.split("}")[-1]

    def cell(el: ET.Element | None) -> tuple[int, int, int, int] | None:
        if el is None:
            return None
        vals = [el.find(f"xdr:{k}", _NS) for k in ("col", "colOff", "row", "rowOff")]
        if any(v is None or v.text is None for v in vals):
            return None
        return tuple(int(v.text) for v in vals)  # type: ignore[return-value]

    if tag == "absoluteAnchor":
        pos, ext = anchor.find("xdr:pos", _NS), anchor.find("xdr:ext", _NS)
        if pos is None or ext is None:
            return None
        return Slot(int(pos.get("x")), int(pos.get("y")), int(ext.get("cx")), int(ext.get("cy")))

    start = cell(anchor.find("xdr:from", _NS))
    if start is None:
        return None
    x0, y0 = grid.x(start[0], start[1]), grid.y(start[2], start[3])
    if tag == "oneCellAnchor":
        ext = anchor.find("xdr:ext", _NS)
        if ext is None:
            return None
        return Slot(x0, y0, int(ext.get("cx")), int(ext.get("cy")))
    end = cell(anchor.find("xdr:to", _NS))
    if end is None:
        return None
    x1, y1 = grid.x(end[0], end[1]), grid.y(end[2], end[3])
    return Slot(x0, y0, max(0, x1 - x0), max(0, y1 - y0))


def read_template_layout(work_dir: str, sheet_xml_path: str, columns: ColumnLayout) -> TemplateLayout:
    """The PY Sample tab's narrative and exhibit slots, in reading order.

    Must be called BEFORE the tab's drawing is replaced. Never raises: a
    drawing it cannot make sense of yields an empty layout, and the caller
    falls back to the grid -- an unusual template is a layout question, not
    a reason to lose the workpaper.
    """
    try:
        drawing = _drawing_path(work_dir, sheet_xml_path)
        if drawing is None:
            return TemplateLayout()
        with open(sheet_xml_path, encoding="utf-8") as fh:
            default_row, rows = _row_heights(fh.read())
        grid = _Grid(columns, default_row, rows)
        root = ET.parse(drawing).getroot()
    except Exception:  # noqa: BLE001
        return TemplateLayout()

    names = [el.get("name", "") for el in root.iter(f"{{{_NS['xdr']}}}cNvPr")]
    generated = any(n.startswith(m) for n in names for m in _GENERATED_MARKERS)

    images: list[Slot] = []
    narrative: Slot | None = None
    narrative_area = 0
    for anchor in root:
        # Only shapes that are DIRECT children of an anchor. A picture inside
        # a group is positioned relative to the group's own transform, and
        # reading it as if it were absolute would put the slot somewhere else.
        box = _anchor_box(anchor, grid)
        if box is None or box.cx <= 0 or box.cy <= 0:
            continue
        if anchor.find("xdr:pic", _NS) is not None:
            images.append(box)
            continue
        sp = anchor.find("xdr:sp", _NS)
        if sp is None:
            continue
        text = "".join(t.text or "" for t in sp.iter(f"{{{_NS['a']}}}t"))
        # The narrative is the biggest box with real prose in it. Tickmark
        # letters ("A") and empty outline rectangles never qualify.
        if len(text.strip()) >= 40 and box.cx * box.cy > narrative_area:
            narrative, narrative_area = box, box.cx * box.cy

    return TemplateLayout(narrative=narrative, images=_reading_order(images), generated=generated)


def _reading_order(slots: list[Slot]) -> list[Slot]:
    """Top-to-bottom bands, left-to-right within a band.

    A band is every slot whose top lies within an inch of the band's first
    top -- PY exhibits in one visual row rarely share an exact top edge.
    """
    bands: list[list[Slot]] = []
    for s in sorted(slots, key=lambda s: (s.y, s.x)):
        if bands and s.y - bands[-1][0].y <= EMU_PER_INCH:
            bands[-1].append(s)
        else:
            bands.append([s])
    return [s for band in bands for s in sorted(band, key=lambda s: s.x)]


# ---------------------------------------------------------------------------
# Placing this year's exhibits
# ---------------------------------------------------------------------------


def _fit(w: int, h: int, cell_w: int, cell_h: int) -> tuple[int, int]:
    """Largest size of a w x h image that fits inside the cell, unstretched."""
    s = min(cell_w / w, cell_h / h)
    return max(1, int(w * s)), max(1, int(h * s))


def _median(values: list[int]) -> int:
    ordered = sorted(values)
    return ordered[len(ordered) // 2]


def plan_sample_layout(
    layout: TemplateLayout,
    image_sizes: list[tuple[int, int]],
    *,
    narrative_height_for_width,
    default_narrative_width: int,
) -> tuple[Slot, list[Slot]]:
    """(narrative box, one box per image) in sheet EMU.

    image_sizes are each image's pixel (width, height); only the ratio
    matters. narrative_height_for_width(cx) returns the height the text
    needs at that width -- text in a box narrower than the PY's must be
    allowed to grow rather than be clipped.
    """
    if layout.usable:
        slots = list(layout.images)
        if layout.narrative is not None:
            n = layout.narrative
            narrative = Slot(n.x, n.y, n.cx, max(n.cy, narrative_height_for_width(n.cx)))
        else:
            # The PY had exhibits but no narrative box. Take the free space
            # left of the first exhibit in the top band rather than a fixed
            # width that sits on top of it and pushes that exhibit down a
            # row. Below 3in the text would be a ribbon; keep the default
            # width then and let the exhibit move.
            top_band = [s for s in slots if s.y - slots[0].y <= EMU_PER_INCH]
            free = min(s.x for s in top_band) - GAP
            width = min(default_narrative_width, free) if free >= 3 * EMU_PER_INCH else default_narrative_width
            narrative = Slot(0, min(GRID_TOP, slots[0].y), width, narrative_height_for_width(width))
        cell_w = _median([s.cx for s in slots])
        cell_h = _median([s.cy for s in slots])
        across = max(2, _widest_band(slots))
    else:
        slots = []
        narrative = Slot(0, GRID_TOP, default_narrative_width,
                         narrative_height_for_width(default_narrative_width))
        cell_w, cell_h, across = GRID_CELL_W, GRID_CELL_H, GRID_COLUMNS
        # The first row shares the top band with the narrative: exhibits
        # start to its right, exactly where the PY layouts put them.
        first_x = narrative.right + GAP
        for k in range(across - 1):
            slots.append(Slot(first_x + k * (cell_w + GAP), GRID_TOP, cell_w, cell_h))

    placed: list[Slot] = []
    overflow = 0
    for i, (w, h) in enumerate(image_sizes):
        if i < len(slots):
            slot = slots[i]
        else:
            # Beyond the PY's slots: continue its rhythm (its typical slot
            # size, its number across). Each exhibit sits directly under
            # whatever is already in ITS column -- not under the tallest
            # thing in the row, which left a blank patch under a short
            # narrative box.
            col = overflow % across
            overflow += 1
            x = col * (cell_w + GAP)
            above = [b.bottom for b in (narrative, *placed, *slots[len(placed):])
                     if b.x < x + cell_w and x < b.right]
            slot = Slot(x, (max(above) + GAP) if above else GRID_TOP, cell_w, cell_h)
        fw, fh = _fit(w, h, slot.cx, slot.cy)
        box = Slot(slot.x, slot.y, fw, fh)
        # Never on top of the narrative or of another exhibit. Text that
        # outgrew its PY box, or a PY whose slots overlapped, is resolved by
        # moving this exhibit straight down until it is clear.
        moved = True
        while moved:
            moved = False
            for other in (narrative, *placed):
                if box.overlaps(other):
                    box = Slot(box.x, other.bottom + GAP, box.cx, box.cy)
                    moved = True
        placed.append(box)
    return narrative, placed


def _widest_band(slots: list[Slot]) -> int:
    widest, band_top, count = 0, None, 0
    for s in slots:  # already in reading order
        if band_top is None or s.y - band_top > EMU_PER_INCH:
            band_top, count = s.y, 0
        count += 1
        widest = max(widest, count)
    return widest
