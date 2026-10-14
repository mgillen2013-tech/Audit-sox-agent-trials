"""Everything in a PY workbook that belongs to last year, with where it is.

Pure reading; nothing is changed here. The result is what the rewrite call
is shown and what the apply step edits, so every item carries the exact
span of XML it came from.

NOTHING HERE DECIDES WHAT AN ELEMENT MEANS. Kinds are physical facts --
what the shape is, not what it is for:
  text       a shape with text in it (a conclusion, a note, a "TM 2" label)
  outline    a shape with an outline and no text (a box around a value)
  frame      an outline coinciding with a picture (it moves with the picture)
  connector  a line or arrow
  group      grouped shapes (their text is read for the model)
  picture    an exhibit
  object     an embedded file (an email, typically)
  other      charts and anything else
What each one MEANS -- a tickmark, a conclusion, last year's evidence -- is
read by the model from the PY (agent/mirror/structure.py, rewrite.py).

TabInventory.role is only a HINT from tab names, used when the model's
structure read is unavailable; it never overrides it.
"""

from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from agent.ooxml import ooxml_utils as ox
from agent.ooxml.layout import Slot, _anchor_box, _Grid, _row_heights

NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS = {"xdr": NS_XDR, "a": NS_A}

# Top-level anchors in a drawing part. They never nest, so a lazy match is
# exact. mc:AlternateContent wraps one anchor (embedded objects, newer shapes).
_ANCHOR_RE = re.compile(
    r"<((?:\w+:)?(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)|mc:AlternateContent)\b.*?</\1>",
    re.S,
)
_LABEL_RE = re.compile(r"(?:TM\s*\d+[a-z]?|[A-Z]{1,2}\d?|\d{1,2}|[A-Z]\.?)", re.I)


@dataclass
class StyleClass:
    """One distinct run format inside a text box (font, size, bold, colour)."""

    key: str
    rpr_xml: str  # the <a:rPr .../> element, verbatim from the PY
    bold: bool = False
    color: str = ""


@dataclass
class Paragraph:
    ppr_xml: str                       # <a:pPr .../> or ""
    runs: list[tuple[str, str]]        # (style key, text)
    end_rpr_xml: str = ""              # <a:endParaRPr .../> or ""

    @property
    def text(self) -> str:
        return "".join(t for _k, t in self.runs)


@dataclass
class DrawingItem:
    id: str              # "<cNvPr id>", unique within its drawing
    name: str
    kind: str
    box: Slot
    span: tuple[int, int]  # character span of the whole anchor in the drawing XML
    paragraphs: list[Paragraph] = field(default_factory=list)
    embed: str = ""        # picture relationship id
    image_path: str = ""   # PY picture file, for reading what it showed

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs).strip()


@dataclass
class CellItem:
    ref: str
    text: str
    formula: bool
    style: int | None


@dataclass
class TabInventory:
    name: str
    sheet_num: int
    role: str
    drawing_part: str = ""          # "drawing2.xml"
    drawing_xml: str = ""
    items: list[DrawingItem] = field(default_factory=list)
    styles: dict[str, StyleClass] = field(default_factory=dict)
    cells: dict[str, CellItem] = field(default_factory=dict)

    def by_kind(self, *kinds: str) -> list[DrawingItem]:
        return [i for i in self.items if i.kind in kinds]


@dataclass
class Inventory:
    work_dir: str
    tabs: list[TabInventory]

    def role(self, role: str) -> list[TabInventory]:
        return [t for t in self.tabs if t.role == role]


# ---------------------------------------------------------------------------


def _ns_decls(drawing_xml: str) -> str:
    root = re.search(r"<(?:\w+:)?wsDr\b([^>]*)>", drawing_xml)
    decls = " ".join(re.findall(r'xmlns(?::\w+)?="[^"]*"', root.group(1))) if root else ""
    for prefix, uri in (("xdr", NS_XDR), ("a", NS_A),
                        ("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
                        ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
                        ("a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
                        ("a16", "http://schemas.microsoft.com/office/drawing/2014/main")):
        if f"xmlns:{prefix}=" not in decls:
            decls += f' xmlns:{prefix}="{uri}"'
    return decls


def _style_key(rpr_xml: str) -> str:
    """Formatting that matters to a reader, ignoring bookkeeping attributes
    (lang, dirty, err, baseline=0) that differ between otherwise identical
    runs and would split one visual style into several."""
    s = re.sub(r'\s(?:lang|altLang|dirty|err|noProof|smtClean)="[^"]*"', "", rpr_xml)
    s = re.sub(r'\sbaseline="0"', "", s)
    return s


def _runs_of(p_xml: str) -> tuple[str, list[tuple[str, str]], str]:
    ppr = re.search(r"<a:pPr\b[^>]*?(?:/>|>.*?</a:pPr>)", p_xml, re.S)
    end = re.search(r"<a:endParaRPr\b[^>]*?(?:/>|>.*?</a:endParaRPr>)", p_xml, re.S)
    runs = []
    for r in re.finditer(r"<a:r>(.*?)</a:r>|<a:fld\b.*?</a:fld>|<a:br\b[^>]*/>", p_xml, re.S):
        if r.group(0).startswith("<a:br"):
            runs.append(("<br>", "\n"))
            continue
        body = r.group(1) or r.group(0)
        rpr = re.search(r"<a:rPr\b[^>]*?(?:/>|>.*?</a:rPr>)", body, re.S)
        t = re.search(r"<a:t>(.*?)</a:t>|<a:t\b[^>]*>(.*?)</a:t>|<a:t/>", body, re.S)
        text = ""
        if t:
            text = ox.unescape(t.group(1) if t.group(1) is not None else (t.group(2) or ""),
                               {"&quot;": '"', "&apos;": "'"})
        runs.append((rpr.group(0) if rpr else "", text))
    return (ppr.group(0) if ppr else ""), runs, (end.group(0) if end else "")


def _parse_text(sp_xml: str, styles: dict[str, StyleClass]) -> list[Paragraph]:
    body = re.search(r"<(?:\w+:)?txBody>(.*?)</(?:\w+:)?txBody>", sp_xml, re.S)
    if not body:
        return []
    paragraphs = []
    for p in re.finditer(r"<a:p>(.*?)</a:p>|<a:p/>", body.group(1), re.S):
        ppr, raw_runs, end = _runs_of(p.group(1) or "")
        runs: list[tuple[str, str]] = []
        for rpr, text in raw_runs:
            if rpr == "<br>":
                runs.append(("br", "\n"))
                continue
            key = _style_key(rpr)
            if key not in styles:
                styles[key] = StyleClass(
                    key=f"s{len(styles)}", rpr_xml=rpr,
                    bold=bool(re.search(r'\bb="1"', rpr)),
                    color=(re.search(r'srgbClr val="([0-9A-Fa-f]{6})"', rpr) or [None, ""])[1],
                )
            k = styles[key].key
            if runs and runs[-1][0] == k:
                runs[-1] = (k, runs[-1][1] + text)   # merge same-format runs
            else:
                runs.append((k, text))
        paragraphs.append(Paragraph(ppr, runs, end))
    return paragraphs


def _classify(sp: ET.Element, text: str, box: Slot, pictures: list[Slot]) -> str:
    """What the shape physically is. Never what it is for."""
    nv = sp.find(".//xdr:cNvPr", _NS)
    name = nv.get("name", "") if nv is not None else ""
    if (nv is not None and nv.get("hidden") == "1") or name.lower().startswith("object"):
        return "object"
    if text.strip():
        return "text"
    outline = sp.find(".//xdr:spPr/a:ln", _NS)
    if outline is None or outline.find("a:noFill", _NS) is not None:
        return "other"
    for pic in pictures:
        if _overlap_ratio(box, pic) > 0.8:
            return "frame"
    return "outline"


def _overlap_ratio(a: Slot, b: Slot) -> float:
    ix = max(0, min(a.right, b.right) - max(a.x, b.x))
    iy = max(0, min(a.bottom, b.bottom) - max(a.y, b.y))
    inter = ix * iy
    union = a.cx * a.cy + b.cx * b.cy - inter
    return inter / union if union else 0.0


def read_drawing(work_dir: str, sheet_num: int, drawing_part: str) -> tuple[str, list[DrawingItem], dict[str, StyleClass]]:
    path = os.path.join(work_dir, "xl", "drawings", drawing_part)
    with open(path, encoding="utf-8") as fh:
        xml = fh.read()
    sheet_path = os.path.join(work_dir, "xl", "worksheets", f"sheet{sheet_num}.xml")
    with open(sheet_path, encoding="utf-8") as fh:
        default_h, custom_h = _row_heights(fh.read())
    grid = _Grid(ox.parse_column_widths(sheet_path), default_h, custom_h)
    decls = _ns_decls(xml)

    rels: dict[str, str] = {}
    rels_path = os.path.join(work_dir, "xl", "drawings", "_rels", f"{drawing_part}.rels")
    if os.path.exists(rels_path):
        with open(rels_path, encoding="utf-8") as fh:
            for rel in re.finditer(r"<Relationship\b[^>]*/>", fh.read()):
                rid = re.search(r'\bId="([^"]+)"', rel.group(0))
                tgt = re.search(r'\bTarget="([^"]+)"', rel.group(0))
                if rid and tgt:
                    rels[rid.group(1)] = ox._part_path(work_dir, "xl/drawings", tgt.group(1))

    parsed = []
    for m in _ANCHOR_RE.finditer(xml):
        try:
            root = ET.fromstring(f"<root {decls}>{m.group(0)}</root>")
        except ET.ParseError:
            continue
        anchor = root[0]
        if anchor.tag.endswith("AlternateContent"):
            choice = anchor[0] if len(anchor) else None
            anchor = choice[0] if choice is not None and len(choice) else None
        if anchor is None:
            continue
        box = _anchor_box(anchor, grid)
        parsed.append((m, anchor, box))

    pictures = [b for _m, a, b in parsed if b is not None and a.find("xdr:pic", _NS) is not None]
    styles: dict[str, StyleClass] = {}
    items: list[DrawingItem] = []
    for m, anchor, box in parsed:
        if box is None:
            box = Slot(0, 0, 0, 0)
        shape = next((c for c in anchor if c.tag.split("}")[-1] in ("sp", "pic", "grpSp", "graphicFrame", "cxnSp")), None)
        if shape is None:
            continue
        tag = shape.tag.split("}")[-1]
        nv = shape.find(".//xdr:cNvPr", _NS)
        sid = nv.get("id", "") if nv is not None else ""
        name = nv.get("name", "") if nv is not None else ""
        if tag == "pic":
            blip = shape.find(".//a:blip", _NS)
            rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed", "") if blip is not None else ""
            items.append(DrawingItem(sid, name, "picture", box, m.span(), embed=rid, image_path=rels.get(rid, "")))
        elif tag == "sp":
            sp_xml = m.group(0)
            paragraphs = _parse_text(sp_xml, styles)
            text = "\n".join(p.text for p in paragraphs)
            kind = _classify(shape, text, box, pictures)
            items.append(DrawingItem(sid, name, kind, box, m.span(), paragraphs=paragraphs))
        elif tag == "cxnSp":
            items.append(DrawingItem(sid, name, "connector", box, m.span()))
        elif tag == "grpSp":
            # Read for the model (a grouped screenshot-with-callouts is PY
            # evidence like any other), never edited inside.
            group_text = " ".join(t.text or "" for t in shape.iter(f"{{{NS_A}}}t")).strip()
            n_pics = len(shape.findall(".//xdr:pic", _NS))
            item = DrawingItem(sid, name, "group", box, m.span())
            item.paragraphs = [Paragraph("", [("g", f"[group of shapes: {n_pics} picture(s)] {group_text}".strip())])]
            items.append(item)
        else:
            items.append(DrawingItem(sid, name, "other", box, m.span()))
    return xml, items, {v.key: v for v in styles.values()}


def _drawing_part_of(work_dir: str, sheet_num: int) -> str:
    rels = os.path.join(work_dir, "xl", "worksheets", "_rels", f"sheet{sheet_num}.xml.rels")
    if not os.path.exists(rels):
        return ""
    with open(rels, encoding="utf-8") as fh:
        m = re.search(r'Target="[^"]*?drawings/(drawing\d+\.xml)"', fh.read())
    return m.group(1) if m and os.path.exists(os.path.join(work_dir, "xl", "drawings", m.group(1))) else ""


def _tab_role(name: str, items: list[DrawingItem], cell_count: int) -> str:
    low = name.lower().strip()
    if re.fullmatch(r"\d+", low) or re.fullmatch(r"sample(?:\s+.*\d.*)?", low):
        return "selection"
    if "population" in low:
        return "population"
    if "parameter" in low:
        return "parameters"
    if "lead" in low or (
        not any(i.kind == "picture" for i in items)
        and any(i.kind == "text" and re.search(r"control description|test step|procedure", i.text, re.I) for i in items)
    ):
        return "lead"
    if low.endswith(">") or (cell_count == 0 and not items):
        return "divider"
    if "selection" in low or "sample" in low:
        return "selection_table"
    return "other"


def build_inventory(work_dir: str) -> Inventory:
    """Read a PY workbook already unzipped into work_dir."""
    sst = ox.SharedStrings(work_dir).strings
    from agent.ooxml.build_workpaper import _template_cells

    tabs: list[TabInventory] = []
    for name, num in ox.resolve_sheets(work_dir).items():
        part = _drawing_part_of(work_dir, num)
        xml, items, styles = ("", [], {})
        if part:
            try:
                xml, items, styles = read_drawing(work_dir, num, part)
            except Exception:  # noqa: BLE001 -- an unreadable drawing is left as it is
                xml, items, styles = ("", [], {})
        raw = _template_cells(os.path.join(work_dir, "xl", "worksheets", f"sheet{num}.xml"), sst)
        cells = {
            ref: CellItem(ref, str(text), formula, style)
            for ref, (style, text, formula) in raw.items()
            if str(text).strip() or formula
        }
        role = _tab_role(name, items, len(cells))
        tabs.append(TabInventory(name, num, role, part, xml, items, styles, cells))

    # No per-selection tabs: the tab whose rows look most like a listing is
    # the selection table (a "Sample Selection" tab, typically).
    if not any(t.role == "selection" for t in tabs) and not any(t.role == "selection_table" for t in tabs):
        best = max((t for t in tabs if t.role == "other"), key=lambda t: len(t.cells), default=None)
        if best is not None and best.cells:
            best.role = "selection_table"
    return Inventory(work_dir, tabs)
