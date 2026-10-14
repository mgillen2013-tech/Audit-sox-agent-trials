"""Write the CY edits into a copy of the PY. No judgment here -- only edits.

Every change is made IN the PY's own XML: a text box keeps its anchor,
fill, outline, font and paragraph settings and only its runs change; a
picture keeps its slot and is fitted inside it; a tickmark keeps its shape
and offset from its label and only moves. Anything the plan does not
mention is left exactly as the PY had it.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Any
from xml.sax.saxutils import escape

from agent.mirror.inventory import DrawingItem, Paragraph, TabInventory
from agent.mirror.rewrite import Exhibit, TabPlan, parse_tagged
from agent.ooxml import imaging
from agent.ooxml import ooxml_utils as ox
from agent.ooxml.layout import GAP, Slot, _fit

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass
class ApplyReport:
    unplaced: list[str] = field(default_factory=list)
    removed: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    # (item id, PY text, CY text, PY lead-ins, CY lead-ins) for verification
    text_changes: list[tuple] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Anchors
# ---------------------------------------------------------------------------


def _shape_xml(anchor_xml: str) -> str:
    m = re.search(r"<((?:\w+:)?(?:sp|pic|cxnSp))\b.*?</\1>", anchor_xml, re.S)
    return m.group(0) if m else ""


def _absolute(anchor_xml: str, box: Slot) -> str:
    """The same shape, re-anchored at an absolute position and size."""
    shape = _shape_xml(anchor_xml)
    shape = re.sub(r"<a:off\b[^>]*/>", f'<a:off x="{box.x}" y="{box.y}"/>', shape, count=1)
    shape = re.sub(r"<a:ext\b(?=[^>]*\bcx=)[^>]*/>", f'<a:ext cx="{box.cx}" cy="{box.cy}"/>', shape, count=1)
    return (f'<xdr:absoluteAnchor><xdr:pos x="{box.x}" y="{box.y}"/>'
            f'<xdr:ext cx="{box.cx}" cy="{box.cy}"/>{shape}<xdr:clientData/></xdr:absoluteAnchor>')


# ---------------------------------------------------------------------------
# Text
# ---------------------------------------------------------------------------


_PLAIN_RPR = '<a:rPr lang="en-US"/>'


def _run_xml(rpr: str, text: str) -> str:
    parts = text.split("\n")
    out = []
    for i, part in enumerate(parts):
        if i:
            out.append("<a:br/>")
        if part:
            out.append(f'<a:r>{rpr or _PLAIN_RPR}<a:t xml:space="preserve">{escape(part)}</a:t></a:r>')
    return "".join(out)


def rebuild_text(anchor_xml: str, item: DrawingItem, tab: TabInventory, paragraphs: list[str]) -> str:
    """Replace a text box's paragraphs, keeping everything else about it.

    Each new paragraph takes the PY paragraph's settings at the same
    position (alignment, spacing); a paragraph beyond the PY's count takes
    the last one's. Each run takes the format of its style tag, which is a
    format lifted verbatim from this box.
    """
    rpr_for = {s.key: s.rpr_xml for s in tab.styles.values()}
    valid = set(rpr_for)
    py_paras: list[Paragraph] = item.paragraphs or [Paragraph("", [])]
    body_paras = []
    for i, tagged in enumerate(paragraphs):
        ref = py_paras[min(i, len(py_paras) - 1)]
        default = next((k for k, _t in ref.runs if k != "br"), next(iter(valid), "s0"))
        runs = parse_tagged(tagged, valid, default)
        if runs:
            body_paras.append(f"<a:p>{ref.ppr_xml}" + "".join(_run_xml(rpr_for.get(k, ""), t) for k, t in runs) + "</a:p>")
        else:
            end = ref.end_rpr_xml or next((p.end_rpr_xml for p in py_paras if p.end_rpr_xml), "")
            body_paras.append(f"<a:p>{ref.ppr_xml}{end}</a:p>")
    tx = re.search(r"(<(?:\w+:)?txBody>)(.*?)(</(?:\w+:)?txBody>)", anchor_xml, re.S)
    if not tx:
        return anchor_xml
    head = re.match(r"\s*(<a:bodyPr\b[^>]*?(?:/>|>.*?</a:bodyPr>))?\s*(<a:lstStyle\b[^>]*?(?:/>|>.*?</a:lstStyle>))?", tx.group(2), re.S)
    body_pr = head.group(1) or "<a:bodyPr/>"
    lst = head.group(2) or "<a:lstStyle/>"
    # Longer than the PY text: let it spill visibly rather than be clipped
    # off the bottom of the box, where nobody would know it was there.
    new_len = sum(len(re.sub(r"</?s\d+>", "", p)) for p in paragraphs)
    if new_len > len(item.text) * 1.15:
        body_pr = body_pr.replace('vertOverflow="clip"', 'vertOverflow="overflow"')
    new_tx = f"{tx.group(1)}{body_pr}{lst}{''.join(body_paras)}{tx.group(3)}"
    return anchor_xml[: tx.start()] + new_tx + anchor_xml[tx.end():]


# ---------------------------------------------------------------------------
# Cells
# ---------------------------------------------------------------------------


def set_cells(sheet_path: str, values: dict[str, Any], sst: ox.SharedStrings,
              styles_by_ref: dict[str, int | None]) -> None:
    """Set (or clear, value None) individual cells, keeping each cell's
    format and every other cell of the sheet as it is.

    The PY cell decides the KIND of value too: where the PY held text, CY
    text is written ("07/01/2026" stays text); where it held a number or
    date, the CY value is written as one, so the PY's number or date format
    applies. Writing a date into a text-formatted PY cell shows a serial
    number; writing text into a date cell loses the format.
    """
    if not values:
        return
    with open(sheet_path, encoding="utf-8") as fh:
        xml = fh.read()
    xml = re.sub(r"<sheetData\s*/>", "<sheetData></sheetData>", xml, count=1)
    by_row: dict[int, dict[str, Any]] = {}
    for ref, v in values.items():
        m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
        if m:
            by_row.setdefault(int(m.group(2)), {})[m.group(1)] = v
    for row, cols in sorted(by_row.items()):
        row_re = re.compile(rf'<row\b[^>]*?\br="{row}"[^>]*?(?:/>|>(.*?)</row>)', re.S)
        m = row_re.search(xml)
        cells: dict[str, str] = {}
        if m and m.group(1):
            for c in re.finditer(r'<c\b[^>]*?\br="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', m.group(1), re.S):
                cells[c.group(1)] = c.group(0)
        for col, value in cols.items():
            ref = f"{col}{row}"
            old_style = re.search(r'\bs="(\d+)"', cells.get(col, ""))
            style = int(old_style.group(1)) if old_style else styles_by_ref.get(ref)
            was_text = bool(re.search(r'\bt="(?:s|str|inlineStr)"', cells.get(col, "")))
            if value is None or value == "":
                cells[col] = f'<c r="{ref}"' + (f' s="{style}"' if style is not None else "") + "/>"
            elif was_text and not isinstance(value, (int, float)) and not hasattr(value, "year"):
                cells[col] = ox.cell_xml(ref, str(value), style=style, sst=sst)
            else:
                cells[col] = ox.cell_xml(ref, _typed(value), style=style, sst=sst)
        ordered = "".join(cells[c] for c in sorted(cells, key=ox.col_letter_to_index))
        if m:
            open_tag = re.sub(r'\sspans="[^"]*"', "", re.match(r"<row\b[^>]*?(?=/?>)", m.group(0)).group(0))
            xml = xml[: m.start()] + f"{open_tag}>{ordered}</row>" + xml[m.end():]
        else:
            later = [x for x in re.finditer(r'<row\b[^>]*?\br="(\d+)"', xml) if int(x.group(1)) > row]
            at = later[0].start() if later else xml.index("</sheetData>")
            xml = xml[:at] + f'<row r="{row}">{ordered}</row>' + xml[at:]
    with open(sheet_path, "w", encoding="utf-8") as fh:
        fh.write(xml)


def _typed(value: Any) -> Any:
    """Numbers as numbers and dates as dates, so the PY cell's own number or
    date format applies; text as text."""
    if isinstance(value, (int, float)) or hasattr(value, "year"):
        return value
    s = str(value).strip()
    import datetime as _dt

    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return _dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    plain = s.replace(",", "").replace("$", "")
    if re.fullmatch(r"-?\d+(?:\.\d+)?", plain) and not (len(plain) > 1 and plain.startswith("0") and "." not in plain):
        return float(plain) if "." in plain else int(plain)
    return s


def recalc_on_open(work_dir: str, sheet_path: str) -> None:
    """Drop cached formula results on this sheet and ask Excel to
    recalculate on open, so an IA recalculation never shows PY numbers."""
    with open(sheet_path, encoding="utf-8") as fh:
        xml = fh.read()
    xml = re.sub(r"(<c\b[^>]*>\s*<f\b[^>]*?(?:/>|>.*?</f>))\s*<v>.*?</v>", r"\1", xml, flags=re.S)
    with open(sheet_path, "w", encoding="utf-8") as fh:
        fh.write(xml)
    wb_path = os.path.join(work_dir, "xl", "workbook.xml")
    with open(wb_path, encoding="utf-8") as fh:
        wb = fh.read()
    if "<calcPr" in wb:
        if "fullCalcOnLoad" not in wb:
            wb = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wb, count=1)
    else:
        wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>', 1) if "<extLst" not in wb else \
            wb.replace("<extLst", '<calcPr fullCalcOnLoad="1"/><extLst', 1)
    with open(wb_path, "w", encoding="utf-8") as fh:
        fh.write(wb)


# ---------------------------------------------------------------------------
# Embedded objects
# ---------------------------------------------------------------------------


def remove_embedded_objects(work_dir: str, sheet_num: int) -> bool:
    """Remove a tab's embedded files (the PY's attached emails). They are
    PY evidence; a CY workpaper carrying them would present last year's
    support as this year's."""
    sheet_path = os.path.join(work_dir, "xl", "worksheets", f"sheet{sheet_num}.xml")
    with open(sheet_path, encoding="utf-8") as fh:
        xml = fh.read()
    if "<oleObjects" not in xml and "oleObject" not in xml:
        return False
    xml = re.sub(r"<mc:AlternateContent\b(?:(?!</mc:AlternateContent>).)*?<oleObjects?\b.*?</mc:AlternateContent>", "", xml, flags=re.S)
    xml = re.sub(r"<oleObjects\b.*?</oleObjects>", "", xml, flags=re.S)
    # The VML that drew the object's icon goes too, unless it also holds
    # cell comments (which are not PY evidence and stay).
    rels_path = os.path.join(work_dir, "xl", "worksheets", "_rels", f"sheet{sheet_num}.xml.rels")
    if os.path.exists(rels_path):
        with open(rels_path, encoding="utf-8") as fh:
            rels = fh.read()
        for rel in re.finditer(r"<Relationship\b[^>]*/>", rels):
            tag = rel.group(0)
            rid = re.search(r'\bId="([^"]+)"', tag).group(1)
            target = re.search(r'\bTarget="([^"]+)"', tag).group(1)
            if "vmlDrawing" in target:
                vml_path = ox._part_path(work_dir, "xl/worksheets", target)
                vml = open(vml_path, encoding="utf-8", errors="ignore").read() if os.path.exists(vml_path) else ""
                if 'ObjectType="Note"' not in vml:
                    xml = re.sub(rf'<legacyDrawing\b[^>]*r:id="{re.escape(rid)}"[^>]*/>', "", xml)
                    rels = rels.replace(tag, "")
            elif target.endswith((".bin", ".emf", ".wmf")) and "printerSettings" not in target:
                if not re.search(rf'r:id="{re.escape(rid)}"', xml):
                    rels = rels.replace(tag, "")
        with open(rels_path, "w", encoding="utf-8") as fh:
            fh.write(rels)
    with open(sheet_path, "w", encoding="utf-8") as fh:
        fh.write(xml)
    return True


# ---------------------------------------------------------------------------
# One selection tab's drawing
# ---------------------------------------------------------------------------


def _crop(image, crop: dict | None, notes: list[str], label: str):
    """Cut a CY page down to the part the PY showed, as the model named it
    (first and last line of that part). Full width, so nothing to either
    side of a value is lost. Not found: the whole page, and a note."""
    if not crop or not (crop.get("from") or crop.get("to")):
        return image
    try:
        words = imaging.ocr_words(image)
    except Exception:  # noqa: BLE001
        return image
    tops = []
    for key in ("from", "to"):
        text = str(crop.get(key) or "").strip()
        if not text:
            continue
        token = text.split()[0]
        try:
            _x, y, _w, h = imaging.find_anchor_box(words, token)
            tops.append((y, y + h))
        except ValueError:
            notes.append(f"{label}: crop line {text!r} not found; the whole page is shown")
            return image
    if not tops:
        return image
    pad = 20
    y0 = max(0, min(t for t, _b in tops) - pad)
    y1 = min(image.height, max(b for _t, b in tops) + pad)
    return image.crop((0, y0, image.width, y1)) if y1 - y0 > 20 else image


@dataclass
class PlacedExhibit:
    exhibit: Exhibit
    box: Slot
    image: Any  # PIL image
    words: list | None = None


def _locate(placed: PlacedExhibit, anchor: str) -> Slot | None:
    """The anchor text's box on a placed exhibit, in sheet EMU."""
    if placed.words is None:
        try:
            placed.words = imaging.ocr_words(placed.image)
        except Exception:  # noqa: BLE001
            placed.words = []
    tokens = anchor.split()
    attempts = [(tokens[0], len(tokens) - 1)] if tokens else []
    attempts += [(t, 0) for t in tokens[1:]]
    for t in list(tokens):
        stripped = t.strip("$()").replace(",", "")
        if stripped and stripped != t:
            attempts.append((stripped, 0))
    for needle, extra in attempts:
        try:
            px, py, pw, ph = imaging.find_anchor_box(placed.words, needle, extra_words=extra)
        except ValueError:
            continue
        scale = placed.box.cx / placed.image.width
        return Slot(placed.box.x + int(px * scale), placed.box.y + int(py * scale),
                    max(1, int(pw * scale)), max(1, int(ph * scale)))
    return None


def _with_flips(xml: str, flip_h: bool, flip_v: bool) -> str:
    def fix(m: re.Match) -> str:
        tag = re.sub(r'\sflip[HV]="[^"]*"', "", m.group(0))
        attrs = (' flipH="1"' if flip_h else "") + (' flipV="1"' if flip_v else "")
        return tag[:-1] + attrs + ">" if not tag.endswith("/>") else tag[:-2] + attrs + "/>"
    return re.sub(r"<a:xfrm\b[^>]*>", fix, xml, count=1)


def apply_tab(
    work_dir: str,
    tab: TabInventory,
    plan: TabPlan,
    exhibits: dict[str, Exhibit],
    support_dir: str,
    *,
    data_row_cell_box,   # (heading) -> Slot | None
) -> ApplyReport:
    """Carry out the model's decisions for one tab's drawing. Nothing here
    decides what an element is; an element the plan does not mention is
    left exactly as the PY had it."""
    report = ApplyReport()
    xml = tab.drawing_xml
    rels_path = os.path.join(work_dir, "xl", "drawings", "_rels", f"{tab.drawing_part}.rels")
    rels = open(rels_path, encoding="utf-8").read() if os.path.exists(rels_path) else (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    media_dir = os.path.join(work_dir, "xl", "media")
    os.makedirs(media_dir, exist_ok=True)
    by_id = {i.id: i for i in tab.items}
    final: dict[str, str] = {}          # element id -> replacement anchor XML ("" = removed)
    appended: list[str] = []

    def anchor_of(item: DrawingItem) -> str:
        return xml[item.span[0]: item.span[1]]

    def _load(x: Exhibit):
        return imaging.load_and_prepare(os.path.join(support_dir, x.source_file), x.page)[0]

    def _new_media(image) -> str:
        n = 1
        while os.path.exists(os.path.join(media_dir, f"image{n}.png")):
            n += 1
        image.save(os.path.join(media_dir, f"image{n}.png"))
        rid = f"rIdMirror{n}"
        nonlocal rels
        rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                            f'officeDocument/2006/relationships/image" Target="../media/image{n}.png"/></Relationships>', 1)
        return rid

    # --- pictures: CY exhibits into the PY's slots (frames follow) ---
    placed: dict[str, PlacedExhibit] = {}
    new_box: dict[str, Slot] = {}
    frames = tab.by_kind("frame")
    pictures = tab.by_kind("picture")
    used_frames: set[str] = set()
    for pic in pictures:
        frame = next((f for f in frames if f.id not in used_frames and _overlap(f.box, pic.box) > 0.8), None)
        if frame:
            used_frames.add(frame.id)
        decision = plan.pictures.get(pic.id)
        if decision is None:
            report.unplaced.append(f"PY exhibit {pic.name!r} was not decided on; left in place (it is last year's)")
            continue
        exhibit = exhibits.get(decision["exhibit"]) if decision.get("exhibit") else None
        if exhibit is None:
            final[pic.id] = ""
            if frame:
                final[frame.id] = ""
            report.removed.append(f"PY exhibit {pic.name!r} (no CY counterpart)")
            continue
        try:
            image = _crop(_load(exhibit), decision.get("crop"), report.notes, f"{exhibit.source_file} p.{exhibit.page + 1}")
        except Exception as exc:  # noqa: BLE001
            final[pic.id] = ""
            if frame:
                final[frame.id] = ""
            report.unplaced.append(f"{exhibit.source_file} p.{exhibit.page + 1} could not be read: {exc}")
            continue
        w, h = _fit(image.width, image.height, pic.box.cx, pic.box.cy)
        box = Slot(pic.box.x, pic.box.y, w, h)
        rid = _new_media(image)
        final[pic.id] = _absolute(re.sub(r'r:embed="[^"]*"', f'r:embed="{rid}"', anchor_of(pic), count=1), box)
        if frame:
            final[frame.id] = _absolute(anchor_of(frame), box)
        placed[exhibit.id] = PlacedExhibit(exhibit, box, image)
        new_box[pic.id] = box

    # --- CY exhibits with no PY slot: below everything, the PY's slot size ---
    if plan.extra_exhibits:
        cell_w = sorted(p.box.cx for p in pictures)[len(pictures) // 2] if pictures else int(5.5 * 914_400)
        cell_h = sorted(p.box.cy for p in pictures)[len(pictures) // 2] if pictures else int(7 * 914_400)
        across = max(1, len({p.box.x // 914_400 for p in pictures})) if pictures else 2
        top = max([i.box.bottom for i in tab.items] + [500_000]) + GAP
        frame_xml = anchor_of(frames[0]) if frames else ""
        pic_xml = anchor_of(pictures[0]) if pictures else ""
        for n, x_id in enumerate(plan.extra_exhibits):
            exhibit = exhibits[x_id]
            try:
                image = _load(exhibit)
            except Exception as exc:  # noqa: BLE001
                report.unplaced.append(f"{exhibit.source_file} p.{exhibit.page + 1} could not be read: {exc}")
                continue
            row, col = divmod(n, across)
            w, h = _fit(image.width, image.height, cell_w, cell_h)
            box = Slot(col * (cell_w + GAP), top + row * (cell_h + GAP), w, h)
            rid = _new_media(image)
            if pic_xml:
                anchor = re.sub(r'r:embed="[^"]*"', f'r:embed="{rid}"', pic_xml, count=1)
                anchor = re.sub(r'(<(?:\w+:)?cNvPr\b[^>]*\bid=")\d+(")', rf"\g<1>{90000 + n}\2", anchor, count=1)
            else:
                anchor = _fresh_pic(rid, 90000 + n)
            appended.append(_absolute(anchor, box))
            if frame_xml:
                fr = re.sub(r'(<(?:\w+:)?cNvPr\b[^>]*\bid=")\d+(")', rf"\g<1>{91000 + n}\2", frame_xml, count=1)
                appended.append(_absolute(fr, box))
            placed[exhibit.id] = PlacedExhibit(exhibit, box, image)

    # --- elements: exactly what the model decided ---
    body: dict[str, str] = {}
    for eid, d in plan.elements.items():
        item = by_id.get(eid)
        if item is None:
            continue
        xml_ = anchor_of(item)
        if d.get("paragraphs") and item.kind == "text":
            xml_ = rebuild_text(xml_, item, tab, d["paragraphs"])
            report.text_changes.append((eid, item.text, "\n".join(
                re.sub(r"</?s\d+>", "", p).replace("<br/>", "\n") for p in d["paragraphs"]).strip(),
                _labels(item, tab), _cy_labels(d["paragraphs"], item, tab)))
        elif item.kind == "text":
            report.text_changes.append((eid, item.text, "" if d["action"] == "remove" else item.text,
                                        _labels(item, tab), _labels(item, tab)))
        body[eid] = xml_
        if d["action"] == "remove":
            final[eid] = ""
            report.removed.append(f"{item.kind} {item.name!r}")
        elif d["action"] in ("keep", "replace"):
            if xml_ != anchor_of(item):
                final[eid] = xml_
        if d["action"] == "keep" and item.kind == "object":
            report.notes.append(f"embedded file {item.name!r} kept: it is last year's -- attach the CY equivalent")

    # Placements: direct targets first, then "beside" (which needs the
    # other element's new position), then lines between elements.
    pending = {eid: d for eid, d in plan.elements.items() if d["action"] == "place" and eid in by_id}
    for _pass in range(4):
        for eid, d in list(pending.items()):
            item, on = by_id[eid], d["on"]
            target: Slot | None = None
            if "exhibit" in on:
                hit = _locate(placed[on["exhibit"]], on["anchor"]) if on["exhibit"] in placed else None
                if hit is not None:
                    pad = 45_720
                    rect = Slot(hit.x - pad, hit.y - pad, hit.cx + 2 * pad, hit.cy + 2 * pad)
                    target = rect if item.kind != "text" else Slot(rect.right + 45_720, rect.y, item.box.cx, item.box.cy)
            elif "heading" in on:
                cell = data_row_cell_box(on["heading"])
                if cell is not None:
                    target = cell if item.kind != "text" else Slot(cell.right + 45_720, cell.y, item.box.cx, item.box.cy)
            elif "beside" in on:
                other = on["beside"]
                if other in pending:
                    continue  # wait for it
                ref_new = new_box.get(other, by_id[other].box if other in by_id else None)
                if ref_new is not None:
                    target = _label_spot(item.box, by_id[other].box, ref_new)
            pending.pop(eid)
            if target is None:
                final[eid] = ""
                report.unplaced.append(
                    f"{item.kind} {item.text[:20] or item.name!r}: target {on} not found on the CY evidence, so it was removed")
                continue
            new_box[eid] = target
            final[eid] = _absolute(body.get(eid, anchor_of(item)), target)
    for eid in pending:  # circular "beside" references
        report.unplaced.append(f"element {eid}: placed beside an element that was never placed; left where it was")

    for eid, d in plan.elements.items():
        if d["action"] != "connect" or eid not in by_id:
            continue
        a = new_box.get(d["from"], by_id[d["from"]].box)
        b = new_box.get(d["to"], by_id[d["to"]].box)
        x1, y1 = a.x + a.cx // 2, a.y + a.cy // 2
        # To the nearest edge of the target, so the arrowhead touches it.
        x2 = b.x if x1 < b.x else (b.right if x1 > b.right else b.x + b.cx // 2)
        y2 = b.y if y1 < b.y else (b.bottom if y1 > b.bottom else b.y + b.cy // 2)
        box = Slot(min(x1, x2), min(y1, y2), max(1, abs(x2 - x1)), max(1, abs(y2 - y1)))
        final[eid] = _with_flips(_absolute(anchor_of(by_id[eid]), box), x2 < x1, y2 < y1)

    if any(by_id[e].kind == "object" for e, d in plan.elements.items() if d["action"] == "remove" and e in by_id):
        remove_embedded_objects(work_dir, tab.sheet_num)

    # Write back, last span first so earlier spans stay valid.
    for eid in sorted(final, key=lambda e: -by_id[e].span[0]):
        s0, e0 = by_id[eid].span
        xml = xml[:s0] + final[eid] + xml[e0:]
    if appended:
        close = re.search(r"</(?:\w+:)?wsDr>\s*$", xml)
        at = close.start() if close else len(xml)
        xml = xml[:at] + "".join(appended) + xml[at:]
    # Shapes written here use the xdr:, a: and r: prefixes. A drawing from a
    # non-Excel writer may declare none of them (a default namespace
    # instead); undeclared, they make the part malformed.
    root = re.search(r"<((?:\w+:)?wsDr)\b[^>]*>", xml)
    if root:
        decls = ""
        for prefix, uri in (("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"),
                            ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"), ("r", R_NS)):
            if f"xmlns:{prefix}=" not in root.group(0):
                decls += f' xmlns:{prefix}="{uri}"'
        if decls:
            xml = xml[: root.start()] + f"<{root.group(1)}{decls}" + xml[root.start() + 1 + len(root.group(1)):]

    with open(os.path.join(work_dir, "xl", "drawings", tab.drawing_part), "w", encoding="utf-8") as fh:
        fh.write(xml)
    os.makedirs(os.path.dirname(rels_path), exist_ok=True)
    with open(rels_path, "w", encoding="utf-8") as fh:
        fh.write(rels)
    return report


def _labels(item: DrawingItem, tab: TabInventory) -> list[str]:
    """Each paragraph's lead-in, where the PY set it apart by FORMAT (a bold
    or coloured first run followed by a differently formatted one) -- e.g.
    "Test Step 1: ". Derived from the PY's formatting, not from a list of
    words, so it works whatever a PY calls its sections."""
    out = []
    for p in item.paragraphs:
        runs = [(k, t) for k, t in p.runs if k != "br"]
        if len(runs) >= 2 and runs[0][0] != runs[1][0]:
            st = tab.styles.get(runs[0][0])
            if st and (st.bold or (st.color and st.color.upper() != "000000")):
                out.append(runs[0][1].strip())
    return out


def _cy_labels(paragraphs: list[str], item: DrawingItem, tab: TabInventory) -> list[str]:
    from agent.mirror.rewrite import parse_tagged

    valid = set(tab.styles)
    out = []
    for p in paragraphs:
        runs = parse_tagged(p, valid, "")
        if len(runs) >= 2 and runs[0][0] != runs[1][0]:
            st = tab.styles.get(runs[0][0])
            if st and (st.bold or (st.color and st.color.upper() != "000000")):
                out.append(runs[0][1].strip())
    return out


def _label_spot(py_label: Slot, py_box: Slot, new_box: Slot) -> Slot:
    """Where the label goes beside the moved box: on the SAME SIDE of it as
    in the PY, at the same gap. Copying the raw PY offset put labels on top
    of the value whenever the CY box was wider than the PY's."""
    lcx, lcy = py_label.x + py_label.cx / 2, py_label.y + py_label.cy / 2
    if lcx >= py_box.right:                       # right of the box
        x = new_box.right + max(0, py_label.x - py_box.right)
    elif lcx <= py_box.x:                         # left of the box
        x = new_box.x - max(0, py_box.x - py_label.right) - py_label.cx
    else:                                         # over it horizontally
        x = new_box.x + (py_label.x - py_box.x)
    if lcy >= py_box.bottom:                      # below
        y = new_box.bottom + max(0, py_label.y - py_box.bottom)
    elif lcy <= py_box.y:                         # above
        y = new_box.y - max(0, py_box.y - py_label.bottom) - py_label.cy
    else:                                         # level with it
        y = new_box.y + (new_box.cy - py_label.cy) // 2
    # Never on top of the value: if the PY label overlapped its box (both
    # axes "over"), move it to the right of the new box.
    candidate = Slot(int(max(0, x)), int(max(0, y)), py_label.cx, py_label.cy)
    if candidate.overlaps(new_box):
        candidate = Slot(new_box.right + 45_720, new_box.y + (new_box.cy - py_label.cy) // 2, py_label.cx, py_label.cy)
    return candidate


def _overlap(a: Slot, b: Slot) -> float:
    from agent.mirror.inventory import _overlap_ratio

    return _overlap_ratio(a, b)


def _fresh_pic(rid: str, shape_id: int) -> str:
    return (f'<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1" cy="1"/>'
            f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{shape_id}" name="Picture {shape_id}"/>'
            f'<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
            f'<xdr:blipFill><a:blip xmlns:r="{R_NS}" r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:absoluteAnchor>')
