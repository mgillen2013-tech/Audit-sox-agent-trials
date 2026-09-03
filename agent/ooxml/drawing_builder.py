"""
Builds xl/drawings/drawingN.xml content (images + callout rectangles +
narrative textboxes) from structured input, using xdr:absoluteAnchor
throughout.

Why absoluteAnchor everywhere: a oneCellAnchor/twoCellAnchor positions a
shape relative to a specific column/row, which means overlaying a callout on
top of an *image* (not a cell) would require you to first figure out which
column/row the image's anchor cell sits in, then add a colOff/rowOff -- and
that offset is only well-defined if you also account for that anchor cell's
own width, easy to get subtly wrong. absoluteAnchor instead positions
everything in EMU from the sheet's top-left corner, so "callout is 583px
right and 210px down from where this image starts" is just addition, once
you know the image's own absolute (x, y).

The tradeoff: absoluteAnchor shapes do NOT move if someone inserts/deletes
rows or columns above them. That's the right tradeoff for a workpaper that
is generated once and then left alone (which is how these are used) -- it
would be the wrong choice for a template a human is expected to keep editing
by hand afterward.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from PIL import Image

from agent.ooxml.models import (
    SampleItem, EvidenceImage, NarrativeParagraph, ParametersTab,
)
from agent.ooxml.ooxml_utils import ColumnLayout, cell_box_emu
from agent.ooxml import imaging as ocr_utils

RED = "FF0000"
BLACK = "000000"

NS_DECL = ('xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"')


@dataclass
class DrawingResult:
    xml: str
    # list of (relationship_id, media_filename, PIL.Image) to write to xl/media
    # and reference from the matching drawingN.xml.rels
    media: list[tuple[str, str, Image.Image]] = field(default_factory=list)
    # (image name, letter, anchor_text, reason) for every callout whose anchor
    # could not be located on its image. See the try/except in
    # build_sample_drawing: these are dropped from the drawing rather than
    # killing the build, so they MUST be reported by whoever renders -- a
    # tickmark that is silently absent is exactly the failure this project
    # keeps designing against.
    unplaced_callouts: list[tuple[str, str, str, str]] = field(default_factory=list)

    def rels_xml(self) -> str:
        rel_items = "".join(
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
            f'officeDocument/2006/relationships/image" Target="../media/{fn}"/>'
            for rid, fn, _im in self.media
        )
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                f'2006/relationships">{rel_items}</Relationships>')


# ---------------------------------------------------------------------------
# Low-level shape templates
# ---------------------------------------------------------------------------

def _pic_xml(shape_id: int, name: str, rid: str, x: int, y: int, cx: int, cy: int) -> str:
    return f'''<xdr:absoluteAnchor>
<xdr:pos x="{x}" y="{y}"/>
<xdr:ext cx="{cx}" cy="{cy}"/>
<xdr:pic>
<xdr:nvPicPr>
<xdr:cNvPr id="{shape_id}" name="{name}"/>
<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
</xdr:nvPicPr>
<xdr:blipFill>
<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rid}"/>
<a:stretch><a:fillRect/></a:stretch>
</xdr:blipFill>
<xdr:spPr>
<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
</xdr:spPr>
</xdr:pic>
<xdr:clientData/>
</xdr:absoluteAnchor>'''


def _rect_label_xml(shape_id: int, name: str, x: int, y: int, cx: int, cy: int,
                     label: str, sz: int = 1000) -> str:
    return f'''<xdr:absoluteAnchor>
<xdr:pos x="{x}" y="{y}"/>
<xdr:ext cx="{cx}" cy="{cy}"/>
<xdr:sp macro="" textlink="">
<xdr:nvSpPr>
<xdr:cNvPr id="{shape_id}" name="{name}"/>
<xdr:cNvSpPr/>
</xdr:nvSpPr>
<xdr:spPr>
<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:noFill/>
<a:ln w="19050" cmpd="sng"><a:solidFill><a:srgbClr val="{RED}"/></a:solidFill></a:ln>
</xdr:spPr>
<xdr:style>
<a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>
<a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>
<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>
<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
</xdr:style>
<xdr:txBody>
<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" anchorCtr="1"><a:noAutofit/></a:bodyPr>
<a:lstStyle/>
<a:p><a:pPr algn="ctr"/><a:r>
<a:rPr lang="en-US" sz="{sz}" b="1"><a:solidFill><a:srgbClr val="{RED}"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr>
<a:t>{label}</a:t>
</a:r></a:p>
</xdr:txBody>
</xdr:sp>
<xdr:clientData/>
</xdr:absoluteAnchor>'''


def _narrative_xml(shape_id: int, name: str, x: int, y: int, cx: int, cy: int,
                    paragraphs: list[NarrativeParagraph], font_sz: int = 1000) -> str:
    body = ""
    for para in paragraphs:
        if not para.runs:
            body += f'<a:p><a:endParaRPr lang="en-US" sz="{font_sz}"/></a:p>'
            continue
        runs_xml = ""
        for run in para.runs:
            b = "1" if run.bold else "0"
            color = RED if run.color == "RED" else BLACK
            runs_xml += (f'<a:r><a:rPr lang="en-US" sz="{font_sz}" b="{b}">'
                         f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                         f'<a:latin typeface="Aptos"/></a:rPr>'
                         f'<a:t xml:space="preserve">{_xml_escape(run.text)}</a:t></a:r>')
        body += f"<a:p>{runs_xml}</a:p>"
    return f'''<xdr:absoluteAnchor>
<xdr:pos x="{x}" y="{y}"/>
<xdr:ext cx="{cx}" cy="{cy}"/>
<xdr:sp macro="" textlink="">
<xdr:nvSpPr>
<xdr:cNvPr id="{shape_id}" name="{name}"/>
<xdr:cNvSpPr txBox="1"/>
</xdr:nvSpPr>
<xdr:spPr>
<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:schemeClr val="accent4"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill>
<a:ln w="9525" cmpd="sng"><a:solidFill><a:srgbClr val="{RED}"/></a:solidFill></a:ln>
</xdr:spPr>
<xdr:style>
<a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>
<a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>
<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>
<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
</xdr:style>
<xdr:txBody>
<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="t"/>
<a:lstStyle/>
{body}
</xdr:txBody>
</xdr:sp>
<xdr:clientData/>
</xdr:absoluteAnchor>'''


def _xml_escape(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# ---------------------------------------------------------------------------
# High-level builder
# ---------------------------------------------------------------------------

def build_sample_drawing(
    sample_items: list[SampleItem],
    raw_data_column_layout: ColumnLayout,
    *,
    narrative_x: int = 0,
    narrative_cx: int = 5_545_854,
    narrative_row_height_px: int = 20,
    gap_emu: int = 300_000,
    image_gap_emu: int = 400_000,
    start_y_emu: int = 500_000,
) -> DrawingResult:
    """Builds the full drawing for the Sample tab: for every SampleItem, its
    raw-data-row callouts, narrative textbox, and evidence images (each with
    their own callouts), stacked vertically per sample and one sample after
    another.

    Assumes sample N's raw data lives in sheet row (N+1) (row 1 = header).
    """
    shapes: list[str] = []
    unplaced: list[tuple[str, str, str, str]] = []
    media: list[tuple[str, str, Image.Image]] = []
    shape_id = 2
    rid_counter = 1
    media_filename_counter = 1

    y_cursor = start_y_emu

    for sample_index, sample in enumerate(sample_items):
        row_num = sample_index + 2  # row 1 is header

        # --- raw data callouts (over the cell grid, not an image) ---
        for tie_out in sample.raw_data_tie_outs:
            col_start, col_end = tie_out.columns[0], tie_out.columns[-1]
            x, y, cx, cy = cell_box_emu(
                raw_data_column_layout, col_start, col_end, row_num,
                row_height_px=narrative_row_height_px,
            )
            shapes.append(_rect_label_xml(
                shape_id, f"Rectangle raw_{sample_index}_{tie_out.letter}",
                x, y, cx, cy, tie_out.letter,
            ))
            shape_id += 1

        # --- narrative textbox ---
        narrative_cy = _estimate_textbox_height(sample.narrative)
        shapes.append(_narrative_xml(
            shape_id, f"TextBox narrative_{sample_index}",
            narrative_x, y_cursor, narrative_cx, narrative_cy, sample.narrative,
        ))
        shape_id += 1
        block_bottom = y_cursor + narrative_cy

        # --- evidence images (placed to the right of the narrative for the
        # first image, then stacked full-width below) ---
        first = True
        image_x = narrative_x + narrative_cx + gap_emu
        image_y = y_cursor

        for img_spec in sample.evidence_images:
            im, dpi = ocr_utils.load_and_prepare(img_spec.source_path, img_spec.pdf_page)
            scale = img_spec.display_width_emu / im.width
            cx = img_spec.display_width_emu
            cy = int(im.height * scale)

            if not first:
                image_x = narrative_x
                image_y = block_bottom + image_gap_emu

            rid = f"rId{rid_counter}"
            fn = f"image{media_filename_counter}.png"
            rid_counter += 1
            media_filename_counter += 1
            media.append((rid, fn, im))

            shapes.append(_pic_xml(
                shape_id, f"Picture {img_spec.name}", rid, image_x, image_y, cx, cy,
            ))
            shape_id += 1

            # --- callouts on this image, located via OCR ---
            if img_spec.tie_outs:
                words = ocr_utils.ocr_words(im)
                for tie_out in img_spec.tie_outs:
                    try:
                        px, py, pw, ph = ocr_utils.find_anchor_box(
                            words, tie_out.anchor_text,
                            occurrence=tie_out.occurrence,
                            extra_words=tie_out.extra_words,
                        )
                    except ValueError as exc:
                        # Originally this propagated, reasoning that a missing
                        # anchor should fail loudly rather than box the wrong
                        # value. That is right for a human building one file by
                        # hand, and wrong here: the anchor text is chosen by a
                        # model from a value it read, OCR may render that value
                        # slightly differently, and losing an entire workpaper
                        # over one unlocatable box is far worse than losing the
                        # box. So drop the callout, keep the image, and record
                        # why -- unplaced, never unreported.
                        unplaced.append(
                            (img_spec.name, tie_out.letter, tie_out.anchor_text, str(exc))
                        )
                        continue
                    bx = image_x + int(px * scale)
                    by = image_y + int(py * scale)
                    bw = int(pw * scale)
                    bh = int(ph * scale)
                    shapes.append(_rect_label_xml(
                        shape_id,
                        f"Rectangle {img_spec.name}_{tie_out.letter}",
                        bx, by, bw, bh, tie_out.letter,
                    ))
                    shape_id += 1

            block_bottom = max(block_bottom, image_y + cy)
            first = False

        y_cursor = block_bottom + image_gap_emu

    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           f'<xdr:wsDr {NS_DECL}>' + "".join(shapes) + "</xdr:wsDr>")
    return DrawingResult(xml=xml, media=media, unplaced_callouts=unplaced)


def build_parameters_drawing(params: ParametersTab) -> DrawingResult:
    im = Image.open(params.screenshot_path)
    native_dpi = im.info.get("dpi", (96, 96))[0] or 96
    scale_to_100pct = 914_400 / native_dpi

    if params.display_width_emu:
        cx = params.display_width_emu
        scale = cx / im.width
    else:
        scale = scale_to_100pct
        cx = int(im.width * scale)
    cy = int(im.height * scale)

    shapes = []
    media = []
    shape_id = 2
    media.append(("rId1", "image1.png", im))
    shapes.append(_pic_xml(shape_id, "Picture params", "rId1", 0, 0, cx, cy))
    shape_id += 1

    narrative_y = cy + 250_000
    narrative_cy = _estimate_textbox_height(params.narrative, sz=1100)
    shapes.append(_narrative_xml(
        shape_id, "TextBox params_narrative", 0, narrative_y,
        6_057_900, narrative_cy, params.narrative, font_sz=1100,
    ))
    shape_id += 1

    if params.callouts:
        words = ocr_utils.ocr_words(im)
        for c in params.callouts:
            px, py, pw, ph = ocr_utils.find_anchor_box(
                words, c.anchor_text, occurrence=c.occurrence,
                extra_words=c.extra_words,
            )
            bx, by, bw, bh = (int(px * scale), int(py * scale),
                               int(pw * scale), int(ph * scale))
            shapes.append(_rect_label_xml(
                shape_id, f"Rectangle params_{c.anchor_text[:10]}",
                bx, by, bw, bh, "",  # parameters callouts have no letter in
                                      # the original template
            ))
            shape_id += 1

    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           f'<xdr:wsDr {NS_DECL}>' + "".join(shapes) + "</xdr:wsDr>")
    return DrawingResult(xml=xml, media=media)


def _estimate_textbox_height(paragraphs: list[NarrativeParagraph], *,
                              sz: int = 1000, chars_per_line: int = 95,
                              line_height_emu: int = 165_000) -> int:
    """Rough line-count estimate so the narrative box is tall enough not to
    clip text. Excel textboxes don't auto-grow reliably when set via raw XML
    (a:spAutoFit is unreliable across LibreOffice/Excel), so this errs on
    the generous side. Tune chars_per_line to your box width / font size if
    text is being clipped or the box looks oversized."""
    lines = 0
    for p in paragraphs:
        text_len = sum(len(r.text) for r in p.runs)
        lines += max(1, -(-text_len // chars_per_line))  # ceil div
    return max(line_height_emu * lines, 500_000)
