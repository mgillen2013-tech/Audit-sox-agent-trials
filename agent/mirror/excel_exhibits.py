"""Excel support -> pictures the workpaper can show.

The workpaper shows support as pictures in the PY's picture slots. PDFs
and images have pages to picture; an Excel file did not, so evidence the
testing agent read from a workbook never reached the workpaper, and the
PY's pictures were left in its place. This makes two kinds of picture:

- a screenshot pasted on a worksheet, exactly as pasted;
- the cells the agent cited, drawn as a plain table (values as the
  workbook last calculated them, bold and fills kept, the header row
  included), the way a preparer would screenshot the range.

Pictures are written to an _exhibits folder inside the support folder and
are then handled like any other image support.
"""

from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path

EXCEL_SUFFIXES = (".xlsx", ".xlsm")
EXHIBIT_DIR = "_exhibits"
_MAX_ROWS, _MAX_COLS = 60, 20
_LOC_RE = re.compile(r"^(?P<sheet>.+?)!(?:image(?P<img>\d+)|(?P<a>[A-Z]+\d+)(?::(?P<b>[A-Z]+\d+))?)$")


def parse_location(location: str) -> tuple[str, int | None, tuple[int, int, int, int] | None] | None:
    """'Sheet1!B3:D9' -> ('Sheet1', None, (min_row, min_col, max_row, max_col));
    'Sheet1!image2' -> ('Sheet1', 2, None); anything else -> None."""
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

    m = _LOC_RE.match((location or "").strip())
    if not m:
        return None
    if m.group("img"):
        return m.group("sheet"), int(m.group("img")), None
    cells = [coordinate_from_string(m.group(k)) for k in ("a", "b") if m.group(k)]
    rows = [r for _c, r in cells]
    cols = [column_index_from_string(c) for c, _r in cells]
    return m.group("sheet"), None, (min(rows), min(cols), max(rows), max(cols))


def _font(size: int, bold: bool = False):
    from PIL import ImageFont

    for name in (("arialbd.ttf", "DejaVuSans-Bold.ttf") if bold else ("arial.ttf", "DejaVuSans.ttf")):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default(size=size)


def _text(value) -> str:
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%m/%d/%Y")
    if isinstance(value, float):
        return f"{value:,.2f}" if abs(value) >= 1000 or value != int(value) else str(int(value))
    return str(value)


def render_range(ws, min_row: int, min_col: int, max_row: int, max_col: int):
    """The cells as a table image, plus their text (for matching)."""
    from openpyxl.utils import get_column_letter
    from PIL import Image, ImageDraw

    max_row = min(max_row, min_row + _MAX_ROWS - 1)
    max_col = min(max_col, min_col + _MAX_COLS - 1)
    regular, bold = _font(14), _font(14, bold=True)
    pad, row_h = 6, 24
    measure = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    widths = []
    for c in range(min_col, max_col + 1):
        # Wide enough for the widest value as actually drawn (bold where
        # bold), never narrower than the workbook's own column width.
        drawn = max((measure.textlength(_text(ws.cell(r, c).value),
                                        font=bold if getattr(ws.cell(r, c).font, "b", False) else regular)
                     for r in range(min_row, max_row + 1)), default=0)
        declared = (ws.column_dimensions[get_column_letter(c)].width or 0) * 7
        widths.append(int(min(max(drawn, declared, 32), 420)) + 2 + 2 * pad)  # +2: never round a value off
    image = Image.new("RGB", (sum(widths) + 1, (max_row - min_row + 1) * row_h + 1), "white")
    draw = ImageDraw.Draw(image)
    lines = []
    y = 0
    for r in range(min_row, max_row + 1):
        x, row_text = 0, []
        for i, c in enumerate(range(min_col, max_col + 1)):
            cell = ws.cell(r, c)
            fill = getattr(cell.fill, "fgColor", None)
            rgb = getattr(fill, "rgb", None) if getattr(cell.fill, "fill_type", None) == "solid" else None
            if isinstance(rgb, str) and len(rgb) == 8 and rgb not in ("00000000", "FFFFFFFF"):
                draw.rectangle([x, y, x + widths[i], y + row_h], fill="#" + rgb[2:])
            draw.rectangle([x, y, x + widths[i], y + row_h], outline="#b7b7b7")
            text = _text(cell.value)
            if text:
                font = bold if getattr(cell.font, "b", False) else regular
                limit = widths[i] - 2 * pad
                while text and draw.textlength(text, font=font) > limit:
                    text = text[:-1]
                draw.text((x + pad, y + 4), text, fill="black", font=font)
                row_text.append(_text(cell.value))
            x += widths[i]
        lines.append(" | ".join(row_text))
        y += row_h
    return image, "\n".join(lines)


def _clean(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|\r\n]+', "_", text).strip()


def make_exhibits(path: Path, locations: list[str], out_dir: Path) -> list[tuple[str, str, list[str]]]:
    """Pictures for the cited locations of one workbook.

    Returns (picture file name relative to the support folder, its text,
    the cited locations it shows). Cited cells on one sheet become ONE
    picture covering all of them, from the sheet's header row down; each
    cited pasted screenshot becomes its own picture.
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    folder = out_dir / EXHIBIT_DIR
    folder.mkdir(parents=True, exist_ok=True)
    out: list[tuple[str, str, list[str]]] = []
    ranges: dict[str, list[tuple[str, tuple[int, int, int, int]]]] = {}
    for loc in dict.fromkeys(locations):
        # The agent may prefix the file name, as in "Recon.xlsx!Sheet1!B3".
        bare = loc.split("!", 1)[1] if loc.lower().startswith(path.name.lower() + "!") else loc
        parsed = parse_location(bare)
        if parsed is None or parsed[0] not in wb.sheetnames:
            continue
        sheet, img, box = parsed
        ws = wb[sheet]
        if img is not None:
            images = getattr(ws, "_images", [])
            if not 1 <= img <= len(images):
                continue
            from PIL import Image

            name = f"{EXHIBIT_DIR}/{_clean(path.stem)} - {_clean(sheet)} - picture {img}.png"
            Image.open(BytesIO(images[img - 1]._data())).convert("RGB").save(out_dir / name)
            out.append((name, "", [loc]))
        else:
            ranges.setdefault(sheet, []).append((loc, box))
    for sheet, cited in ranges.items():
        ws = wb[sheet]
        min_row = 1 if min(b[0] for _l, b in cited) - 1 <= _MAX_ROWS // 2 else min(b[0] for _l, b in cited)
        min_col = min(b[1] for _l, b in cited)
        max_row = max(b[2] for _l, b in cited)
        max_col = max(b[3] for _l, b in cited)
        # A single cited cell is shown with its whole row, for context.
        if min_col == max_col:
            min_col, max_col = 1, max(max_col, min(ws.max_column, _MAX_COLS))
        image, text = render_range(ws, min_row, min_col, max_row, max_col)
        name = f"{EXHIBIT_DIR}/{_clean(path.stem)} - {_clean(sheet)} rows {min_row}-{max_row}.png"
        image.save(out_dir / name)
        out.append((name, text, [loc for loc, _b in cited]))
    if not out and locations:
        # Cited, but at no location that can be read back: show the first
        # sheet rather than nothing.
        ws = wb.worksheets[0]
        image, text = render_range(ws, 1, 1, ws.max_row, ws.max_column)
        name = f"{EXHIBIT_DIR}/{_clean(path.stem)} - {_clean(ws.title)} rows 1-{min(ws.max_row, _MAX_ROWS)}.png"
        image.save(out_dir / name)
        out.append((name, text, list(locations)))
    return out
