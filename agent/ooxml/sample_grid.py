"""Where the sampled item's row goes on a PY sample tab, and under which headings.

The builder used to assume two things about every PY sample tab: headings
in row 1, the sampled item in row 2 -- and the item's values written into
columns A, B, C... in whatever order THIS year's listing happened to have
its columns. Both held for the first PY it was built against. Neither is a
property of PY workpapers in general:

- a tab with a title or a note above its headings had the title treated as
  headings and everything from row 2 down -- headings included -- replaced;
- an extract whose columns came out in a different order this year put
  every value under the wrong heading, which is worse than a blank because
  it looks right.

So both are now READ from the tab. The heading row is the row whose cells
best match the listing's column names; the item goes directly beneath it;
each value goes under the heading with its name. Nothing above the heading
row is touched.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from agent.ooxml import ooxml_utils as ox

# How far down to look for the heading row. PY tabs put a title or a short
# note above the table, not pages of it.
_SCAN_ROWS = 40


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def _match(listing: str, heading: str) -> int:
    """2 = same name, 1 = one contains the other, 0 = unrelated.

    Containment covers the common real case: an E1 extract heading like
    "Invoice Number F0411.VINV" against a PY heading of "Invoice Number".
    Short names are exact-only -- "No" is contained in far too much.
    """
    a, b = _norm(listing), _norm(heading)
    if not a or not b:
        return 0
    if a == b:
        return 2
    if min(len(a), len(b)) >= 5 and (a in b or b in a):
        return 1
    return 0


@dataclass
class SampleGrid:
    header_row: int
    data_row: int
    # Listing column letter (A, B, C... in listing order) -> template letter.
    column_for: dict[str, str] = field(default_factory=dict)
    # Template letter -> heading text to WRITE (columns the PY lacked, or
    # every column when nothing matched).
    headings_to_write: dict[str, str] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)


def plan_grid(template_cells: dict[str, tuple], listing_headers: dict[str, str]) -> SampleGrid:
    """Decide the heading row, the data row and the column for each value.

    template_cells: {ref: (style, text, has_formula)} from the PY tab.
    listing_headers: {listing letter: column name} in listing order.
    """
    rows: dict[int, dict[str, str]] = {}
    for ref, (_style, text, _f) in template_cells.items():
        m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
        if m and int(m.group(2)) <= _SCAN_ROWS and str(text).strip():
            rows.setdefault(int(m.group(2)), {})[m.group(1)] = str(text)

    names = [n for n in listing_headers.values() if n]
    best_row, best_score, best_map = None, 0, {}
    for r in sorted(rows):
        used: set[str] = set()
        mapping: dict[str, str] = {}
        score = 0
        # Exact names first, so "Invoice Number" never loses its column to a
        # looser containment match from another field.
        for strength in (2, 1):
            for letter, name in listing_headers.items():
                if letter in mapping or not name:
                    continue
                for col, heading in sorted(rows[r].items(), key=lambda kv: ox.col_letter_to_index(kv[0])):
                    if col not in used and _match(name, heading) == strength:
                        mapping[letter], score = col, score + strength
                        used.add(col)
                        break
        if score > best_score:
            best_row, best_score, best_map = r, score, mapping

    needed = 2 if len(names) > 1 else 1
    if best_row is None or len(best_map) < needed:
        # Nothing on the tab reads as this listing's headings. Keep the
        # long-standing layout (headings row 1, item row 2) but WRITE the
        # listing's names into the heading row: values under someone else's
        # labels look right and are wrong.
        letters = {letter: letter for letter in listing_headers}
        return SampleGrid(
            header_row=1, data_row=2, column_for=letters,
            headings_to_write={letter: name for letter, name in listing_headers.items() if name},
            notes=["the prior-year sample tab's headings don't match this year's listing, so "
                   "the listing's own column names were written as the headings"],
        )

    # A listing column with no matching heading. If the PY has an UNUSED
    # column between where its listing neighbours landed, it goes there --
    # that is where the PY had it, under whatever the PY called it (a vendor
    # name under a second "Payee Number" heading, in the real one). Only
    # strictly between both neighbours, so it can never be pushed under an
    # unrelated heading such as an empty "Void Date". Otherwise it goes
    # after the PY's last column, with its own heading.
    order = list(listing_headers)
    column_for = dict(best_map)
    used = set(column_for.values())
    last = max(ox.col_letter_to_index(c) for c in rows[best_row])
    headings: dict[str, str] = {}
    for pos, letter in enumerate(order):
        if letter in column_for:
            continue
        left = next((column_for[o] for o in reversed(order[:pos]) if o in best_map), None)
        right = next((column_for[o] for o in order[pos + 1:] if o in best_map), None)
        lo = ox.col_letter_to_index(left) if left else 0
        hi = ox.col_letter_to_index(right) if right else last + 1
        between = [ox.col_index_to_letter(i) for i in range(lo + 1, hi)
                   if ox.col_index_to_letter(i) not in used]
        if left and right and between:
            column_for[letter] = between[0]
        else:
            last += 1
            column_for[letter] = ox.col_index_to_letter(last)
            if listing_headers[letter]:
                headings[column_for[letter]] = listing_headers[letter]
        used.add(column_for[letter])
    notes = []
    if best_row != 1:
        notes.append(f"the prior-year sample tab's headings are on row {best_row}; "
                     f"rows above them were kept")
    if headings:
        notes.append("columns not on the prior-year sample tab were added after its last "
                     f"column: {', '.join(headings.values())}")
    return SampleGrid(best_row, best_row + 1, column_for, headings, notes)


def write_headings(sheet_xml_path: str, row: int, headings: dict[str, str],
                   sst: ox.SharedStrings, style: int | None) -> None:
    """Set heading cells on an existing row, keeping the row's other cells
    and Excel's required left-to-right cell order."""
    if not headings:
        return
    with open(sheet_xml_path, encoding="utf-8") as fh:
        xml = fh.read()
    xml = re.sub(r"<sheetData\s*/>", "<sheetData></sheetData>", xml, count=1)
    row_re = re.compile(rf'<row\b[^>]*?\br="{row}"[^>]*?(?:/>|>(.*?)</row>)', re.S)
    m = row_re.search(xml)
    cells: dict[str, str] = {}
    if m and m.group(1):
        for c in re.finditer(r'<c\b[^>]*?\br="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', m.group(1), re.S):
            cells[c.group(1)] = c.group(0)
    for col, text in headings.items():
        cells[col] = ox.cell_xml(f"{col}{row}", text, style=style, sst=sst)
    ordered = "".join(cells[c] for c in sorted(cells, key=ox.col_letter_to_index))
    new_row = f'<row r="{row}">{ordered}</row>'
    if m:
        open_tag = re.match(r"<row\b[^>]*?(?=/?>)", m.group(0)).group(0)
        open_tag = re.sub(r'\sspans="[^"]*"', "", open_tag)
        new_row = f"{open_tag}>{ordered}</row>"
        xml = xml[: m.start()] + new_row + xml[m.end():]
    else:
        later = [x for x in re.finditer(r'<row\b[^>]*?\br="(\d+)"', xml) if int(x.group(1)) > row]
        at = later[0].start() if later else xml.index("</sheetData>")
        xml = xml[:at] + new_row + xml[at:]
    with open(sheet_xml_path, "w", encoding="utf-8") as fh:
        fh.write(xml)
