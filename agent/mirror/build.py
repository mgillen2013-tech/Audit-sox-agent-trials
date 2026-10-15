"""Build a CY workpaper by editing a copy of the PY. See agent/mirror/__init__.py.

Order of work:
  1. read the PY's organisation (structure.py -- the model, checked by code)
  2. lay out the CY tabs the same way: each CY selection gets the same group
     of tabs a PY selection had, named in the PY's convention
  3. write each selection's record under the PY's headings
  4. per tab, the model decides every element (rewrite.py); code applies it
  5. population written back if supplied; PY evidence tabs flagged
  6. checks: structure, leftovers across the whole workbook, placement,
     support; then a model pass -- on the Mirror Check and Support Check tabs
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from agent.mirror import verify
from agent.mirror.apply import (
    ApplyReport, apply_tab, recalc_on_open, set_cells,
)
from agent.mirror.facts import selection_facts
from agent.mirror.inventory import Inventory, TabInventory, build_inventory
from agent.mirror.rewrite import RELOCATE_SYSTEM, Exhibit, SelectionFacts, call_model, tab_prompt, validate
from agent.mirror.structure import Structure, read_structure
from agent.ooxml import imaging
from agent.ooxml import ooxml_utils as ox
from agent.ooxml.layout import Slot, _Grid, _row_heights
from agent.ooxml.sample_grid import _match, plan_grid
from agent.ooxml.summary_sheet import DRAFT_BANNER, GENERATED_HEADINGS, add_summary_sheet, add_table_sheet

MIRROR_CHECK_COLUMNS = [("A", "Severity", 10), ("B", "Tab", 16), ("C", "Item", 18),
                        ("D", "Issue", 90), ("E", "Found by", 14)]
SUPPORT_CHECK_COLUMNS = [("A", "Match", 10), ("B", "Tab", 14), ("C", "PY document", 50),
                         ("D", "CY counterpart", 40), ("E", "Comment", 60)]
_CHECK_HEADINGS = tuple(t for _c, t, _w in MIRROR_CHECK_COLUMNS[:4])
_SUPPORT_HEADINGS = tuple(t for _c, t, _w in SUPPORT_CHECK_COLUMNS[:4])


@dataclass
class MirrorOutcome:
    path: Path
    flags: list[verify.CheckFlag] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    tab_names: dict[str, str] = field(default_factory=dict)  # sample_id -> first CY tab
    support: list[dict] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------


def _remove_generated_tabs(work_dir: str) -> None:
    """A previous run's Summary / Mirror Check / Support Check, by headings."""
    for name, num in list(ox.resolve_sheets(work_dir).items()):
        rows = ox.first_rows_text(work_dir, num, rows=2)
        if any(tuple(r[:3]) == GENERATED_HEADINGS or tuple(r[:4]) in (_CHECK_HEADINGS, _SUPPORT_HEADINGS)
               for r in rows):
            ox.delete_sheet(work_dir, name)


def _derived_names(group: list[str], py_index: int, cy_index: int, taken: set[str]) -> list[str]:
    """Names for a CY selection's tabs when the model gave none: the PY
    names with the PY selection's number swapped for the CY one where the
    number appears, else the PY name with the CY number appended."""
    out = []
    for name in group:
        new = re.sub(rf"(?<!\d){py_index + 1}(?!\d)", str(cy_index + 1), name, count=1)
        if new == name and cy_index != py_index:
            new = f"{name} ({cy_index + 1})"
        base, k = new[:31], 2
        while base.lower() in taken:
            base, k = f"{new[:28]} {k}", k + 1
        taken.add(base.lower())
        out.append(base)
    return out


def _arrange(work_dir: str, inv: Inventory, structure: Structure, n: int, notes: list[str]) -> list[list[tuple[str, str]]]:
    """Lay out CY selection tabs the way the PY laid out its selections.

    Returns, per CY selection, [(CY tab name, PY tab it was made from)].
    CY selection i is made from PY selection i's tabs, or the first PY
    selection's for selections beyond the PY's count. PY selection groups
    beyond the CY count are last year's items and are removed.
    """
    groups = structure.groups
    if not groups:
        return [[] for _ in range(n)]
    taken = {t.name.lower() for t in inv.tabs if not any(t.name in g for g in groups)}
    names: list[list[str]] = []
    for i in range(n):
        template = groups[min(i, len(groups) - 1)]
        want = structure.cy_names[i] if i < len(structure.cy_names) and structure.cy_names[i] else []
        if not want or any(w.lower() in taken for w in want):
            want = _derived_names(template, min(i, len(groups) - 1), i, taken)
            if structure.from_model:
                notes.append(f"CY tab names for selection {i + 1} were derived from the PY names: {want}")
        else:
            taken |= {w.lower() for w in want}
        names.append(want)

    # Copies first, from the pristine PY tabs, under temporary names.
    sheets = ox.resolve_sheets(work_dir)
    after = groups[-1][-1]
    temp: list[list[str]] = []
    for i in range(n):
        if i < len(groups):
            temp.append(list(groups[i]))
            continue
        made = []
        for j, py_tab in enumerate(groups[0]):
            t = f"__cy_{i}_{j}"
            ox.clone_sheet(work_dir, sheets[py_tab], t, after=after, keep_drawing=True)
            made.append(t)
            after = t
        temp.append(made)
    for g in groups[n:]:
        for t in g:
            ox.delete_sheet(work_dir, t)
    # Two-step rename so "2" -> "1" style swaps can never collide.
    sheets = ox.resolve_sheets(work_dir)
    k = 0
    for g in temp:
        for t in g:
            ox.rename_sheet(work_dir, sheets[t], f"__rename_{k}")
            k += 1
    sheets = ox.resolve_sheets(work_dir)
    k = 0
    out = []
    for i, g in enumerate(names):
        template = groups[min(i, len(groups) - 1)]
        pairs = []
        for j, final in enumerate(g):
            ox.rename_sheet(work_dir, sheets[f"__rename_{k}"], final)
            pairs.append((final, template[j]))
            k += 1
        out.append(pairs)
    return out


# ---------------------------------------------------------------------------
# What the PY elements sit on / showed (read from the PY itself)
# ---------------------------------------------------------------------------

_OCR_CACHE: dict[str, tuple[Any, list]] = {}


def _ocr(path: str):
    if path not in _OCR_CACHE:
        try:
            from PIL import Image

            im = Image.open(path).convert("RGB")
            _OCR_CACHE[path] = (im, imaging.ocr_words(im))
        except Exception:  # noqa: BLE001 -- emf/wmf and the like: no text, still a slot
            _OCR_CACHE[path] = (None, [])
    return _OCR_CACHE[path]


def _grid(work_dir: str, tab: TabInventory) -> _Grid:
    path = os.path.join(work_dir, "xl", "worksheets", f"sheet{tab.sheet_num}.xml")
    with open(path, encoding="utf-8") as fh:
        default_h, custom_h = _row_heights(fh.read())
    return _Grid(ox.parse_column_widths(path), default_h, custom_h)


def _py_reading(tab: TabInventory, grid: _Grid) -> tuple[dict[str, str], dict[str, str], str]:
    """(what each PY picture showed, what each other element sits on, all
    PY evidence text on this tab)."""
    pictures = tab.by_kind("picture")
    pic_notes: dict[str, str] = {}
    evidence: list[str] = []
    for p in pictures:
        _im, words = _ocr(p.image_path) if p.image_path else (None, [])
        text = " ".join(w.text for w in sorted(words, key=lambda w: (w.top, w.left)))
        pic_notes[p.id] = text[:300] or "(no readable text)"
        evidence.append(text)
    marks: dict[str, str] = {}
    for item in tab.items:
        if item.kind in ("picture", "frame", "group", "object"):
            continue
        host = next((p for p in pictures if p.box.x <= item.box.x + item.box.cx / 2 <= p.box.right
                     and p.box.y <= item.box.y + item.box.cy / 2 <= p.box.bottom), None)
        if host is not None:
            im, words = _ocr(host.image_path) if host.image_path else (None, [])
            if im is None:
                continue
            sx, sy = im.width / max(1, host.box.cx), im.height / max(1, host.box.cy)
            x0, y0 = (item.box.x - host.box.x) * sx, (item.box.y - host.box.y) * sy
            x1, y1 = (item.box.right - host.box.x) * sx, (item.box.bottom - host.box.y) * sy
            inside = [w.text for w in words if x0 <= w.left + w.width / 2 <= x1 and y0 <= w.top + w.height / 2 <= y1]
            if inside:
                marks[item.id] = f"{' '.join(inside)[:80]!r} on PY picture {host.id}"
            continue
        # On the cell grid: the cells under it.
        c0, c1 = grid.cell_at(item.box.x, item.box.y), grid.cell_at(item.box.right - 1, item.box.bottom - 1)
        m0, m1 = re.fullmatch(r"([A-Z]+)(\d+)", c0), re.fullmatch(r"([A-Z]+)(\d+)", c1)
        under = []
        for r in range(int(m0.group(2)), min(int(m1.group(2)), int(m0.group(2)) + 3) + 1):
            for ci in range(ox.col_letter_to_index(m0.group(1)), min(ox.col_letter_to_index(m1.group(1)), ox.col_letter_to_index(m0.group(1)) + 6) + 1):
                cell = tab.cells.get(f"{ox.col_index_to_letter(ci)}{r}")
                if cell is not None and cell.text.strip():
                    under.append(f"{cell.ref}={cell.text.strip()[:30]}")
        if under and item.kind != "text":
            marks[item.id] = "cells " + ", ".join(under[:6])
    return pic_notes, marks, " ".join(evidence)


# ---------------------------------------------------------------------------
# Data rows
# ---------------------------------------------------------------------------


def _cells_for_grid(tab: TabInventory, only_row: int | None = None) -> dict[str, tuple]:
    return {ref: (c.style, c.text, c.formula) for ref, c in tab.cells.items()
            if only_row is None or re.fullmatch(rf"[A-Z]+{only_row}", ref)}


def _listing(facts: SelectionFacts) -> dict[str, str]:
    return {ox.col_index_to_letter(i + 1): k for i, k in enumerate(facts.data_row)}


def _write_record(work: str, tab: TabInventory, f: SelectionFacts, rows: tuple[int, list[int]] | None,
                  sst, flags: list, notes: list) -> tuple[int, int, dict[str, str]]:
    """The selection's record under the PY's headings. Returns (header row,
    data row, {column: heading})."""
    listing = _listing(f)
    if not listing:
        return (rows[0], rows[1][0] if rows and rows[1] else rows[0] + 1, {}) if rows else (1, 2, {})
    header_row = rows[0] if rows and rows[0] else None
    plan = plan_grid(_cells_for_grid(tab, header_row), listing) if header_row else plan_grid(_cells_for_grid(tab), listing)
    if header_row and plan.header_row != header_row:
        plan.header_row, plan.data_row = header_row, header_row + 1
    data_rows = rows[1] if rows and rows[1] else [plan.data_row]
    header = {re.sub(r"\d+", "", ref): c.text for ref, c in tab.cells.items()
              if re.fullmatch(rf"[A-Z]+{plan.header_row}", ref)}
    if plan.headings_to_write and len(plan.headings_to_write) == len(listing):
        flags.append(verify.CheckFlag(
            "high", tab.name, "data row",
            "this year's sample listing has none of the PY tab's column headings, so the selection's "
            "record was not written -- check the listing is the same report as last year", "structure"))
        return plan.header_row, data_rows[0], header
    sheet_path = os.path.join(work, "xl", "worksheets", f"sheet{tab.sheet_num}.xml")
    values: dict[str, Any] = {}
    for r in data_rows:  # last year's record(s) out
        for ref in [x for x in tab.cells if re.fullmatch(rf"[A-Z]+{r}", x)]:
            values[ref] = None
    for letter, heading in listing.items():
        values[f"{plan.column_for[letter]}{data_rows[0]}"] = f.data_row.get(heading)
    for col, heading in plan.headings_to_write.items():
        values[f"{col}{plan.header_row}"] = heading
    set_cells(sheet_path, values, sst, {})
    if len(data_rows) > 1:
        notes.append(f"[{tab.name}] the PY selection had {len(data_rows)} record rows; CY's one record is on "
                     f"row {data_rows[0]} and the others were cleared")
    notes += [f"[{tab.name}] {n}" for n in plan.notes]
    return plan.header_row, data_rows[0], header


def _write_rows(work: str, tab: TabInventory, header_row: int, first: int, last: int,
                records: list[dict], sst, notes: list) -> bool:
    """Replace a block of PY rows (a selection table, a population) with CY
    records under the PY's headings. Rows outside the block are untouched."""
    if not records:
        return False
    listing = {ox.col_index_to_letter(i + 1): k for i, k in enumerate(records[0])}
    plan = plan_grid(_cells_for_grid(tab, header_row) if header_row else _cells_for_grid(tab), listing)
    if plan.headings_to_write and len(plan.headings_to_write) == len(listing):
        notes.append(f"[{tab.name}] none of this year's columns match the PY headings; left for review")
        return False
    header_row = header_row or plan.header_row
    if not first:
        used = sorted({int(re.sub(r"[A-Z]+", "", r)) for r in tab.cells if int(re.sub(r"[A-Z]+", "", r)) > header_row})
        first = header_row + 1
        last = first - 1
        for r in used:  # the contiguous block under the headings
            if r == last + 1:
                last = r
            else:
                break
    sheet_path = os.path.join(work, "xl", "worksheets", f"sheet{tab.sheet_num}.xml")
    values: dict[str, Any] = {}
    for r in range(first, last + 1):
        for ref in [x for x in tab.cells if re.fullmatch(rf"[A-Z]+{r}", x)]:
            values[ref] = None
    for i, rec in enumerate(records):
        for letter, heading in listing.items():
            col = plan.column_for.get(letter)
            if col:
                values[f"{col}{first + i}"] = rec.get(heading)
    set_cells(sheet_path, values, sst, {})
    if first + len(records) - 1 > last:
        notes.append(f"[{tab.name}] CY has more rows than the PY block ({len(records)} vs {last - first + 1}); "
                     "they continue below it")
    return True


# ---------------------------------------------------------------------------


def _facts_lines(f: SelectionFacts | None, all_facts: list[SelectionFacts], spec: dict) -> list[str]:
    lines = [f"Control: {spec.get('control_id', '')} -- {spec.get('control_objective_text', '')}"]
    for s in spec.get("test_steps", []):
        lines.append(f"Test step {s.get('test_step_id')}: {s.get('test_step_text', '')}")
    for g in ([f] if f is not None else all_facts):
        lines.append(f"Selection {g.sample_id} record: " + json.dumps(g.data_row, default=str))
        lines += [json.dumps(s, default=str) for s in g.steps]
    return lines


def _ensure_png(work_dir: str) -> None:
    ct_path = os.path.join(work_dir, "[Content_Types].xml")
    with open(ct_path, encoding="utf-8") as fh:
        ct = fh.read()
    if not re.search(r'<Default\b[^>]*Extension="png"', ct, re.I):
        ct = ct.replace("<Override", '<Default Extension="png" ContentType="image/png"/><Override', 1) \
            if "<Override" in ct else ct.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>', 1)
        with open(ct_path, "w", encoding="utf-8") as fh:
            fh.write(ct)


def _cy_text_locations(work: str, tabs: set[str]) -> list[tuple[str, str, str, bool]]:
    """Every cell and shape text on the given CY tabs."""
    inv = build_inventory(work)
    out = []
    for t in inv.tabs:
        if t.name not in tabs:
            continue
        for ref, c in t.cells.items():
            out.append((t.name, ref, c.text, False))
        for i in t.items:
            if i.text:
                out.append((t.name, f"{i.kind} {i.name}", i.text, False))
    return out


def kind_is_label(text: str, labels: set[str]) -> bool:
    """A text that IS a tickmark label (the mark itself), not text citing one."""
    return re.sub(r"\s+", " ", (text or "").strip()) in labels


def _collect_tickmarks(tab: TabInventory, plan, report: ApplyReport, exhibits: list[Exhibit],
                       declared: list, placed_labels: dict, tick_lines: dict) -> None:
    """Where this tab's tickmark labels ended up, and a line per declared
    tickmark for the final review (label -> value placed on -> citing text)."""
    pictures = tab.by_kind("picture")
    files = {x.id: f"{x.source_file} p.{x.page + 1}" for x in exhibits}
    for item in tab.items:
        if item.kind != "text":
            continue
        d = plan.elements.get(item.id, {})
        text = (re.sub(r"</?s\d+>", "", " ".join(d["paragraphs"])) if d.get("paragraphs") else item.text).strip()
        if not text or len(text) > 12:
            continue
        status = report.status.get(item.id, "kept")
        on_picture = any(p.box.x <= item.box.x + item.box.cx / 2 <= p.box.right
                         and p.box.y <= item.box.y + item.box.cy / 2 <= p.box.bottom for p in pictures)
        if status == "placed":
            on = report.placed_on.get(item.id, {})
            where = (f"{on.get('anchor')!r} on {files.get(on.get('exhibit'), on.get('exhibit'))}" if "exhibit" in on
                     else f"data row '{on.get('heading')}'" if "heading" in on else "beside another mark")
        elif status == "kept" and not on_picture:
            where = "kept at its PY position on the cell grid"
        else:
            continue
        placed_labels.setdefault(re.sub(r"\s+", " ", text), []).append((tab.name, where))
    texts = {i.id: i.text for i in tab.items}
    for t in plan.tickmarks:
        declared.append((tab.name, t))
        on = report.placed_on.get(t["element"], {})
        where = (f"{on.get('anchor')!r} on {files.get(on.get('exhibit'), on.get('exhibit'))}" if "exhibit" in on
                 else report.status.get(t["element"], "not placed"))
        citing = []
        for r in t.get("referenced_by", []):
            d = plan.elements.get(r, {})
            citing.append(re.sub(r"</?s\d+>", "", " ".join(d["paragraphs"]))[:200] if d.get("paragraphs")
                          else (texts.get(r) or (plan.cells.get(r, {}).get("value") or r))[:200])
        tick_lines.setdefault(tab.name, []).append(
            f"{t['label']}: placed on {where}; said to mark {t.get('marks', '')!r}; cited by {citing}")


def mirror_workpaper(
    spec: dict[str, Any],
    results: dict[str, dict],
    py_path: Path,
    output_path: Path,
    support_dir: Path,
    manifests: dict | None,
    client: Any,
    model: str,
    *,
    summary_rows_fn=None,
    ledger: Any = None,
    run_model_check: bool = True,
    population_records: list[dict] | None = None,
) -> MirrorOutcome:
    work = str(output_path) + ".mirror"
    ox.extract_template(str(py_path), work)
    # A pristine read of the PY, for what it cited (tickmark labels) --
    # the working copy is edited in place.
    ox.extract_template(str(py_path), str(output_path) + ".pyread")
    _remove_generated_tabs(work)
    inv = build_inventory(work)
    facts = selection_facts(results, spec, manifests, support_dir)
    sst = ox.SharedStrings(work)
    flags: list[verify.CheckFlag] = []
    notes: list[str] = []
    support_rows: list[dict] = []

    # --- 1. how the PY is organised ---
    structure, problems = read_structure(client, model, inv, [(f.sample_id, f.data_row) for f in facts], ledger=ledger)
    notes += [f"structure: {p}" for p in problems] + structure.notes
    if not structure.from_model:
        flags.append(verify.CheckFlag("medium", "-", "structure",
                                      "the PY's organisation was inferred from tab names because the model's "
                                      "reading was unavailable or failed checks -- confirm the tabs", "structure"))

    # PY facts for the leftover scan: from the PY's working tabs, before any edit.
    working = {n for n, k in structure.kinds.items() if k in ("selection", "edit")}
    if structure.table:
        working.add(structure.table[0])
    py_texts, py_values, py_evidence = [], [], []
    for t in inv.tabs:
        if t.name not in working:
            continue
        g = _grid(work, t)
        _pn, _mk, ev = _py_reading(t, g)
        py_evidence.append(ev)
        py_texts += [i.text for i in t.items if i.text]
        # Record rows (the selection's own line from the listing) are facts;
        # every other cell is text, judged like the shapes' text.
        if t.name in structure.rows:
            record_rows = set(structure.rows[t.name][1]) or {structure.rows[t.name][0] + 1}
        elif structure.table and structure.table[0] == t.name:
            record_rows = set(range(structure.table[2], structure.table[3] + 1))
        else:
            record_rows = {2} if structure.kinds.get(t.name) == "selection" else set()
        for ref, c in t.cells.items():
            row = int(re.sub(r"[A-Z]+", "", ref))
            (py_values if row in record_rows else py_texts).append(c.text)
    py_fact_list, py_name_candidates = verify.py_facts(py_texts, py_values, " ".join(py_evidence))

    # --- 2. CY tabs laid out like the PY's ---
    layout = _arrange(work, inv, structure, len(facts), notes)
    inv = build_inventory(work)
    by_name = {t.name: t for t in inv.tabs}
    tab_names = {f.sample_id: layout[i][0][0] for i, f in enumerate(facts) if layout[i]}

    edited_tabs: set[str] = set()
    checks: list[tuple[str, list, str]] = []   # (tab, text changes, facts) for the final model pass

    def relocate(exhibit: Exhibit, wanted: str, lines: list[str]) -> str | None:
        """A second look (short answer, cheap): which exact text on this
        page is the value the tickmark was meant for?"""
        if not lines:
            return None
        raw, err = call_model(client, model,
                              f"VALUE WANTED: {wanted}\nDOCUMENT: {exhibit.source_file} p.{exhibit.page + 1}\n"
                              "PAGE TEXT LINES:\n" + "\n".join(lines[:120]),
                              ledger=ledger, label="tickmark second look", system=RELOCATE_SYSTEM, max_tokens=200)
        text = str(raw.get("text", "")).strip() if not err else ""
        return text or None

    def edit(tab: TabInventory, purpose: str, f: SelectionFacts | None, exhibits: list[Exhibit],
             data_row: int | None, header_row: int | None, header: dict[str, str]):
        g = _grid(work, tab)
        pic_notes, marks, _ev = _py_reading(tab, g)

        def cell_box(heading: str) -> Slot | None:
            if data_row is None:
                return None
            col = next((c for c, h in sorted(header.items(), key=lambda kv: -_match(heading, kv[1]))
                        if _match(heading, h)), None)
            if not col:
                return None
            i = ox.col_letter_to_index(col) - 1
            return Slot(g.x(i, 0), g.y(data_row - 1, 0), g.x(i + 1, 0) - g.x(i, 0), g.y(data_row, 0) - g.y(data_row - 1, 0))

        skip = {r for r in (header_row, data_row) if r}
        prompt = tab_prompt(tab, purpose=purpose, facts_lines=_facts_lines(f, facts, spec), exhibits=exhibits,
                            marks=marks, picture_notes=pic_notes, grid=g, skip_rows=skip)
        raw, err = call_model(client, model, prompt, ledger=ledger, label=f"tab {tab.name}")
        plan = validate(raw, tab, {x.id for x in exhibits}, error=err)
        report = ApplyReport()
        if tab.drawing_part:
            report = apply_tab(work, tab, plan, {x.id: x for x in exhibits}, str(support_dir),
                               data_row_cell_box=cell_box, relocate=relocate)
        sheet_path = os.path.join(work, "xl", "worksheets", f"sheet{tab.sheet_num}.xml")
        if plan.cells:
            set_cells(sheet_path, {ref: (None if e["action"] == "clear" else e.get("value"))
                                   for ref, e in plan.cells.items()}, sst, {})
        if any(c.formula for c in tab.cells.values()):
            recalc_on_open(work, sheet_path)
        notes.extend(f"[{tab.name}] {n}" for n in plan.notes + report.notes)
        for eid in plan.unaddressed:
            item = next((i for i in tab.items if i.id == eid), None)
            if item is not None:
                flags.append(verify.CheckFlag("medium", tab.name, f"{item.kind} {item.name}",
                                              "not addressed by the editor; left exactly as in the PY", "placement"))
        flags.extend(verify.CheckFlag("medium", tab.name, "-", u, "placement") for u in report.unplaced)
        flags.extend(verify.structural_checks(tab.name, report.text_changes))
        flags.extend(verify.support_flags(tab.name, plan.support))
        for s in plan.support:
            x = next((e for e in exhibits if e.id == s["cy_exhibit"]), None)
            support_rows.append({**s, "tab": tab.name,
                                 "cy": f"{x.source_file} p.{x.page + 1}" if x else ""})
        checks.append((tab.name, report.text_changes, "\n".join(_facts_lines(f, facts, spec))))
        edited_tabs.add(tab.name)
        return plan, report

    # --- 3/4. selection tabs ---
    groups_done: list[tuple[str, list[str], list, dict]] = []   # (selection, tabs, declared, placed labels)
    tick_lines: dict[str, list[str]] = {}
    for i, f in enumerate(facts):
        group = layout[i] if i < len(layout) else []
        declared: list[tuple[str, dict]] = []
        placed_labels: dict[str, list[tuple[str, str]]] = {}
        for j, (cy_name, py_name) in enumerate(group):
            tab = by_name[cy_name]
            rows = structure.rows.get(py_name)
            header_row, data_row, header = (None, None, {})
            if j == 0 or rows:
                header_row, data_row, header = _write_record(work, tab, f, rows, sst, flags, notes)
            others = [n for n, _p in group if n != cy_name]
            purpose = (f"documents CY selection {f.sample_id}"
                       + (f" (tab {j + 1} of {len(group)} for this selection; the others: {others})" if len(group) > 1 else ""))
            plan, report = edit(tab, purpose, f, f.exhibits, data_row, header_row, header)
            _collect_tickmarks(tab, plan, report, f.exhibits, declared, placed_labels, tick_lines)
        groups_done.append((f.sample_id, [n for n, _p in group], declared, placed_labels))

    # --- 4. tabs for the whole control (lead sheets, conclusions) ---
    all_exhibits = {x.source_file + str(x.page): x for f in facts for x in f.exhibits}
    for name, kind in structure.kinds.items():
        if kind == "edit" and name in by_name and name not in edited_tabs:
            edit(by_name[name], f"the control's {name} tab (whole control; CY selection tabs: "
                                f"{[n for g in layout for n, _p in g]})", None, list(all_exhibits.values()),
                 None, None, {})

    # --- selections as rows on one tab ---
    if structure.table and structure.table[0] in by_name:
        tname, hr, first, last = structure.table
        if _write_rows(work, by_name[tname], hr, first, last, [f.data_row for f in facts if f.data_row], sst, notes):
            edited_tabs.add(tname)
        notes.append("the PY lists selections as rows on one tab: conclusions are on the Summary tab")

    # --- 5. population and other PY evidence ---
    for name, kind in structure.kinds.items():
        if name not in by_name:
            continue
        if kind == "population":
            hr = structure.population[1] if structure.population and structure.population[0] == name else 0
            if population_records and _write_rows(work, by_name[name], hr, 0, 0, population_records, sst, notes):
                notes.append(f"[{name}] replaced with this year's population ({len(population_records)} rows)")
                continue
            flags.append(verify.CheckFlag("medium", name, "-", "still holds the PY population -- upload the CY "
                                          "population to have it written here", "PY data"))
        elif kind in ("parameters", "py_evidence"):
            flags.append(verify.CheckFlag("medium", name, "-", f"still holds last year's {kind.replace('_', ' ')} "
                                          "-- replace with this year's", "PY data"))

    # --- 6. the leftover scan, over every working tab of the CY file ---
    sst.write()
    cy_fact_text = "\n".join(_facts_lines(None, facts, spec)) + "\n" + "\n".join(
        x.text + " " + x.description for f in facts for x in f.exhibits)
    common = " ".join(c.text for t in inv.tabs for ref, c in t.cells.items()
                      if t.name in structure.rows and re.fullmatch(rf"[A-Z]+{structure.rows[t.name][0]}", ref))
    locations = _cy_text_locations(work, edited_tabs)
    flags.extend(verify.leftover_scan(py_fact_list, locations, cy_fact_text, common))
    candidates = verify.name_hits(py_name_candidates, locations, cy_fact_text)
    # Tickmarks: cited in text <-> placed on evidence, across each selection's tabs.
    py_labels = verify.py_cited_labels([t for t in build_inventory(str(output_path) + ".pyread").tabs]) \
        if os.path.isdir(str(output_path) + ".pyread") else set()
    for sel, tabs_, declared, placed_labels in groups_done:
        labels = {re.sub(r"\s+", " ", x.strip()) for x in py_labels} | \
                 {re.sub(r"\s+", " ", t["label"].strip()) for _t, t in declared}
        texts = [(tab, where, text, kind_is_label(text, labels)) for tab, where, text, _k in
                 _cy_text_locations(work, set(tabs_))]
        flags.extend(verify.tickmark_checks(sel, py_labels, declared, placed_labels, texts))

    for tab_name, changes, facts_text in checks:
        if run_model_check:
            flags.extend(verify.model_check(client, model, tab_name, changes, facts_text, ledger=ledger,
                                            name_candidates=candidates.get(tab_name),
                                            tickmarks=tick_lines.get(tab_name)))
        elif candidates.get(tab_name):
            flags.extend(verify.CheckFlag("low", tab_name, "-", f"possible PY name {n!r} in CY text -- check",
                                          "leftovers") for n in candidates[tab_name])

    if summary_rows_fn is not None:
        add_summary_sheet(work, summary_rows_fn(tab_names), sst)
    flags = verify.sort_flags(flags)
    rows = [[f.severity.upper(), f.tab, f.item, f.issue, f.found_by] for f in flags] or \
           [["OK", "-", "-", "No differences from the prior-year layout and wording were found.", "all checks"]]
    pos = 1 if summary_rows_fn is not None else 0
    add_table_sheet(work, "Mirror Check", MIRROR_CHECK_COLUMNS, rows, sst, position=pos,
                    red_rows={i for i, f in enumerate(flags) if f.severity == "high"},
                    amber_rows={i for i, f in enumerate(flags) if f.severity == "medium"})
    srows = [[s["match"].upper(), s["tab"], s["py_document"], s.get("cy") or "-", s.get("comment", "")]
             for s in support_rows] or [["-", "-", "No PY documents were identified on the selection tabs.", "-", ""]]
    add_table_sheet(work, "Support Check", SUPPORT_CHECK_COLUMNS, srows, sst, position=pos + 1,
                    red_rows={i for i, s in enumerate(support_rows) if s["match"] == "missing"},
                    amber_rows={i for i, s in enumerate(support_rows) if s["match"] == "partial"})
    sst.write()

    for num in ox.resolve_sheets(work).values():
        ox.stamp_print_header(os.path.join(work, "xl", "worksheets", f"sheet{num}.xml"), DRAFT_BANNER)
    _ensure_png(work)
    ox.repair_builtin_name_scopes(work)
    ox.remove_orphan_media(work)
    ox.remove_calc_chain(work)
    problems = ox.package_problems(work)
    if problems:
        raise ox.TemplateMismatch("the mirrored file failed its integrity check: " + "; ".join(problems[:5]))
    ox.repackage(work, str(output_path))
    return MirrorOutcome(output_path, flags, notes, tab_names, support_rows)
