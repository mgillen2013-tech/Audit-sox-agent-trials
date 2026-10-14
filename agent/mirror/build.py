"""Build a CY workpaper by editing a copy of the PY. See agent/mirror/__init__.py."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from agent.mirror import verify
from agent.mirror.apply import (
    ApplyReport, apply_tab, recalc_on_open, remove_embedded_objects, set_cells,
)
from agent.mirror.facts import selection_facts
from agent.mirror.inventory import Inventory, TabInventory, build_inventory
from agent.mirror.rewrite import (
    SelectionFacts, call_model, lead_prompt, tab_prompt, validate,
)
from agent.ooxml import imaging
from agent.ooxml import ooxml_utils as ox
from agent.ooxml.layout import Slot, _Grid, _row_heights
from agent.ooxml.sample_grid import _match, plan_grid
from agent.ooxml.summary_sheet import DRAFT_BANNER, GENERATED_HEADINGS, add_summary_sheet, add_table_sheet

MIRROR_CHECK_COLUMNS = [("A", "Severity", 10), ("B", "Tab", 16), ("C", "Item", 18),
                        ("D", "Issue", 90), ("E", "Found by", 14)]
_CHECK_HEADINGS = tuple(t for _c, t, _w in MIRROR_CHECK_COLUMNS[:4])


@dataclass
class MirrorOutcome:
    path: Path
    flags: list[verify.CheckFlag] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    tab_names: dict[str, str] = field(default_factory=dict)  # sample_id -> CY tab


# ---------------------------------------------------------------------------


def _remove_generated_tabs(work_dir: str) -> None:
    """A previous run's Summary / Mirror Check, recognised by their headings."""
    for name, num in list(ox.resolve_sheets(work_dir).items()):
        rows = ox.first_rows_text(work_dir, num, rows=2)
        if any(tuple(r[:3]) == GENERATED_HEADINGS or tuple(r[:4]) == _CHECK_HEADINGS for r in rows):
            ox.delete_sheet(work_dir, name)


def _cy_names(py_tabs: list[str], facts: list[SelectionFacts]) -> list[str]:
    """CY tab names in the PY's own convention."""
    n = len(facts)
    if py_tabs and all(t.strip().isdigit() for t in py_tabs):
        return [str(i + 1) for i in range(n)]
    base = py_tabs[0] if py_tabs else "Sample"
    if re.fullmatch(r"sample(?:\s+.*)?", base, re.I):
        from agent.ooxml.build_workpaper import sample_tab_name

        return [sample_tab_name("Sample", f.sample_id, i, n) for i, f in enumerate(facts)]
    return [(py_tabs[i] if i < len(py_tabs) else f"{base} ({i + 1})")[:31] for i in range(n)]


def _arrange_selection_tabs(work_dir: str, inv: Inventory, facts: list[SelectionFacts]) -> dict[str, str]:
    """One CY tab per selection, built from the PY tab at the same position
    (or the PY's first, for selections beyond the PY's count). Surplus PY
    selection tabs -- last year's items -- are removed. Returns
    {sample_id: CY tab name}."""
    py = [t for t in inv.tabs if t.role == "selection"]
    names = _cy_names([t.name for t in py], facts)
    # Clones first, from the pristine PY tab, under temporary names.
    temp: list[str] = []
    after = py[-1].name
    for i in range(len(py), len(facts)):
        t = f"__cy_{i}"
        ox.clone_sheet(work_dir, py[0].sheet_num, t, after=after, keep_drawing=True)
        temp.append(t)
        after = t
    for t in py[len(facts):]:
        ox.delete_sheet(work_dir, t.name)
    current = [t.name for t in py[: len(facts)]] + temp
    # Two-step rename so "2" -> "1" style swaps can never collide.
    sheets = ox.resolve_sheets(work_dir)
    for i, name in enumerate(current):
        ox.rename_sheet(work_dir, sheets[name], f"__rename_{i}")
    sheets = ox.resolve_sheets(work_dir)
    for i, final in enumerate(names):
        ox.rename_sheet(work_dir, sheets[f"__rename_{i}"], final)
    return {f.sample_id: names[i] for i, f in enumerate(facts)}


def _cells_for_grid(tab: TabInventory) -> dict[str, tuple]:
    return {ref: (c.style, c.text, c.formula) for ref, c in tab.cells.items()}


def _listing(facts: SelectionFacts) -> dict[str, str]:
    return {ox.col_index_to_letter(i + 1): k for i, k in enumerate(facts.data_row)}


def _grid_geometry(work_dir: str, tab: TabInventory) -> _Grid:
    path = os.path.join(work_dir, "xl", "worksheets", f"sheet{tab.sheet_num}.xml")
    with open(path, encoding="utf-8") as fh:
        default_h, custom_h = _row_heights(fh.read())
    return _Grid(ox.parse_column_widths(path), default_h, custom_h)


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


def _py_notes(tab: TabInventory, header: dict[str, str], data_row: int, grid: _Grid):
    """What each PY picture showed, and what each PY tickmark marked -- read
    from the PY itself, so the model can match roles without guessing."""
    pictures = tab.by_kind("picture")
    pic_notes = {}
    for p in pictures:
        im, words = _ocr(p.image_path) if p.image_path else (None, [])
        text = " ".join(w.text for w in sorted(words, key=lambda w: (w.top, w.left)))
        pic_notes[p.id] = (text[:300] or "(no readable text)")
    from agent.mirror.apply import _pair_tickmarks

    tm_notes = {}
    row_top = grid.y(data_row - 1, 0)
    row_bottom = grid.y(data_row, 0)
    for label_id, box in _pair_tickmarks(tab).items():
        if box is None:
            continue
        if box.box.y < row_bottom and box.box.bottom > row_top - 200_000:
            cols = [c for c in header if grid.x(ox.col_letter_to_index(c) - 1, 0) < box.box.right
                    and grid.x(ox.col_letter_to_index(c), 0) > box.box.x]
            tm_notes[label_id] = "data row: " + ", ".join(
                f"{header[c]}={tab.cells.get(f'{c}{data_row}').text if tab.cells.get(f'{c}{data_row}') else ''}"
                for c in cols)
            continue
        host = next((p for p in pictures if p.box.x <= box.box.x <= p.box.right and p.box.y <= box.box.y <= p.box.bottom), None)
        if host is None:
            continue
        im, words = _ocr(host.image_path) if host.image_path else (None, [])
        if im is None:
            continue
        sx, sy = im.width / max(1, host.box.cx), im.height / max(1, host.box.cy)
        x0, y0 = (box.box.x - host.box.x) * sx, (box.box.y - host.box.y) * sy
        x1, y1 = (box.box.right - host.box.x) * sx, (box.box.bottom - host.box.y) * sy
        inside = [w.text for w in words if w.left + w.width / 2 >= x0 and w.left + w.width / 2 <= x1
                  and w.top + w.height / 2 >= y0 and w.top + w.height / 2 <= y1]
        tm_notes[label_id] = f"{' '.join(inside) or '(value not readable)'} on PY picture {host.id}"
    return pic_notes, tm_notes


def _ensure_png(work_dir: str) -> None:
    ct_path = os.path.join(work_dir, "[Content_Types].xml")
    with open(ct_path, encoding="utf-8") as fh:
        ct = fh.read()
    if not re.search(r'<Default\b[^>]*Extension="png"', ct, re.I):
        ct = ct.replace("<Override", '<Default Extension="png" ContentType="image/png"/><Override', 1) \
            if "<Override" in ct else ct.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>', 1)
        with open(ct_path, "w", encoding="utf-8") as fh:
            fh.write(ct)


def _facts_text(f: SelectionFacts) -> str:
    import json

    return json.dumps({"data_row": f.data_row, "steps": f.steps,
                       "exhibits": [x.__dict__ for x in f.exhibits]}, default=str)


# ---------------------------------------------------------------------------


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
) -> MirrorOutcome:
    work = str(output_path) + ".mirror"
    ox.extract_template(str(py_path), work)
    _remove_generated_tabs(work)
    inv = build_inventory(work)
    facts = selection_facts(results, spec, manifests, support_dir)
    sst = ox.SharedStrings(work)
    flags: list[verify.CheckFlag] = []
    notes: list[str] = []
    tab_names: dict[str, str] = {}

    has_selection_tabs = any(t.role == "selection" for t in inv.tabs)
    if has_selection_tabs:
        tab_names = _arrange_selection_tabs(work, inv, facts)
        inv = build_inventory(work)
    by_name = {t.name: t for t in inv.tabs}

    # --- each selection tab: data row, then the model's edits ---
    for f in facts if has_selection_tabs else []:
        tab = by_name[tab_names[f.sample_id]]
        sheet_path = os.path.join(work, "xl", "worksheets", f"sheet{tab.sheet_num}.xml")
        listing = _listing(f)
        grid_plan = plan_grid(_cells_for_grid(tab), listing) if listing else None
        header_row = grid_plan.header_row if grid_plan else 1
        data_row = grid_plan.data_row if grid_plan else 2
        header = {re.sub(r"\d+", "", ref): c.text for ref, c in tab.cells.items()
                  if re.fullmatch(rf"[A-Z]+{header_row}", ref)}
        geometry = _grid_geometry(work, tab)

        if grid_plan and grid_plan.headings_to_write and len(grid_plan.headings_to_write) == len(listing):
            # Nothing on the PY tab matched the listing's columns. Writing
            # our own headings over the PY's would be the opposite of
            # mirroring, and writing values under the PY's would label them
            # wrongly -- so neither: the row is left for a reviewer.
            flags.append(verify.CheckFlag(
                "high", tab.name, "data row",
                "this year's sample listing has none of the PY tab's column headings, so the "
                "selection's row was not written -- check the listing is the same report as last year",
                "structure"))
            grid_plan = None
        if grid_plan:
            values: dict[str, Any] = {f"{c}{data_row}": None for c in
                                      {re.sub(r"\d+", "", r) for r in tab.cells if re.fullmatch(rf"[A-Z]+{data_row}", r)}}
            for letter, heading in listing.items():
                values[f"{grid_plan.column_for[letter]}{data_row}"] = f.data_row.get(heading)
            for col, heading in grid_plan.headings_to_write.items():
                values[f"{col}{header_row}"] = heading
            set_cells(sheet_path, values, sst, {})
            notes += [f"[{tab.name}] {n}" for n in grid_plan.notes]

        def cell_box(heading: str, _tab=tab, _row=data_row, _g=geometry, _h=header) -> Slot | None:
            col = next((c for c, h in sorted(_h.items(), key=lambda kv: -_match(heading, kv[1])) if _match(heading, h)), None)
            if not col:
                return None
            i = ox.col_letter_to_index(col) - 1
            x0, x1 = _g.x(i, 0), _g.x(i + 1, 0)
            y0, y1 = _g.y(_row - 1, 0), _g.y(_row, 0)
            return Slot(x0, y0, x1 - x0, y1 - y0)

        pic_notes, tm_notes = _py_notes(tab, header, data_row, geometry)
        prompt = tab_prompt(tab, f, picture_notes=pic_notes, tickmark_notes=tm_notes,
                            data_rows={data_row}, header_row=header_row)
        raw, err = call_model(client, model, prompt, ledger=ledger, label=f"tab {tab.name}")
        plan = validate(raw, tab, {x.id for x in f.exhibits}, error=err)
        report = ApplyReport()
        if tab.drawing_part:
            report = apply_tab(work, tab, plan, {x.id: x for x in f.exhibits}, str(support_dir),
                               data_row_cell_box=cell_box)
        if plan.cells:
            set_cells(sheet_path, {ref: (None if e["action"] == "clear" else e.get("value"))
                                   for ref, e in plan.cells.items()}, sst, {})
        if any(c.formula for c in tab.cells.values()):
            recalc_on_open(work, sheet_path)
        if remove_embedded_objects(work, tab.sheet_num):
            report.removed.append("an embedded file (PY evidence)")
        notes += [f"[{tab.name}] {n}" for n in plan.notes]

        flags += [verify.CheckFlag("medium", tab.name, "-", u, "placement") for u in report.unplaced]
        flags += verify.structural_checks(tab.name, report.text_changes)
        py_row_values = [c.text for ref, c in inv_py_row(work, tab, data_row)]
        flags += verify.leftover_checks(tab.name, report.text_changes, py_row_values, _facts_text(f))
        if run_model_check:
            flags += verify.model_check(client, model, tab.name, report.text_changes, _facts_text(f), ledger=ledger)

    # --- the lead sheet ---
    all_steps = [s for f in facts for s in f.steps]
    for tab in [t for t in inv.tabs if t.role == "lead" and t.drawing_part]:
        raw, err = call_model(client, model, lead_prompt(tab, all_steps, list(tab_names.values())),
                              ledger=ledger, label=f"lead {tab.name}")
        plan = validate(raw, tab, set(), error=err)
        report = apply_tab(work, tab, plan, {}, str(support_dir), data_row_cell_box=lambda _h: None)
        flags += verify.structural_checks(tab.name, report.text_changes)
        if run_model_check:
            flags += verify.model_check(client, model, tab.name, report.text_changes,
                                        "\n".join(_facts_text(f) for f in facts), ledger=ledger)

    # --- one selection table instead of per-item tabs ---
    if not has_selection_tabs:
        for tab in [t for t in inv.tabs if t.role == "selection_table"][:1]:
            _write_selection_table(work, tab, facts, sst, notes)
        notes.append("the prior-year file has no per-selection tabs: conclusions are on the Summary tab")

    # --- tabs still holding last year's data ---
    for t in inv.tabs:
        if t.role in ("population", "parameters", "other") and (t.cells or t.items):
            flags.append(verify.CheckFlag(
                "medium", t.name, "-",
                "still holds the prior year's data (report extract / parameters screenshot) -- replace with CY",
                "PY data"))

    sst.write()
    if summary_rows_fn is not None:
        add_summary_sheet(work, summary_rows_fn(tab_names), sst)
    flags = verify.sort_flags(flags)
    rows = [[f.severity.upper(), f.tab, f.item, f.issue, f.found_by] for f in flags] or \
           [["OK", "-", "-", "No differences from the prior-year layout and wording were found.", "all checks"]]
    # Right after the Summary (or first, without one): what to look at
    # before the evidence.
    add_table_sheet(work, "Mirror Check", MIRROR_CHECK_COLUMNS, rows, sst,
                    position=1 if summary_rows_fn is not None else 0,
                    red_rows={i for i, f in enumerate(flags) if f.severity == "high"},
                    amber_rows={i for i, f in enumerate(flags) if f.severity == "medium"})
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
    return MirrorOutcome(output_path, flags, notes, tab_names)


def inv_py_row(work: str, tab: TabInventory, data_row: int):
    """The PY's own values on the data row (from the inventory, read before
    the row was rewritten) -- the facts most likely to leak into CY text."""
    return [(ref, c) for ref, c in tab.cells.items() if re.fullmatch(rf"[A-Z]+{data_row}", ref)]


def _write_selection_table(work: str, tab: TabInventory, facts: list[SelectionFacts], sst, notes: list[str]) -> None:
    listing = _listing(facts[0]) if facts and facts[0].data_row else {}
    if not listing:
        return
    grid_plan = plan_grid(_cells_for_grid(tab), listing)
    if grid_plan.headings_to_write and len(grid_plan.headings_to_write) == len(listing):
        notes.append(f"[{tab.name}] this year's sample listing has none of the PY table's column "
                     "headings, so last year's selections were left for a reviewer to replace")
        return
    sheet_path = os.path.join(work, "xl", "worksheets", f"sheet{tab.sheet_num}.xml")
    py_rows = sorted({int(re.sub(r"[A-Z]+", "", r)) for r in tab.cells
                      if int(re.sub(r"[A-Z]+", "", r)) > grid_plan.header_row})
    values: dict[str, Any] = {}
    for r in py_rows:  # clear last year's selections
        for ref in [x for x in tab.cells if re.fullmatch(rf"[A-Z]+{r}", x)]:
            values[ref] = None
    for i, f in enumerate(facts):
        row = grid_plan.data_row + i
        for letter, heading in _listing(f).items():
            col = grid_plan.column_for.get(letter)
            if col:
                values[f"{col}{row}"] = f.data_row.get(heading)
    set_cells(sheet_path, values, sst, {})
    notes += [f"[{tab.name}] {n}" for n in grid_plan.notes]
