"""How the PY workpaper is organised -- read by the model, checked by code.

Every earlier version decided this with name patterns: a tab called "1" or
"Sample" is one selection, "lead" is the lead sheet, and so on. Each
pattern is a guess, and a PY that breaks one -- two tabs per selection
("1a", "1b"), a tab per test step, a lead sheet called "Overview", a lead
written in cells -- was mangled or left as last year's without a word.

So the model is shown every tab (its name, its cells, the text and kind of
its shapes) and asked how the workbook is organised. Code then checks the
answer against the workbook itself; anything it cannot verify is refused,
and the name patterns are used only as a flagged fallback.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Any

from agent.mirror.inventory import Inventory, TabInventory

KINDS = ("selection", "edit", "population", "parameters", "py_evidence", "static")

SYSTEM = """\
You are reading last year's (PY) SOX testing workbook to work out how it is
organised, so this year's (CY) workbook can be built with exactly the same
organisation. You do not edit anything.

For every tab, say what it is:
  selection    documents ONE PY selection (sampled item) -- or one part of one
               selection's testing, when a selection spans several tabs
  edit         holds testing narrative/results/sign-offs for the whole control
               that must be updated for CY (a lead sheet, a conclusion tab) --
               whether written in text boxes or in cells
  population   last year's population listing
  parameters   last year's report parameters / IPE evidence
  py_evidence  other last-year data that CY must replace (report extracts etc.)
  static       nothing year-specific (dividers, instructions, blank tabs)

Group selection tabs by PY selection, in order: which tabs together document
selection 1, which selection 2, and so on. A selection may have one tab or
several (e.g. "1a"/"1b", or a tab per test step).

If instead the PY lists selections as ROWS on one tab (no per-selection tabs),
give that tab as selection_table with its heading row and the rows used.

For each selection tab, give its heading row and the row(s) holding that
selection's own record (the line(s) from the listing/extract), if any.

Then name the CY tabs: for each CY selection, the names of its tab(s), in the
PY's own naming convention (if PY used "1","2" use "1","2"; if "1a","1b" use
"2a","2b" for the second selection; and so on). Max 31 characters, none of
: \\ / ? * [ ].

Reply with JSON only:
{"tabs": [{"name": "...", "kind": "...", "why": "<= 12 words"}],
 "py_selections": [{"py_id": "1", "tabs": ["..."]}],
 "selection_rows": [{"tab": "...", "header_row": 1, "data_rows": [2]}],
 "selection_table": {"tab": "...", "header_row": 3, "first_row": 4, "last_row": 4} or null,
 "population": {"tab": "...", "header_row": 1} or null,
 "cy_tab_names": [["..."], ["..."]],
 "notes": ["anything unusual about this workbook"]}
"""


@dataclass
class Structure:
    kinds: dict[str, str]                         # tab -> kind
    groups: list[list[str]]                       # PY selection -> its tabs, in order
    rows: dict[str, tuple[int, list[int]]]        # selection tab -> (header row, data rows)
    table: tuple[str, int, int, int] | None       # (tab, header row, first, last)
    population: tuple[str, int] | None
    cy_names: list[list[str]]                     # CY selection -> its tab names
    notes: list[str] = field(default_factory=list)
    from_model: bool = True


def _tab_digest(tab: TabInventory, max_rows: int = 14) -> str:
    rows: dict[int, list[str]] = {}
    for ref, c in tab.cells.items():
        m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
        if m and c.text.strip():
            rows.setdefault(int(m.group(2)), []).append(f"{m.group(1)}={c.text.strip()[:40]}")
    lines = [f"TAB {tab.name!r}: {len(tab.cells)} non-empty cells over {len(rows)} rows; "
             f"shapes: " + (", ".join(f"{k}x{n}" for k, n in _count(tab).items()) or "none")]
    for r in sorted(rows)[:max_rows]:
        lines.append(f"  row {r}: " + "; ".join(rows[r][:14]))
    if len(rows) > max_rows:
        lines.append(f"  ... rows up to {max(rows)}")
    for item in [i for i in tab.items if i.kind in ("text", "group")][:12]:
        lines.append(f"  shape {item.kind} '{item.name}': {' '.join(item.text.split())[:160]}")
    return "\n".join(lines)


def _count(tab: TabInventory) -> dict[str, int]:
    out: dict[str, int] = {}
    for i in tab.items:
        out[i.kind] = out.get(i.kind, 0) + 1
    return out


def structure_prompt(inv: Inventory, cy_selections: list[tuple[str, dict]]) -> str:
    parts = ["PY WORKBOOK:"] + [_tab_digest(t) for t in inv.tabs]
    parts += ["", f"CY: {len(cy_selections)} selection(s): "
              + "; ".join(f"{sid} = {json.dumps(row, default=str)[:160]}" for sid, row in cy_selections)]
    return "\n".join(parts)


_BAD_NAME = re.compile(r"[:\\/?*\[\]]")


def validate(raw: dict, inv: Inventory, n_cy: int) -> tuple[Structure | None, list[str]]:
    """The model's reading, checked against the workbook. (None, problems)
    when it cannot be trusted."""
    names = [t.name for t in inv.tabs]
    problems: list[str] = []
    kinds: dict[str, str] = {}
    for entry in raw.get("tabs", []) or []:
        n, k = str(entry.get("name", "")), str(entry.get("kind", ""))
        if n in names and k in KINDS:
            kinds[n] = k
    for n in names:
        if n not in kinds:
            problems.append(f"tab {n!r} was not described")

    groups: list[list[str]] = []
    seen: set[str] = set()
    for g in raw.get("py_selections", []) or []:
        tabs = [str(t) for t in g.get("tabs", []) or []]
        if not tabs or any(t not in names for t in tabs):
            problems.append(f"selection group {tabs} names a tab that does not exist")
            continue
        if seen & set(tabs):
            problems.append(f"a tab is in two selection groups: {sorted(seen & set(tabs))}")
            continue
        seen |= set(tabs)
        groups.append(tabs)
    stray = [n for n, k in kinds.items() if k == "selection" and n not in seen]
    if stray:
        problems.append(f"selection tabs in no group: {stray}")

    rows: dict[str, tuple[int, list[int]]] = {}
    for r in raw.get("selection_rows", []) or []:
        t = str(r.get("tab", ""))
        try:
            hr = int(r.get("header_row") or 0)
            dr = [int(x) for x in r.get("data_rows", []) or []]
        except (TypeError, ValueError):
            continue
        if t in seen and hr >= 0:
            rows[t] = (hr, dr)

    table = None
    st = raw.get("selection_table")
    if isinstance(st, dict) and str(st.get("tab")) in names:
        try:
            table = (str(st["tab"]), int(st["header_row"]), int(st["first_row"]), int(st["last_row"]))
        except (KeyError, TypeError, ValueError):
            problems.append("selection_table rows are not numbers")
    pop = raw.get("population")
    population = None
    if isinstance(pop, dict) and str(pop.get("tab")) in names:
        try:
            population = (str(pop["tab"]), int(pop.get("header_row") or 1))
        except (TypeError, ValueError):
            pass

    cy_names: list[list[str]] = []
    if groups:
        proposed = raw.get("cy_tab_names", []) or []
        for i in range(n_cy):
            want = len(groups[min(i, len(groups) - 1)])
            got = [str(x)[:31] for x in proposed[i]] if i < len(proposed) and isinstance(proposed[i], list) else []
            if len(got) != want or any(not x.strip() or _BAD_NAME.search(x) for x in got):
                problems.append(f"CY tab names for selection {i + 1} are missing or invalid: {got}")
                got = []
            cy_names.append(got)
        flat = [n for g in cy_names for n in g]
        others = {n.lower() for n in names if n not in seen}
        if len({n.lower() for n in flat}) != len(flat) or any(n.lower() in others for n in flat):
            problems.append("CY tab names repeat, or clash with another tab")
            cy_names = [[] for _ in range(n_cy)]

    if not groups and table is None and any(k in ("selection",) for k in kinds.values()):
        problems.append("selection tabs were named but not grouped")
    if len(kinds) < max(1, len(names) // 2):
        problems.append("most tabs were not described")
        return None, problems
    if any(p for p in problems if "does not exist" in p or "two selection groups" in p or "not grouped" in p):
        return None, problems
    # A tab the reading skipped is left exactly as in the PY and flagged --
    # "py_evidence" is the kind that gets a reviewer's attention.
    for n in names:
        kinds.setdefault(n, "py_evidence")
    return Structure(kinds, groups, rows, table, population, cy_names,
                     [str(n) for n in raw.get("notes", []) or []]), problems


def fallback(inv: Inventory, n_cy: int) -> Structure:
    """Name patterns, when the model's reading is unavailable. Flagged."""
    role_to_kind = {"selection": "selection", "lead": "edit", "population": "population",
                    "parameters": "parameters", "other": "py_evidence", "divider": "static",
                    "selection_table": "py_evidence"}
    kinds = {t.name: role_to_kind.get(t.role, "py_evidence") for t in inv.tabs}
    groups = [[t.name] for t in inv.tabs if t.role == "selection"]
    table = None
    tbl = next((t for t in inv.tabs if t.role == "selection_table"), None)
    if not groups and tbl is not None:
        kinds[tbl.name] = "selection"
        table = (tbl.name, 0, 0, 0)  # rows found by heading match
    pop = next((t.name for t in inv.tabs if t.role == "population"), None)
    return Structure(kinds, groups, {}, table, (pop, 1) if pop else None, [[] for _ in range(n_cy)],
                     ["the PY's organisation was inferred from tab names, not read by the model -- check it"],
                     from_model=False)


def read_structure(client: Any, model: str, inv: Inventory, cy_selections: list[tuple[str, dict]],
                   *, ledger: Any = None) -> tuple[Structure, list[str]]:
    from agent.mirror.rewrite import call_model

    raw, err = call_model(client, model, structure_prompt(inv, cy_selections), ledger=ledger,
                          label="structure", system=SYSTEM)
    if err:
        return fallback(inv, len(cy_selections)), [f"structure read failed ({err})"]
    found, problems = validate(raw, inv, len(cy_selections))
    if found is None:
        return fallback(inv, len(cy_selections)), problems
    return found, problems
