"""The judgment step: "make this PY tab say what is true for CY".

The model is shown every element of a PY tab as it physically is -- each
shape with its text (formatting marked), what it sits on, the element
nearest to it; each picture and what it showed; each line, group and
embedded file; the cells -- next to the CY facts. It decides, per element,
what it is and what happens to it: keep, rewrite, remove, re-place onto
the CY evidence, or re-connect. Code decides nothing about meaning: no
"a label is TM + digits", no "a box belongs to the nearest label", no
"an unlabelled box is deleted". Anything the model does not address is
KEPT and reported, never removed.

The same call answers the second question: does the CY support mirror
the PY support? Every PY document (picture, embedded file, pasted table)
is matched to its CY counterpart or reported missing.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Any

from agent.mirror.inventory import DrawingItem, TabInventory
from agent.ooxml.layout import Slot

_MAX_TOKENS = 16_000

SYSTEM = """\
You update one tab of last year's (PY) SOX testing workpaper for this year
(CY). You are an editor: the PY is the house style and is kept wherever it is
still true. You decide what each PY element is and what happens to it.

For EVERY element listed, return one entry:
  keep     unchanged (still true for CY, or not year-specific)
  replace  same element, CY text (give "paragraphs")
  remove   no CY counterpart (e.g. a note about a PY-only issue)
  place    move it onto the CY evidence: "on": {"exhibit": "X1", "anchor": "<exact
           text on that page>"} or {"heading": "<data-row heading>"} or
           {"beside": "<id of another element you place>"}. A text element can be
           placed AND given new "paragraphs".
  connect  a line/arrow: give "from" and "to" element ids to re-draw it between

Text rules (strict):
1. Same structure: same paragraphs, order, labels ("Test Step 1:", "TM 2",
   "Note:"), blank lines and <br/> breaks. Only facts change: names, dates,
   amounts, document numbers, counts, and result wording where CY differs.
2. Style tags <s0>..</s0>: each is a format from the PY; use the same tag for
   the same role. Never invent a tag number.
3. Every CY fact must come from CY FACTS. Never carry a PY name, date, amount or
   number into CY. If CY evidence does not show what a PY sentence states, say
   what CY shows, in the PY's own phrasing for that situation.
4. Results must match the testing agent's conclusion per test step.
5. Keep length within ~25% of the PY element.
6. Anchors are copied character-for-character from the CY exhibit text: one
   value, not a sentence.

PICTURES: give each PY picture the CY exhibit playing the same role (invoice
for invoice) or null. If the PY picture showed only PART of a page, give a
"crop": {"from": "<first CY text line of that part>", "to": "<last line>"}.
CY exhibits with no PY picture go in extra_exhibits.

CELLS: update cells that state PY facts (sign-offs, dates, amounts, results).
Formula cells are read-only.

SUPPORT: for every PY document on this tab (each picture, embedded file, group
of screenshots, pasted evidence table), say whether CY has the same kind of
document: {"py_ref": "<id or 'cells'>", "py_document": "<what it is>",
"cy_exhibit": "X1" or null, "match": "same|partial|missing", "comment": "..."}.

Reply with JSON only:
{"elements": [{"id": "...", "action": "...", "paragraphs": ["..."], "on": {...},
               "from": "...", "to": "..."}],
 "pictures": [{"id": "...", "exhibit": "X1" or null, "crop": {"from": "...", "to": "..."} or null}],
 "extra_exhibits": ["X4"],
 "cells": [{"ref": "B5", "action": "keep|replace|clear", "value": "..."}],
 "support": [...],
 "notes": ["anything a reviewer must know"]}
"""


@dataclass
class Exhibit:
    """One CY document page the workpaper can show."""

    id: str             # "X1"
    source_file: str
    page: int           # 0-indexed
    description: str    # what the testing agent said it is
    text: str = ""      # its text, for matching to a PY slot


@dataclass
class SelectionFacts:
    sample_id: str
    data_row: dict[str, Any]                  # listing heading -> value
    steps: list[dict[str, Any]]               # per test step, this selection only
    exhibits: list[Exhibit] = field(default_factory=list)


@dataclass
class TabPlan:
    elements: dict[str, dict] = field(default_factory=dict)
    pictures: dict[str, dict] = field(default_factory=dict)   # id -> {"exhibit", "crop"}
    extra_exhibits: list[str] = field(default_factory=list)
    cells: dict[str, dict] = field(default_factory=dict)
    support: list[dict] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    unaddressed: list[str] = field(default_factory=list)      # element ids the model skipped


# ---------------------------------------------------------------------------
# Prompt
# ---------------------------------------------------------------------------


def tagged(item: DrawingItem) -> list[str]:
    # A line break INSIDE a paragraph is <br/>: shown as a new line it read as
    # an extra paragraph, and a faithful copy was flagged as restructured.
    return ["".join("<br/>" if k == "br" else f"<{k}>{t}</{k}>" for k, t in p.runs)
            for p in item.paragraphs]


def style_legend(tab: TabInventory) -> str:
    parts = []
    for s in tab.styles.values():
        bits = ["bold" if s.bold else "regular"]
        if s.color and s.color.upper() != "000000":
            bits.append(f"colour {s.color}")
        size = re.search(r'\bsz="(\d+)"', s.rpr_xml)
        if size:
            bits.append(f"{int(size.group(1)) / 100:g}pt")
        parts.append(f"{s.key} = {', '.join(bits)}")
    return "; ".join(parts) or "(none)"


def _where(box: Slot, grid) -> str:
    if grid is None:
        return ""
    return f"at {grid.cell_at(box.x, box.y)}-{grid.cell_at(box.right, box.bottom)}"


def _nearest(item: DrawingItem, items: list[DrawingItem]) -> str:
    cx, cy = item.box.x + item.box.cx / 2, item.box.y + item.box.cy / 2
    best, best_d = None, None
    for other in items:
        if other is item or other.kind in ("frame", "picture"):
            continue
        ox_, oy_ = other.box.x + other.box.cx / 2, other.box.y + other.box.cy / 2
        d = ((cx - ox_) ** 2 + (cy - oy_) ** 2) ** 0.5
        if best_d is None or d < best_d:
            best, best_d = other, d
    if best is None:
        return ""
    label = " ".join(best.text.split())[:30]
    return f"; nearest: {best.id} ({best.kind}{', ' + repr(label) if label else ''}, {best_d / 914_400:.1f}in)"


def tab_prompt(
    tab: TabInventory,
    *,
    purpose: str,
    facts_lines: list[str],
    exhibits: list[Exhibit],
    marks: dict[str, str],
    picture_notes: dict[str, str],
    grid=None,
    skip_rows: set[int] = frozenset(),
) -> str:
    lines = [f"PY TAB {tab.name!r} -- {purpose}", "", f"STYLE TAGS: {style_legend(tab)}", "", "PY ELEMENTS:"]
    for item in tab.items:
        if item.kind == "frame":
            continue  # frames move with their picture
        head = f"- {item.id} [{item.kind}] '{item.name}' {_where(item.box, grid)}"
        if item.kind == "picture":
            lines.append(f"{head} showed: {picture_notes.get(item.id, '(unreadable)')}")
            continue
        extra = f"; sits on: {marks[item.id]}" if item.id in marks else ""
        lines.append(head + extra + _nearest(item, tab.items))
        if item.kind == "text":
            for p in tagged(item):
                lines.append(f"    | {p}")
        elif item.kind in ("group",):
            lines.append(f"    | {item.text[:300]}")
        elif item.kind == "object":
            lines.append("    | an embedded file (opens from the workbook)")
    cells = [c for ref, c in tab.cells.items()
             if (m := re.fullmatch(r"[A-Z]+(\d+)", ref)) and int(m.group(1)) not in skip_rows]
    if cells:
        lines += ["", "PY CELLS:"]
        for c in sorted(cells, key=lambda c: (int(re.sub(r"\D", "", c.ref)), len(c.ref), c.ref))[:150]:
            lines.append(f"- {c.ref}: {c.text[:120]!r}" + (" [formula, read-only]" if c.formula else ""))
        if len(cells) > 150:
            lines.append(f"... {len(cells) - 150} more cells not shown (kept as they are)")
    lines += ["", "CY FACTS (the only source of CY content):"] + facts_lines
    lines += ["", "CY EXHIBITS:"]
    for x in exhibits:
        snippet = " ".join(x.text.split())[:500]
        lines.append(f"- {x.id}: {x.source_file} p.{x.page + 1} -- {x.description}"
                     + (f" | text: {snippet}" if snippet else ""))
    if not exhibits:
        lines.append("(none)")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Call + validation
# ---------------------------------------------------------------------------


def call_model(client: Any, model: str, user: str, *, ledger: Any = None, label: str = "",
               system: str = SYSTEM, max_tokens: int = _MAX_TOKENS) -> tuple[dict, str]:
    """(parsed JSON, error). Never raises: a failed call means "keep
    everything" for that tab, reported -- the PY copy is still valid."""
    try:
        response = client.messages.create(
            model=model, max_tokens=max_tokens, system=system,
            messages=[{"role": "user", "content": user}],
        )
        if ledger is not None and getattr(response, "usage", None) is not None:
            try:
                ledger.record(response.usage, group="Mirror", label=label)
            except Exception:  # noqa: BLE001 -- observation only
                pass
        raw = "".join(b.text for b in response.content if getattr(b, "type", "") == "text").strip()
        start, end = raw.find("{"), raw.rfind("}")
        if start == -1 or end == -1:
            return {}, "the model did not return JSON"
        return json.loads(raw[start:end + 1]), ""
    except Exception as exc:  # noqa: BLE001
        return {}, f"{type(exc).__name__}: {exc}"[:200]


_TAG_RE = re.compile(r"<(s\d+)>(.*?)</\1>", re.S)


def parse_tagged(paragraph: str, valid: set[str], default: str) -> list[tuple[str, str]]:
    """'<s0>Test Step 1: </s0><s1>IA noted...</s1>' -> [(s0, ...), (s1, ...)].
    Text outside tags takes the default style; unknown tags take the default."""
    paragraph = paragraph.replace("<br/>", "\n").replace("<br>", "\n")
    runs: list[tuple[str, str]] = []
    pos = 0
    for m in _TAG_RE.finditer(paragraph):
        if m.start() > pos and paragraph[pos:m.start()]:
            runs.append((default, paragraph[pos:m.start()]))
        runs.append((m.group(1) if m.group(1) in valid else default, m.group(2)))
        pos = m.end()
    if pos < len(paragraph):
        runs.append((default, re.sub(r"</?s\d+>", "", paragraph[pos:])))
    return [(k, t) for k, t in runs if t != ""]


ACTIONS = ("keep", "replace", "remove", "place", "connect")


def validate(raw: dict, tab: TabInventory, exhibit_ids: set[str], *, error: str = "") -> TabPlan:
    plan = TabPlan()
    if error:
        plan.notes.append(f"the edit call for this tab failed ({error}); the PY tab was left unchanged")
    ids = {i.id for i in tab.items if i.kind not in ("frame", "picture")}
    for entry in raw.get("elements", []) or []:
        if not isinstance(entry, dict):
            continue
        eid, action = str(entry.get("id", "")), str(entry.get("action", "keep")).lower()
        if eid not in ids or action not in ACTIONS:
            continue
        clean: dict = {"action": action}
        if isinstance(entry.get("paragraphs"), list):
            clean["paragraphs"] = [str(p) for p in entry["paragraphs"]]
        if action == "replace" and "paragraphs" not in clean:
            continue
        if action == "place":
            on = entry.get("on") or {}
            if on.get("exhibit") in exhibit_ids and str(on.get("anchor", "")).strip():
                clean["on"] = {"exhibit": on["exhibit"], "anchor": str(on["anchor"]).strip()}
            elif str(on.get("heading", "")).strip():
                clean["on"] = {"heading": str(on["heading"]).strip()}
            elif str(on.get("beside", "")) in ids:
                clean["on"] = {"beside": str(on["beside"])}
            else:
                plan.notes.append(f"element {eid}: 'place' with no usable target; kept where it was")
                clean = {"action": "keep"}
        if action == "connect":
            if str(entry.get("from")) in ids and str(entry.get("to")) in ids:
                clean.update({"from": str(entry["from"]), "to": str(entry["to"])})
            else:
                clean = {"action": "keep"}
        plan.elements[eid] = clean
    plan.unaddressed = sorted(ids - set(plan.elements))
    for missing in plan.unaddressed:
        plan.elements[missing] = {"action": "keep"}

    pic_ids = {i.id for i in tab.by_kind("picture")}
    for entry in raw.get("pictures", []) or []:
        if isinstance(entry, dict) and str(entry.get("id")) in pic_ids:
            x = entry.get("exhibit")
            crop = entry.get("crop") if isinstance(entry.get("crop"), dict) else None
            plan.pictures[str(entry["id"])] = {"exhibit": x if x in exhibit_ids else None, "crop": crop}
    for missing in sorted(pic_ids - set(plan.pictures)):
        plan.unaddressed.append(missing)  # a PY picture the model did not decide on: kept, flagged
    used = {p["exhibit"] for p in plan.pictures.values() if p["exhibit"]}
    plan.extra_exhibits = [x for x in raw.get("extra_exhibits", []) or [] if x in exhibit_ids and x not in used]

    for entry in raw.get("cells", []) or []:
        if isinstance(entry, dict) and re.fullmatch(r"[A-Z]+\d+", str(entry.get("ref", ""))):
            ref = str(entry["ref"])
            cell = tab.cells.get(ref)
            if cell is not None and cell.formula:
                continue  # formulas are never edited
            action = str(entry.get("action", "keep")).lower()
            if action in ("replace", "clear"):
                plan.cells[ref] = {"action": action, "value": entry.get("value")}

    for entry in raw.get("support", []) or []:
        if isinstance(entry, dict) and entry.get("py_document"):
            match = str(entry.get("match", "")).lower()
            plan.support.append({
                "py_ref": str(entry.get("py_ref", "")),
                "py_document": str(entry.get("py_document", ""))[:160],
                "cy_exhibit": entry.get("cy_exhibit") if entry.get("cy_exhibit") in exhibit_ids else None,
                "match": match if match in ("same", "partial", "missing") else ("missing" if not entry.get("cy_exhibit") else "partial"),
                "comment": str(entry.get("comment", ""))[:200],
            })
    plan.notes += [str(n) for n in raw.get("notes", []) or [] if str(n).strip()]
    return plan
