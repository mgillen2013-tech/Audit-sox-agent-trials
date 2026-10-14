"""The one judgment step: "rewrite last year's item for this year".

The model is shown each PY item exactly as written -- every text box with
its formatting marked, every non-data cell, every exhibit slot and what it
held, every tickmark and what it pointed at -- next to the CY facts the
testing agent established. It returns, per item: keep it, replace it, or
remove it. Replacements must read like the PY: same sentences, same order,
same labels and line breaks, same length within reason. Only facts change.

What it is NOT asked to do: decide where anything goes, how anything looks,
or whether the control passed. Positions and formats come from the PY;
pass/fail comes from the testing agent's conclusions, which it must report
faithfully.

Output is strict JSON, validated here. Anything missing or malformed
degrades to "keep" for that item, with a note -- never to a guess.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Any

from agent.mirror.inventory import DrawingItem, TabInventory

_MAX_TOKENS = 12_000

SYSTEM = """\
You update last year's (PY) SOX testing workpaper for this year (CY). You are
an editor, not an author: the PY wording is the house style and must be kept.

For every PY item you are given, decide:
  keep     the item is still true for CY exactly as written
  replace  rewrite it for CY
  remove   it has no CY counterpart (e.g. a note about a PY-only issue)

RULES FOR REPLACEMENTS -- these are strict:
1. Same structure. Same number of paragraphs, same order, same labels
   ("Test Step 1:", "TM 2", "Note:"), same blank lines, same <br/> line breaks.
   Change only the facts:
   names, dates, amounts, document numbers, counts, and the result wording
   where the CY result differs.
2. Same style tags. Text is written as <s0>..</s0><s1>..</s1>; each tag is a
   format taken from the PY (bold, red, etc.). Use the same tag for the same
   role as the PY did. Never invent a tag number that was not shown.
3. Every CY fact must come from the CY FACTS below. Never carry a PY fact
   (a PY name, date, amount, document number) into CY. If CY evidence does
   not show something the PY sentence states, say what CY shows instead --
   in the PY's own phrasing for that situation.
4. The result must match the testing agent's conclusion for that test step
   and selection. Where the PY said "Test Satisfied." and CY is an exception,
   use the PY's exception phrasing if the PY has one, else the closest
   phrasing in the same style ("Exception noted.").
5. Keep length within about 25% of the PY item. No new sections.
6. Tickmark references in text ("See TM2") must refer to tickmarks you place.

EXHIBITS: each PY picture slot held a document. Assign the CY exhibit that
plays the same role (invoice to the invoice slot, approval to the approval
slot). A slot with no CY counterpart gets null. CY exhibits with no PY slot
go in extra_exhibits.

TICKMARKS: each PY tickmark label marked one value. Place it on the CY
exhibit and the exact CY text that plays the same role, as it appears in the
document (copy it character for character; one value, not a sentence). Or
put it on the selection's data row by heading. Remove it if CY has no
counterpart.

CELLS: cells other than the headings and the selection's data row (e.g. an
IA recalculation). Formula cells are shown for context and are never edited.

Reply with JSON only:
{"text": [{"id": "...", "action": "keep|replace|remove", "paragraphs": ["<s0>...</s0>", ""]}],
 "cells": [{"ref": "AD9", "action": "keep|replace|clear", "value": "..."}],
 "pictures": [{"id": "...", "exhibit": "X1" or null}],
 "extra_exhibits": ["X4"],
 "tickmarks": [{"id": "...", "action": "place|remove", "exhibit": "X1", "anchor": "...",
                "data_row_heading": "..."}],
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
    text: dict[str, dict] = field(default_factory=dict)
    cells: dict[str, dict] = field(default_factory=dict)
    pictures: dict[str, str | None] = field(default_factory=dict)
    extra_exhibits: list[str] = field(default_factory=list)
    tickmarks: dict[str, dict] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Prompt
# ---------------------------------------------------------------------------


def _tagged(item: DrawingItem) -> list[str]:
    out = []
    for p in item.paragraphs:
        # A line break INSIDE a paragraph is <br/>, never a new line here:
        # shown as a new line it read as an extra paragraph, and a faithful
        # copy was then flagged as having changed the structure.
        out.append("".join(
            "<br/>" if k == "br" else f"<{k}>{t}</{k}>" for k, t in p.runs
        ))
    return out


def _style_legend(tab: TabInventory) -> str:
    parts = []
    for s in tab.styles.values():
        bits = ["bold" if s.bold else "regular"]
        if s.color and s.color.upper() not in ("000000",):
            bits.append(f"colour {s.color}")
        size = re.search(r'\bsz="(\d+)"', s.rpr_xml)
        if size:
            bits.append(f"{int(size.group(1)) / 100:g}pt")
        parts.append(f"{s.key} = {', '.join(bits)}")
    return "; ".join(parts) or "(none)"


def tab_prompt(
    tab: TabInventory,
    facts: SelectionFacts,
    *,
    picture_notes: dict[str, str],
    tickmark_notes: dict[str, str],
    data_rows: set[int],
    header_row: int,
) -> str:
    """The user message for one CY selection tab, built from its PY tab."""
    lines = [f"PY TAB {tab.name!r} -> CY selection {facts.sample_id!r}", ""]
    lines.append(f"STYLE TAGS: {_style_legend(tab)}")
    lines.append("")
    lines.append("PY TEXT ITEMS:")
    for item in tab.by_kind("text"):
        lines.append(f"- id {item.id} ({item.name}):")
        for p in _tagged(item):
            lines.append(f"    | {p}")
    lines.append("")
    lines.append("PY TICKMARKS (label -> what it marked in the PY):")
    for item in tab.by_kind("label"):
        lines.append(f"- id {item.id}: {item.text!r} -> {tickmark_notes.get(item.id, 'unknown')}")
    lines.append("")
    lines.append("PY PICTURE SLOTS (what each held in the PY):")
    for item in tab.by_kind("picture"):
        lines.append(f"- id {item.id}: {picture_notes.get(item.id, '(unreadable)')}")
    cells = [
        c for ref, c in tab.cells.items()
        if (m := re.fullmatch(r"[A-Z]+(\d+)", ref)) and int(m.group(1)) not in data_rows | {header_row}
    ]
    if cells:
        lines.append("")
        lines.append("PY CELLS (other than headings and data row):")
        for c in sorted(cells, key=lambda c: (int(re.sub(r"\D", "", c.ref)), c.ref))[:80]:
            lines.append(f"- {c.ref}: {c.text!r}" + (" [formula, read-only]" if c.formula else ""))

    lines += ["", "CY FACTS (the only source of CY content):"]
    lines.append(f"Selection {facts.sample_id} data row: " + json.dumps(facts.data_row, default=str))
    for s in facts.steps:
        lines.append(json.dumps(s, default=str))
    lines.append("")
    lines.append("CY EXHIBITS:")
    for x in facts.exhibits:
        snippet = " ".join(x.text.split())[:400]
        lines.append(f"- {x.id}: {x.source_file} p.{x.page + 1} -- {x.description}" + (f" | text: {snippet}" if snippet else ""))
    return "\n".join(lines)


def lead_prompt(tab: TabInventory, steps: list[dict[str, Any]], cy_tab_names: list[str]) -> str:
    lines = [f"PY TAB {tab.name!r} (the control's lead sheet)", "",
             f"STYLE TAGS: {_style_legend(tab)}", "", "PY TEXT ITEMS:"]
    for item in tab.by_kind("text"):
        lines.append(f"- id {item.id} ({item.name}):")
        for p in _tagged(item):
            lines.append(f"    | {p}")
    lines += ["", "CY FACTS:", f"CY selection tabs: {cy_tab_names}"]
    for s in steps:
        lines.append(json.dumps(s, default=str))
    lines += ["", "Keep the control description and test steps unless CY facts show they "
                  "changed. Update tab references to the CY tabs. No pictures, cells or "
                  "tickmarks on this tab: return empty lists for those."]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Call + validation
# ---------------------------------------------------------------------------


def call_model(client: Any, model: str, user: str, *, ledger: Any = None, label: str = "") -> tuple[dict, str]:
    """(parsed JSON, error). Never raises: a failed call means "keep
    everything" for that tab, reported -- the PY copy is still valid."""
    try:
        response = client.messages.create(
            model=model, max_tokens=_MAX_TOKENS, system=SYSTEM,
            messages=[{"role": "user", "content": user}],
        )
        if ledger is not None:
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


def validate(
    raw: dict, tab: TabInventory, exhibit_ids: set[str], *, error: str = ""
) -> TabPlan:
    plan = TabPlan()
    if error:
        plan.notes.append(f"{tab.name}: the rewrite call failed ({error}); PY text left unchanged for review")
    text_ids = {i.id for i in tab.by_kind("text")}
    for entry in raw.get("text", []) or []:
        if not isinstance(entry, dict) or str(entry.get("id")) not in text_ids:
            continue
        action = str(entry.get("action", "keep")).lower()
        if action == "replace" and isinstance(entry.get("paragraphs"), list):
            plan.text[str(entry["id"])] = {"action": "replace", "paragraphs": [str(p) for p in entry["paragraphs"]]}
        elif action == "remove":
            plan.text[str(entry["id"])] = {"action": "remove"}
    for missing in sorted(text_ids - set(plan.text)):
        plan.text.setdefault(missing, {"action": "keep"})

    for entry in raw.get("cells", []) or []:
        if isinstance(entry, dict) and re.fullmatch(r"[A-Z]+\d+", str(entry.get("ref", ""))):
            ref = str(entry["ref"])
            cell = tab.cells.get(ref)
            if cell is not None and cell.formula:
                continue  # formulas are never edited
            action = str(entry.get("action", "keep")).lower()
            if action in ("replace", "clear"):
                plan.cells[ref] = {"action": action, "value": entry.get("value")}

    pic_ids = {i.id for i in tab.by_kind("picture")}
    for entry in raw.get("pictures", []) or []:
        if isinstance(entry, dict) and str(entry.get("id")) in pic_ids:
            x = entry.get("exhibit")
            plan.pictures[str(entry["id"])] = x if x in exhibit_ids else None
    plan.extra_exhibits = [x for x in raw.get("extra_exhibits", []) or [] if x in exhibit_ids
                           and x not in plan.pictures.values()]

    label_ids = {i.id for i in tab.by_kind("label")}
    for entry in raw.get("tickmarks", []) or []:
        if not isinstance(entry, dict) or str(entry.get("id")) not in label_ids:
            continue
        if str(entry.get("action", "")).lower() == "place" and (
            (entry.get("exhibit") in exhibit_ids and str(entry.get("anchor", "")).strip())
            or str(entry.get("data_row_heading", "")).strip()
        ):
            plan.tickmarks[str(entry["id"])] = {
                "exhibit": entry.get("exhibit"),
                "anchor": str(entry.get("anchor", "")).strip(),
                "data_row_heading": str(entry.get("data_row_heading", "")).strip(),
            }
        else:
            plan.tickmarks[str(entry["id"])] = {"remove": True}
    for missing in sorted(label_ids - set(plan.tickmarks)):
        plan.tickmarks[missing] = {"remove": True}
    plan.notes += [str(n) for n in raw.get("notes", []) or [] if str(n).strip()]
    return plan
