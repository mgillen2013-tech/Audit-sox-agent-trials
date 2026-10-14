"""Does the CY workpaper mirror the PY? Checked two ways, both reported.

PLAIN-CODE CHECKS (free, certain):
  structure   every PY text box is accounted for; a replaced box kept the
              PY's paragraph count and its labels ("Test Step 1:", "Note:")
  leftovers   a PY fact -- a document number, date, amount or name from the
              PY's own row or text -- appearing in CY text when the CY facts
              do not contain it. The most damaging failure a mirrored
              workpaper can have, and one a reader cannot spot.
  placement   exhibits and tickmarks that could not be placed
  PY data     tabs still holding last year's data (population, parameters,
              report extracts) that this run had nothing to replace with

MODEL CHECK (the final call): the PY and CY version of every item side by
side, with the CY facts, asking one question per item -- same wording and
structure with only the facts changed, and are those facts supported?
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Any

SEVERITIES = ("high", "medium", "low")


@dataclass
class CheckFlag:
    severity: str
    tab: str
    item: str
    issue: str
    found_by: str  # "structure" | "leftovers" | "placement" | "PY data" | "model"


# Things that identify a specific transaction: long numbers, dates, money.
_FACT_RE = re.compile(
    r"\$?\d{1,3}(?:,\d{3})+(?:\.\d+)?|\$\d+(?:\.\d+)?|\b\d{5,}\b|"
    r"\b\d{1,2}/\d{1,2}/\d{2,4}\b|\b\d{4}-\d{2}-\d{2}\b"
)
_LABEL_RE = re.compile(r"^\s*((?:Test(?: Step)?\s*\d+|Note|TM\s*\d+|Procedures?|Control Description)\s*:?)", re.I)


def _norm(s: str) -> str:
    return re.sub(r"[\s,$]", "", s)


def structural_checks(tab: str, changes: list[tuple[str, str, str]]) -> list[CheckFlag]:
    flags = []
    for item, py, cy in changes:
        if cy == "":
            flags.append(CheckFlag("medium", tab, item, "PY text box removed (no CY counterpart)", "structure"))
            continue
        if py == cy:
            continue
        py_paras = [p for p in py.split("\n")]
        cy_paras = [p for p in cy.split("\n")]
        if len(py_paras) != len(cy_paras):
            flags.append(CheckFlag("medium", tab, item,
                                   f"paragraph count changed ({len(py_paras)} in PY, {len(cy_paras)} in CY)", "structure"))
        py_labels = [m.group(1).lower() for p in py_paras if (m := _LABEL_RE.match(p))]
        cy_labels = [m.group(1).lower() for p in cy_paras if (m := _LABEL_RE.match(p))]
        if py_labels and [re.sub(r"\d+", "#", x) for x in py_labels] != [re.sub(r"\d+", "#", x) for x in cy_labels]:
            flags.append(CheckFlag("medium", tab, item,
                                   f"PY labels {py_labels} not kept (CY has {cy_labels})", "structure"))
        ratio = len(cy) / max(1, len(py))
        if ratio > 1.6 or ratio < 0.5:
            flags.append(CheckFlag("low", tab, item, f"length changed a lot ({len(py)} -> {len(cy)} characters)", "structure"))
    return flags


def leftover_checks(tab: str, changes: list[tuple[str, str, str]], py_values: list[str],
                    cy_facts_text: str) -> list[CheckFlag]:
    """PY facts that survived into CY text without being CY facts too."""
    cy_facts = _norm(cy_facts_text)
    py_facts = {m.group(0) for _i, py, _c in changes for m in _FACT_RE.finditer(py)}
    py_facts |= {v for v in py_values if len(_norm(v)) >= 5 and re.search(r"\d", v)}
    names = {v for v in py_values if re.fullmatch(r"[A-Z][A-Za-z&.,' -]{5,}", v or "")}
    flags = []
    for item, _py, cy in changes:
        if not cy:
            continue
        cy_n = _norm(cy)
        for fact in sorted(py_facts | names):
            f = _norm(fact)
            if len(f) >= 4 and f in cy_n and f not in cy_facts:
                flags.append(CheckFlag("high", tab, item, f"PY fact {fact!r} appears in CY text but not in the CY facts", "leftovers"))
        # People and companies named in PY prose ("approval from Ryan Clark").
        # Less certain than a number -- role phrases are capitalised too --
        # so common workpaper words are excluded and the flag is medium.
        for name in _py_names(changes):
            n = _norm(name).lower()
            if n in cy_n.lower() and n not in cy_facts.lower():
                flags.append(CheckFlag("medium", tab, item, f"PY name {name!r} appears in CY text but not in the CY facts", "leftovers"))
    return flags


_COMMON = {
    "test", "step", "steps", "project", "manager", "branch", "administrator", "account", "customer",
    "invoice", "note", "total", "amount", "satisfied", "exception", "noted", "control", "description",
    "procedures", "procedure", "refer", "tab", "schedule", "value", "values", "approval", "approved",
    "the", "ia", "e1", "brightview", "purchase", "order", "proposal", "extra", "work", "designee",
    "payment", "date", "credit", "memo", "retainage", "sample", "selection", "reference", "references",
    "director", "vice", "president", "senior", "regional", "general", "finance", "accounts", "payable",
    "receivable", "contract", "billing", "sales", "service", "services", "landscape", "llc", "inc",
}


def _py_names(changes: list[tuple[str, str, str]]) -> set[str]:
    names = set()
    for _i, py, _c in changes:
        for m in re.finditer(r"\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})\b", py):
            words = m.group(1).split()
            if not all(w.lower() in _COMMON for w in words) and not any(w.lower() in _COMMON for w in words[:1]):
                names.add(m.group(1))
    return names


VERIFY_SYSTEM = """\
You are reviewing a CY SOX workpaper that was produced by editing last year's
(PY) workpaper. For each item you see the PY text, the CY text, and the CY facts.

The CY text must: (1) keep the PY's wording, sentence structure, labels and
order, changing only facts; (2) state only facts supported by the CY facts;
(3) carry no PY fact; (4) report the same result as the CY facts.

Flag only real problems. Reply with JSON only:
{"flags": [{"item": "...", "severity": "high|medium|low", "issue": "<= 25 words"}]}
An empty list means every item passes.
"""


def model_check(client: Any, model: str, tab: str, changes: list[tuple[str, str, str]],
                cy_facts_text: str, *, ledger: Any = None) -> list[CheckFlag]:
    pairs = [{"item": i, "py": py, "cy": cy} for i, py, cy in changes if py != cy and cy]
    if not pairs or client is None:
        return []
    user = (f"TAB {tab}\n\nCY FACTS:\n{cy_facts_text[:12000]}\n\nITEMS:\n"
            + json.dumps(pairs, ensure_ascii=False, indent=1)[:30000])
    try:
        response = client.messages.create(
            model=model, max_tokens=4000, system=VERIFY_SYSTEM,
            messages=[{"role": "user", "content": user}],
        )
        if ledger is not None:
            try:
                ledger.record(response.usage, group="Mirror check", label=tab)
            except Exception:  # noqa: BLE001
                pass
        raw = "".join(b.text for b in response.content if getattr(b, "type", "") == "text")
        start, end = raw.find("{"), raw.rfind("}")
        data = json.loads(raw[start:end + 1]) if start != -1 else {}
    except Exception as exc:  # noqa: BLE001
        return [CheckFlag("medium", tab, "-", f"the model check could not run ({str(exc)[:80]})", "model")]
    out = []
    for f in data.get("flags", []) or []:
        sev = str(f.get("severity", "medium")).lower()
        out.append(CheckFlag(sev if sev in SEVERITIES else "medium", tab, str(f.get("item", "-")),
                             str(f.get("issue", ""))[:200], "model"))
    return out


def sort_flags(flags: list[CheckFlag]) -> list[CheckFlag]:
    return sorted(flags, key=lambda f: (SEVERITIES.index(f.severity), f.tab, f.item))
