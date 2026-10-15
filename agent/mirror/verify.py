"""Two questions, answered on the Mirror Check tab:

  1. Does the CY workpaper mirror the PY?
  2. Does the CY support mirror the PY support?

Plain-code checks, no rules about what a workpaper contains:
  structure   every PY text element accounted for; where the PY set a
              paragraph's lead-in apart by FORMAT (a bold "Test Step 1: "
              before regular text), CY keeps the same lead-in
  leftovers   the safety net for everything else: every specific fact in
              the PY's working tabs -- numbers, dates in any format, names --
              is looked for in EVERY cell and shape of the CY workbook. Found
              there and not a CY fact = last year's detail left behind,
              wherever it is and however it got there.
  placement   anything that could not be placed, or was not decided on
  support     each PY document matched to a CY one, or missing (from the
              model's comparison, reported as it said)

And a final model pass comparing PY and CY text item by item.
"""

from __future__ import annotations

import datetime as dt
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
    found_by: str  # structure | leftovers | placement | support | PY data | model


def _norm(s: str) -> str:
    return re.sub(r"[\s,$]", "", str(s)).lower()


def structural_checks(tab: str, changes: list[tuple]) -> list[CheckFlag]:
    flags = []
    for change in changes:
        item, py, cy = change[0], change[1], change[2]
        py_labels = change[3] if len(change) > 3 else []
        cy_labels = change[4] if len(change) > 4 else []
        if cy == "" and py:
            flags.append(CheckFlag("medium", tab, item, "PY text element removed (no CY counterpart)", "structure"))
            continue
        if py == cy:
            continue
        n_py, n_cy = len(py.split("\n")), len(cy.split("\n"))
        if n_py != n_cy:
            flags.append(CheckFlag("medium", tab, item, f"paragraph count changed ({n_py} in PY, {n_cy} in CY)", "structure"))
        shape = lambda xs: [re.sub(r"\d+", "#", x.lower()) for x in xs]  # noqa: E731
        if py_labels and shape(py_labels) != shape(cy_labels):
            flags.append(CheckFlag("medium", tab, item, f"PY lead-ins {py_labels} not kept (CY has {cy_labels})", "structure"))
        ratio = len(cy) / max(1, len(py))
        if ratio > 1.6 or ratio < 0.5:
            flags.append(CheckFlag("low", tab, item, f"length changed a lot ({len(py)} -> {len(cy)} characters)", "structure"))
    return flags


# ---------------------------------------------------------------------------
# Leftovers: the safety net
# ---------------------------------------------------------------------------

_MONTHS = {m: i + 1 for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}
_DATE_RES = [
    (re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{2,4})\b"), "mdy"),
    (re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b"), "ymd"),
    (re.compile(r"\b([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})\b"), "Mdy"),
    (re.compile(r"\b(\d{1,2})[- ]([A-Za-z]{3,9})[- ](\d{2,4})\b"), "dMy"),
]
_NUM_RE = re.compile(r"(?<![\w.])\$?\d[\d,]*(?:\.\d+)?(?![\w])")


def _dates(text: str) -> set[str]:
    """Every date in the text, normalised to YYYY-MM-DD, whatever its format."""
    out = set()
    for rx, kind in _DATE_RES:
        for m in rx.finditer(text):
            try:
                if kind == "mdy":
                    mo, d, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                elif kind == "ymd":
                    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                elif kind == "Mdy":
                    mo, d, y = _MONTHS[m.group(1)[:3].lower()], int(m.group(2)), int(m.group(3))
                else:
                    d, mo, y = int(m.group(1)), _MONTHS[m.group(2)[:3].lower()], int(m.group(3))
                y = y + 2000 if y < 100 else y
                out.add(dt.date(y, mo, d).isoformat())
            except (KeyError, ValueError):
                continue
    return out


def _numbers(text: str) -> set[str]:
    """Specific numbers: four or more digits, or with decimals. Dates are
    handled separately; these are document numbers and amounts."""
    out = set()
    for m in _NUM_RE.finditer(text):
        n = m.group(0).replace("$", "").replace(",", "")
        digits = re.sub(r"\D", "", n)
        if len(digits) >= 4 or ("." in n and len(digits) >= 3):
            out.add(n.rstrip("0").rstrip(".") if "." in n else n.lstrip("0") or "0")
    return out


def _names(text: str) -> set[str]:
    """Capitalised words and phrases -- candidates only; py_facts keeps the
    ones that appear in the PY's evidence."""
    out = set()
    for m in re.finditer(r"\b([A-Z][a-zA-Z&'.-]{2,}(?:\s+[A-Z][a-zA-Z&'.-]{1,}){0,3})\b", text):
        out.add(m.group(1).strip())
    return out


@dataclass
class PyFact:
    kind: str     # date | number | name
    value: str    # as written in the PY
    key: str      # normalised for matching


def py_facts(texts: list[str], data_values: list[str], evidence_text: str = "") -> tuple[list[PyFact], list[PyFact]]:
    """The PY's specific facts, from its working tabs' text and data rows.

    A capitalised phrase in PY text counts as a fact only if it also appears
    in the PY's own EVIDENCE (its data rows, the text of its screenshots):
    "Ryan Clark" is on last year's approval email, so it is a fact; "Test
    Step" and "Account Manager" are on no document, so they are wording.
    No list of words decides it.

    Returns (certain facts, name candidates). Numbers, dates and the values
    of the PY's own record rows are certain. Other names are only
    CANDIDATES: whether "Project Manager" is a person or a job title is a
    judgment, so the model check decides, not code."""
    evidence = _norm(evidence_text + " " + " ".join(map(str, data_values)))
    facts: dict[str, PyFact] = {}
    for t in texts + data_values:
        for d in _dates(t):
            facts.setdefault("d" + d, PyFact("date", d, d))
        for n in _numbers(t):
            facts.setdefault("n" + n, PyFact("number", n, n))
    for v in data_values:
        v = str(v).strip()
        if re.search(r"[A-Za-z]{3}", v) and len(v) >= 4 and not _dates(v):
            facts.setdefault("t" + _norm(v), PyFact("name", v, _norm(v)))
    candidates: dict[str, PyFact] = {}
    for t in texts:
        for name in _names(t):
            key = _norm(name)
            if len(key) >= 4 and key in evidence and "t" + key not in facts:
                candidates.setdefault(key, PyFact("name", name, key))
    return list(facts.values()), list(candidates.values())


def name_hits(candidates: list[PyFact], cy_locations: list[tuple[str, str, str, bool]],
              cy_fact_text: str) -> dict[str, list[str]]:
    """Per tab, the PY name candidates that appear in CY text and not in the
    CY facts -- for the model check to judge."""
    cy_norm = _norm(cy_fact_text)
    out: dict[str, list[str]] = {}
    for tab, _where, text, _kept in cy_locations:
        norm = _norm(text)
        for c in candidates:
            if c.key in norm and c.key not in cy_norm and c.value not in out.get(tab, []):
                out.setdefault(tab, []).append(c.value)
    return out


def leftover_scan(
    facts: list[PyFact],
    cy_locations: list[tuple[str, str, str, bool]],   # (tab, where, text, kept-as-is)
    cy_fact_text: str,
    common_text: str,
) -> list[CheckFlag]:
    """Every PY fact found in the CY workbook that is not a CY fact.

    common_text: text the PY and CY share by design (headings, the control
    description, test step wording) -- a 'name' found there is wording, not
    a leftover. Numbers and dates are never excused this way.
    """
    cy_dates = _dates(cy_fact_text)
    cy_nums = _numbers(cy_fact_text)
    cy_norm = _norm(cy_fact_text)
    common = _norm(common_text)
    flags: list[CheckFlag] = []
    seen: set[tuple[str, str]] = set()
    for tab, where, text, kept in cy_locations:
        if not text:
            continue
        dates, nums, norm = _dates(text), _numbers(text), _norm(text)
        for f in facts:
            if f.kind == "date":
                hit = f.key in dates and f.key not in cy_dates
            elif f.kind == "number":
                hit = f.key in nums and f.key not in cy_nums
            else:
                hit = len(f.key) >= 4 and f.key in norm and f.key not in cy_norm and f.key not in common
            if hit and (tab, f.key) not in seen:
                seen.add((tab, f.key))
                flags.append(CheckFlag(
                    "medium" if kept else "high", tab, where,
                    f"PY {f.kind} {f.value!r} is in the CY workpaper but not in the CY facts"
                    + (" (in text the editor kept as still true -- confirm)" if kept else ""),
                    "leftovers"))
    return flags


def support_flags(tab: str, support: list[dict]) -> list[CheckFlag]:
    out = []
    for s in support:
        if s["match"] == "missing":
            out.append(CheckFlag("high", tab, s.get("py_ref") or "-",
                                 f"PY had {s['py_document']!r}; no CY counterpart in the support"
                                 + (f" -- {s['comment']}" if s.get("comment") else ""), "support"))
        elif s["match"] == "partial":
            out.append(CheckFlag("medium", tab, s.get("py_ref") or "-",
                                 f"PY {s['py_document']!r} only partly matched by CY {s.get('cy_exhibit')}"
                                 + (f" -- {s['comment']}" if s.get("comment") else ""), "support"))
    return out


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


def model_check(client: Any, model: str, tab: str, changes: list[tuple],
                cy_facts_text: str, *, ledger: Any = None, name_candidates: list[str] | None = None,
                tickmarks: list[str] | None = None) -> list[CheckFlag]:
    pairs = [{"item": c[0], "py": c[1], "cy": c[2]} for c in changes if c[1] != c[2] and c[2]]
    names = name_candidates or []
    if (not pairs and not names and not tickmarks) or client is None:
        return [CheckFlag("low", tab, "-", f"possible PY name {n!r} in CY text -- check", "leftovers")
                for n in names]
    from agent.mirror.rewrite import call_model

    user = (f"TAB {tab}\n\nCY FACTS:\n{cy_facts_text[:12000]}\n\nITEMS:\n"
            + json.dumps(pairs, ensure_ascii=False, indent=1)[:30000])
    if tickmarks:
        user += ("\n\nTICKMARKS (label -> the CY value it was placed on -> the text citing it):\n"
                 + "\n".join(tickmarks)
                 + "\nFlag (high) any tickmark whose placed value is not what the citing text says it marks.")
    if names:
        user += ("\n\nPY NAMES FOUND IN THIS TAB'S CY TEXT (and not in the CY facts): "
                 + json.dumps(names) + "\nFlag (high) any that is a PY-specific person, company or place "
                 "left over from last year; ignore ordinary words and job titles.")
    data, err = call_model(client, model, user, ledger=ledger, label=f"check {tab}",
                           system=VERIFY_SYSTEM, max_tokens=4000)
    if err:
        return [CheckFlag("medium", tab, "-", f"the model check could not run ({err[:80]})", "model")]
    out = []
    for f in data.get("flags", []) or []:
        sev = str(f.get("severity", "medium")).lower()
        out.append(CheckFlag(sev if sev in SEVERITIES else "medium", tab, str(f.get("item", "-")),
                             str(f.get("issue", ""))[:200], "model"))
    return out


def sort_flags(flags: list[CheckFlag]) -> list[CheckFlag]:
    return sorted(flags, key=lambda f: (SEVERITIES.index(f.severity), f.found_by, f.tab, f.item))


# ---------------------------------------------------------------------------
# Tickmarks: every one cited in text must be on the evidence, and vice versa
# ---------------------------------------------------------------------------


def _cites(text: str, label: str) -> bool:
    """`label` appears in `text` as its own token (not inside a longer
    word or number): "TM 2" is cited by "See TM 2.", not by "TM 22"."""
    if not text or not label:
        return False
    # Case-sensitive: marks are written with fixed capitals ("TM2", "A"), and
    # ignoring case made ordinary words cite them ("a credit memo", "total
    # amount" for a "Total Amount" annotation).
    pattern = r"(?<![\w])" + r"\s*".join(re.escape(p) for p in label.split()) + r"(?![\w])"
    return re.search(pattern, text) is not None


def py_cited_labels(py_tabs: list) -> set[str]:
    """The PY's referenced marks, read from the PY itself: a short text
    shape whose exact text is cited in some OTHER text on the PY's tabs
    (a conclusion saying "See TM2", a note saying "per A"). Whatever the
    PY's convention -- "TM 2", "A", "^" -- this finds it; no list of styles."""
    texts = []
    shapes = []
    for t in py_tabs:
        texts += [(t.name, i.id, i.text) for i in t.items if i.kind == "text" and i.text]
        texts += [(t.name, ref, c.text) for ref, c in t.cells.items() if c.text]
        shapes += [(t.name, i.id, i.text.strip()) for i in t.items if i.kind == "text" and 0 < len(i.text.strip()) <= 12]
    out = set()
    for tab, sid, label in shapes:
        # A citation is a mention INSIDE other text ("See TM2 and Note"), not
        # another place that just says the same words (a cell "Total Amount").
        if any(_cites(text, label) and len(text.strip()) > len(label) + 3
               for tab2, oid, text in texts if not (tab2 == tab and oid == sid)):
            out.add(label)
    return out


def tickmark_checks(
    selection: str,
    py_labels: set[str],
    declared: list[tuple[str, dict]],                 # (tab, plan tickmark entry)
    placed_labels: dict[str, list[tuple[str, str]]],  # label (normalised) -> [(tab, where it was placed)]
    cy_texts: list[tuple[str, str, str, bool]],       # (tab, where, text, is a label element)
) -> list[CheckFlag]:
    """Across all the tabs of one selection:
      - every tickmark CITED in CY text is PLACED on the CY evidence
      - every tickmark the model declared and that text cites is placed
      - a tickmark placed on the evidence that the PY cited is still cited
    """
    norm = lambda s: re.sub(r"\s+", " ", s.strip())  # noqa: E731
    labels = {norm(x) for x in py_labels} | {norm(t["label"]) for _tab, t in declared}
    flags: list[CheckFlag] = []
    cited_in_cy: dict[str, list[tuple[str, str]]] = {}
    for tab, where, text, is_label in cy_texts:
        if is_label:
            continue
        for lab in labels:
            if _cites(text, lab):
                cited_in_cy.setdefault(lab, []).append((tab, where))
    for lab, places in sorted(cited_in_cy.items()):
        if lab not in placed_labels:
            tabs = sorted({t for t, _w in places})
            flags.append(CheckFlag("high", ",".join(tabs), places[0][1],
                                   f"selection {selection}: text cites {lab!r} but no {lab!r} tickmark is placed on "
                                   "the CY evidence", "tickmarks"))
    for tab, t in declared:
        lab = norm(t["label"])
        if t.get("referenced_by") and lab not in placed_labels and lab not in cited_in_cy:
            flags.append(CheckFlag("high", tab, t.get("element") or "-",
                                   f"selection {selection}: tickmark {t['label']!r} ({t.get('marks', '')}) was "
                                   "declared and cited but is not on the CY evidence", "tickmarks"))
    for lab in sorted(placed_labels):
        if lab in {norm(x) for x in py_labels} and lab not in cited_in_cy:
            tab = placed_labels[lab][0][0]
            flags.append(CheckFlag("medium", tab, "-",
                                   f"selection {selection}: tickmark {lab!r} is on the evidence but no text cites it "
                                   "(the PY cited it)", "tickmarks"))
    return flags
