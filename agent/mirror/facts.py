"""CY facts, grouped the way the PY is: by SELECTION, every test step on it.

The previous renderer made one tab per selection PER TEST STEP, so a
three-step control with one selection came out as three near-identical
tabs, where the PY has one tab carrying all three steps' conclusions.
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any

from agent.mirror.excel_exhibits import EXCEL_SUFFIXES, make_exhibits
from agent.mirror.rewrite import Exhibit, SelectionFacts
from agent.schemas import ConclusionOutput, SamplePopulationManifest

_IMAGE_EXT = (".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif")
_PAGE_RE = re.compile(r"\bp(?:age)?\.?\s*(\d+)", re.I)


def _page(location: str) -> int:
    m = _PAGE_RE.search(location or "")
    return max(0, int(m.group(1)) - 1) if m else 0


def _page_text(path: Path, page: int) -> str:
    if path.suffix.lower() != ".pdf" or not path.exists():
        return ""
    try:
        import pdfplumber

        with pdfplumber.open(str(path)) as pdf:
            return (pdf.pages[page].extract_text() or "") if page < len(pdf.pages) else ""
    except Exception:  # noqa: BLE001 -- text is a matching aid, not required
        return ""


def selection_facts(
    results: dict[str, dict],
    spec: dict[str, Any],
    manifests: dict[str, SamplePopulationManifest] | None,
    support_dir: Path,
) -> list[SelectionFacts]:
    """One SelectionFacts per distinct sampled item, in first-seen order."""
    step_text = {s.get("test_step_id"): s.get("test_step_text", "") for s in spec.get("test_steps", [])}
    order: list[str] = []
    rows: dict[str, dict] = {}
    for manifest in (manifests or {}).values():
        for s in manifest.samples:
            if s.sample_id not in rows:
                order.append(s.sample_id)
                rows[s.sample_id] = dict(s.key_fields or {})

    conclusions: list[tuple[str, ConclusionOutput]] = [
        (sid, r["conclusion"]) for sid, r in results.items() if "conclusion" in r
    ]
    for _step, c in conclusions:
        for sr in c.sample_results:
            if sr.sample_id not in rows:
                order.append(sr.sample_id)
                rows[sr.sample_id] = {}
    if not order:
        order = ["1"]
        rows["1"] = {}

    out: list[SelectionFacts] = []
    for sid in order:
        steps: list[dict[str, Any]] = []
        exhibits: dict[tuple[str, int], Exhibit] = {}
        for step_id, c in conclusions:
            verdict = next((sr for sr in c.sample_results if sr.sample_id == sid), None)
            single = len({sr.sample_id for sr in c.sample_results}) <= 1
            if verdict is None and not single:
                continue
            attrs = [a for a in c.attribute_results if a.sample_id in (sid, None) or single]
            cited = {e for a in attrs for e in a.evidence_ids}
            citations = [e for e in c.evidence_citations
                         if e.sample_id in (sid, None) or e.evidence_id in cited or single]
            key_of: dict[str, tuple[str, int]] = {}   # evidence_id -> exhibit key
            excel: dict[str, list] = {}
            for e in citations:
                if e.source_file.lower().endswith(EXCEL_SUFFIXES):
                    excel.setdefault(e.source_file, []).append(e)
                    continue
                if not e.source_file.lower().endswith(_IMAGE_EXT):
                    continue
                key = (e.source_file, _page(e.location))
                key_of[e.evidence_id] = key
                if key not in exhibits:
                    exhibits[key] = Exhibit(
                        id=f"X{len(exhibits) + 1}", source_file=e.source_file, page=key[1],
                        description=e.quote_or_summary[:240],
                        text=_page_text(support_dir / e.source_file, key[1])[:1500],
                    )
            # Evidence read from a workbook becomes a picture too -- before
            # this, Excel support never reached the workpaper and the PY's
            # pictures were left in its place.
            for name, cites in excel.items():
                try:
                    made = make_exhibits(support_dir / name, [c.location for c in cites], support_dir)
                except Exception:  # noqa: BLE001 -- an unreadable workbook just has no picture
                    continue
                for picture, text, locs in made:
                    key = (picture, 0)
                    shown = [cite for cite in cites if cite.location in locs]
                    if key not in exhibits:
                        exhibits[key] = Exhibit(
                            id=f"X{len(exhibits) + 1}", source_file=picture, page=0,
                            description=("; ".join(cite.quote_or_summary for cite in shown) or name)[:240],
                            text=text[:1500],
                        )
                    for cite in shown:
                        key_of[cite.evidence_id] = key
            steps.append({
                "test_step": step_id,
                "test_step_text": step_text.get(step_id, ""),
                "result": (verdict.conclusion if verdict else c.conclusion),
                "result_note": verdict.note if verdict else "",
                "attributes": [
                    {"attribute": a.attribute, "result": a.result, "observed": a.value_observed,
                     "exhibits": sorted({exhibits[key_of[eid]].id for eid in a.evidence_ids
                                         if eid in key_of and key_of[eid] in exhibits})}
                    for a in attrs
                ],
                "exceptions": [x for x in c.exceptions if single or sid in x or f"row_{sid}" in x],
                "testing_agent_summary": c.narrative,
            })
        out.append(SelectionFacts(sid, rows.get(sid, {}), steps, list(exhibits.values())))

    # Every uploaded support page, not only the ones the testing agent cited:
    # "does the CY support mirror the PY support" cannot be answered from a
    # subset. Uncited pages are offered to every selection, marked as such;
    # the editor decides where (if anywhere) they belong.
    for f in out:
        have = {(x.source_file, x.page) for x in f.exhibits}
        for name in spec.get("cy_support_files", []) or []:
            path = support_dir / name
            if not name.lower().endswith(_IMAGE_EXT) or not path.exists():
                continue
            for page in range(_page_count(path))[:10]:
                if (name, page) in have or len(f.exhibits) >= 30:
                    continue
                f.exhibits.append(Exhibit(
                    id=f"X{len(f.exhibits) + 1}", source_file=name, page=page,
                    description="(not cited by the testing agent" + (")" if len(out) == 1 else
                                                                     "; may belong to another selection)"),
                    text=_page_text(path, page)[:1500]))
                have.add((name, page))
    return out


def _page_count(path: Path) -> int:
    if path.suffix.lower() != ".pdf":
        return 1
    try:
        import pdfplumber

        with pdfplumber.open(str(path)) as pdf:
            return len(pdf.pages)
    except Exception:  # noqa: BLE001
        return 1


def exhibit_path(support_dir: Path, exhibit: Exhibit) -> str:
    return os.path.join(str(support_dir), exhibit.source_file)
