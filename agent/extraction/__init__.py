"""Deterministic extraction: raw files -> agent.schemas.EvidenceItem.

Per docs/cy_testing_agent_design.md section 2, this is a batch step that
runs once per upload, before any Claude call. Claude is never handed a PDF
or workbook directly -- only the EvidenceItem list this package produces.
"""

from __future__ import annotations

from pathlib import Path

from agent.schemas import EvidenceItem

from .excel import extract_excel
from .pdf import extract_pdf

_HANDLERS = {
    ".xlsx": extract_excel,
    ".xlsm": extract_excel,
    ".xls": extract_excel,
    ".pdf": extract_pdf,
}


def extract(path: str | Path) -> list[EvidenceItem]:
    """Dispatch to the right extractor by file extension."""
    path = Path(path)
    handler = _HANDLERS.get(path.suffix.lower())
    if handler is None:
        raise ValueError(
            f"no extractor registered for {path.suffix!r} (file: {path.name}). "
            f"Supported: {sorted(_HANDLERS)}"
        )
    return handler(path)


def extract_many(paths: list[str | Path]) -> list[EvidenceItem]:
    """Extract multiple files and merge into one evidence pool.

    A control's CY support is normally more than one file (an Excel recon
    plus a PDF GL export, say), and each single-file extractor numbers
    evidence_id starting fresh at ev_0001 -- calling extract() on two files
    and concatenating the results would silently collide IDs. This
    renumbers globally across the merged set instead, so every evidence_id
    handed to the tool loop is actually unique for the run.
    """
    items: list[EvidenceItem] = []
    for path in paths:
        items.extend(extract(path))

    renumbered = [item.model_copy(update={"evidence_id": f"ev_{i:04d}"}) for i, item in enumerate(items, start=1)]
    return renumbered


__all__ = ["extract", "extract_many", "extract_excel", "extract_pdf"]
