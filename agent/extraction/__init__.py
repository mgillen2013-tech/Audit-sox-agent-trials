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


__all__ = ["extract", "extract_excel", "extract_pdf"]
