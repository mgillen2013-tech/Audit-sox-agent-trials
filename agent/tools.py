"""Deterministic tool implementations from design doc section 3.

These are plain functions over plain data -- no Claude, no network. The
per-test-step loop (agent/loop.py) is what wires these to actual tool_use
blocks and appends every call+result to the audit log; keeping that
concern out of this module is what makes these easy to unit test.
"""

from __future__ import annotations

import itertools
import re
from dataclasses import dataclass, field

from rank_bm25 import BM25Okapi

from agent.schemas import (
    CheckSampleCoverageError,
    CheckSampleCoverageInput,
    CheckSampleCoverageOutput,
    EvidenceItem,
    FlagExceptionInput,
    FlagExceptionOutput,
    RequestAdditionalSupportInput,
    RequestAdditionalSupportOutput,
    SampleCoverage,
    SearchCySupportInput,
    SearchCySupportOutput,
    SearchResult,
)

_TOKEN_RE = re.compile(r"[A-Za-z0-9]+")


def _tokenize(text: str) -> list[str]:
    return [t.lower() for t in _TOKEN_RE.findall(text)]


def _searchable_text(item: EvidenceItem) -> str:
    """The text BM25 indexes and the text a search snippet is built from."""
    if item.extracted_text:
        return item.extracted_text
    if item.extracted_table:
        return "\n".join(" | ".join(row) for row in item.extracted_table)
    if item.source_type == "image_ocr":
        return "[no extractable text -- likely a scanned page, OCR not run]"
    return ""


@dataclass
class ToolContext:
    """Per-control-run state the tools operate against and the audit-log
    bookkeeping the loop needs. One instance per test-step conversation.
    """

    evidence_items: list[EvidenceItem]
    exceptions: dict[str, FlagExceptionInput] = field(default_factory=dict)
    support_requests: dict[str, RequestAdditionalSupportInput] = field(default_factory=dict)
    evidence_ids_returned_by_search: set[str] = field(default_factory=set)

    # Populated/used by agent/loop.py -- kept here (not bolted on post-hoc)
    # so ToolContext is the single source of truth for per-run state.
    model: str = ""
    required_sample_ids: list[str] = field(default_factory=list)
    last_sample_coverage: "SampleCoverage | None" = None
    tool_call_count: int = 0

    _evidence_by_id: dict[str, EvidenceItem] = field(default_factory=dict, init=False)
    _bm25: BM25Okapi | None = field(default=None, init=False)
    _bm25_items: list[EvidenceItem] = field(default_factory=list, init=False)
    _exception_counter: "itertools.count" = field(default_factory=lambda: itertools.count(1), init=False)
    _request_counter: "itertools.count" = field(default_factory=lambda: itertools.count(1), init=False)

    def __post_init__(self) -> None:
        self._evidence_by_id = {item.evidence_id: item for item in self.evidence_items}
        indexable = [item for item in self.evidence_items if _searchable_text(item).strip()]
        self._bm25_items = indexable
        if indexable:
            corpus = [_tokenize(_searchable_text(item)) for item in indexable]
            self._bm25 = BM25Okapi(corpus)


def search_cy_support(inp: SearchCySupportInput, ctx: ToolContext) -> SearchCySupportOutput:
    candidates = ctx._bm25_items
    if inp.evidence_types:
        allowed = set(inp.evidence_types)
        candidates = [item for item in candidates if item.source_type in allowed]

    query_tokens = _tokenize(inp.query)
    if not candidates or not query_tokens:
        return SearchCySupportOutput(results=[])

    # Gate on lexical overlap before ranking, don't rely on "BM25 score > 0"
    # as the no-match signal. BM25's IDF is log((N-n+0.5)/(n+0.5)) -- with
    # the small per-control evidence sets this tool actually sees, a term
    # present in even one of a handful of documents can land n at roughly
    # N/2 and drive idf to ~0, silently dropping a real match. Overlap is
    # the real relevance gate; BM25 just orders the survivors.
    query_token_set = set(query_tokens)
    overlap_items: list[EvidenceItem] = []
    overlap_corpus: list[list[str]] = []
    for item in candidates:
        tokens = _tokenize(_searchable_text(item))
        if query_token_set & set(tokens):
            overlap_items.append(item)
            overlap_corpus.append(tokens)

    if not overlap_items:
        return SearchCySupportOutput(results=[])

    bm25 = BM25Okapi(overlap_corpus)
    scores = bm25.get_scores(query_tokens)

    ranked = sorted(zip(overlap_items, scores), key=lambda pair: pair[1], reverse=True)
    top = ranked[: inp.top_k]

    results = [
        SearchResult(
            evidence_id=item.evidence_id,
            source_file=item.source_file,
            location=item.location,
            snippet=_searchable_text(item)[:280],
            score=round(float(score), 4),
        )
        for item, score in top
    ]

    ctx.evidence_ids_returned_by_search.update(r.evidence_id for r in results)
    return SearchCySupportOutput(results=results)


def check_sample_coverage(
    inp: CheckSampleCoverageInput,
) -> CheckSampleCoverageOutput | CheckSampleCoverageError:
    # required_sample_ids is populated by the loop from the control's
    # SamplePopulationManifest. An empty list means no manifest was found
    # for this test step (see agent/loop.py) -- that's the error case, not
    # "a sample of size zero."
    if not inp.required_sample_ids:
        return CheckSampleCoverageError(error="no_sample_list_found")

    required = set(inp.required_sample_ids)
    found = set(inp.found_evidence_ids) & required
    missing = sorted(required - found)
    total_required = len(required)
    total_found = len(found)

    return CheckSampleCoverageOutput(
        total_required=total_required,
        total_found=total_found,
        missing=missing,
        coverage_pct=round(100.0 * total_found / total_required, 2) if total_required else 100.0,
        complete=not missing,
    )


def flag_exception(inp: FlagExceptionInput, ctx: ToolContext) -> FlagExceptionOutput:
    exception_id = f"exc_{next(ctx._exception_counter):04d}"
    ctx.exceptions[exception_id] = inp
    return FlagExceptionOutput(exception_id=exception_id)


def request_additional_support(
    inp: RequestAdditionalSupportInput, ctx: ToolContext
) -> RequestAdditionalSupportOutput:
    request_id = f"req_{next(ctx._request_counter):04d}"
    ctx.support_requests[request_id] = inp
    return RequestAdditionalSupportOutput(request_id=request_id)
