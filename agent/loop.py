"""Per-test-step tool loop from design doc sections 3, 6, and 7.

This is "the backend" the design doc keeps referring to: it's what appends
every tool call+result to the audit log automatically, what enforces the
forced-close-via-submit_conclusion rule, and what overrides
check_sample_coverage's required_sample_ids server-side rather than trusting
the model to restate the sample list correctly.

The Anthropic client is a constructor argument, not a module-level global,
specifically so this can be tested against a fake client with no network
call and no API key -- see agent/tests/test_loop.py. A real key is only
needed to actually run run_test_step() against the live API.
"""

from __future__ import annotations

import datetime
import json
from dataclasses import dataclass, field
from typing import Any, Callable, Protocol

from agent.schemas import (
    CheckSampleCoverageError,
    CheckSampleCoverageInput,
    ConclusionOutput,
    EvidenceItem,
    FlagExceptionInput,
    ModelMetadata,
    RequestAdditionalSupportInput,
    SampleCoverage,
    SampleItem,
    SamplePopulationManifest,
    SearchCySupportInput,
    SubmitConclusionInput,
)
from agent.tools import (
    ToolContext,
    check_sample_coverage,
    flag_exception,
    request_additional_support,
    search_cy_support,
)
from agent.schemas import validate_citations_against_transcript

# Matches the deployment this is actually run against. It must name a model
# that exists on the Foundry resource -- a default that doesn't just fails
# the run at the first API call, so "a sensible general default" is worth
# less here than "the one that works".
DEFAULT_MODEL = "claude-opus-5"
PROMPT_VERSION = "v1"
MAX_TOOL_ITERATIONS = 15

# Hard spending cap, in COST-WEIGHTED token units (fresh-input-token
# equivalents: cache writes ~1.25x, cache reads ~0.1x, output ~5x -- see
# _usage_tokens), summed across every turn of one test step. A real run
# without any cap hit ~400K input tokens on a SINGLE request before the
# extraction/caching fixes, and separately burned $65 with zero output by
# grinding through MAX_TOOL_ITERATIONS turns without ever calling
# submit_conclusion. An earlier version of this cap summed usage fields
# raw, which counted cache reads at 10x their real price and aborted a
# legitimately-progressing run mid-investigation -- cost-weighting is what
# makes the number mean "roughly proportional to dollars." On Opus-class
# input pricing, 300K weighted units is on the order of a few dollars per
# test step -- the ceiling for a runaway, well above any healthy run.
MAX_TOTAL_TOKENS = 300_000

# Output ceiling per turn. 4096 was too small for a real submit_conclusion:
# a 2-sample step produced a multi-paragraph narrative plus nine evidence
# citations with full quotes, and the JSON was cut off mid-object. The
# fields at the END of the schema (confidence, confidence_rationale,
# additional_support_requests, exceptions) simply never arrived, so
# validation rejected it as "Field required" -- which reads like the model
# forgot them rather than that it was truncated. It retried with a LONGER
# narrative each time and lost one more field, burning ~60K cost-weighted
# tokens on three failures that could never have succeeded.
MAX_OUTPUT_TOKENS = 16_000

_TRUNCATION_NOTICE = (
    "NOTE: your previous response hit the output length limit and was cut off "
    "mid-tool-call. Any 'Field required' error above is a symptom of that "
    "truncation, NOT of a field you forgot -- repeating the same call with a "
    "longer narrative will fail the same way. Resubmit with a materially "
    "shorter response: condense the narrative, and keep each "
    "quote_or_summary to the specific value or sentence that matters rather "
    "than reproducing a whole document."
)

# Section 0 of the design doc: this block goes into the cacheable system
# prompt prefix so the model reads company vocabulary correctly instead of
# treating it as unusual. Keep this in sync with docs/cy_testing_agent_design.md
# section 0 -- it's copied here rather than parsed from the doc so the
# runtime doesn't depend on markdown staying byte-identical.
COMPANY_CONTEXT = """\
COMPANY CONTEXT

BrightView Holdings (NYSE: BV) is the largest provider of commercial
landscaping services in the United States, operating through two reportable
segments:
- Maintenance Services: recurring, largely seasonal/evergreen landscaping
  work (mowing, irrigation, tree care, snow removal) under ongoing
  contracts.
- Development Services: project-based landscape architecture and
  construction for new facilities and major redesigns.

Operating model: 280+ branches perform the work locally (crews, equipment,
customer relationships); accounting and financial reporting are centralized
at corporate. Most SOX controls sit at corporate; a smaller set are
branch-level/transactional (e.g. purchase approvals, timekeeping, system
access, asset/vehicle controls) and are tested via samples drawn across
branches.

Financial systems referenced in evidence:
- E1 (JD Edwards EnterpriseOne) -- the core financial/ERP system.
- BVE1 -- a separate branch-facing instance of E1 used to reconcile branch
  activity back to corporate E1.
- OneStream -- corporate consolidation and financial reporting.

When evidence references one of these systems, a report name, or a branch
number, treat that as normal company vocabulary, not something to flag as
unusual on its own.
"""

SYSTEM_PROMPT = f"""\
You are a SOX control testing agent. You draft conclusions for human
auditors to review and approve -- you do not finalize workpapers, and
nothing you produce is final until a human approves it.

Non-negotiables:
- Cite only evidence returned by search_cy_support. Never invent an
  evidence_id, a quote, or a source location.
- "insufficient_evidence" is exactly as valid a conclusion as "satisfied" --
  there is no penalty for concluding you don't have enough to go on, and no
  extra justification is required for it beyond any other conclusion.
- Prior-year (PY) testing shown to you is precedent for FORMAT and APPROACH
  only -- what the test step looked like, what procedures were performed,
  what "satisfied" looked like structurally. PY having passed says nothing
  about whether this year's evidence supports the same conclusion. Evaluate
  current-year (CY) evidence independently every time.
- Every turn must end in a tool call. Gather evidence with search_cy_support
  before concluding anything. When you have enough to conclude -- including
  concluding you don't -- call submit_conclusion. That is the only way this
  conversation ends.
- IPE: when the evidence includes both a population extract and the report
  parameters it was produced from (a query/parameters screenshot, report
  header, or run log), actually reconcile them -- agree the record count in
  the parameters to the number of rows in the extract, and check the filters
  match the control's scope and period. A record count that does not tie, or
  filters that do not match the control, is an exception to flag, not a
  detail to pass over. Only mark ipe_completeness_accuracy_status
  "validated" when you have performed that reconciliation and it agrees.

{COMPANY_CONTEXT}"""


@dataclass
class TestStepRequest:
    __test__ = False  # not a pytest test class -- name collision with the "Test" prefix convention

    test_step_id: str
    control_id: str
    control_objective_ref: str
    control_objective_text: str
    test_step_text: str
    py_support_excerpts: list[EvidenceItem] = field(default_factory=list)
    # Optional -- if you already have it, it saves the model a search; if
    # not, the PY conclusion is normally readable from py_support_excerpts
    # itself, so there's no need to make a human retype it.
    py_conclusion_text: str = ""
    # From the SamplePopulationManifest, when one exists. Without this the
    # model's ONLY signal about sample size is check_sample_coverage's
    # total_required count, with zero context for why it's that size -- a
    # real run saw the model treat a correct, intentional 1-item sample as
    # suspicious (a CY support filename said "selection 1", implying more
    # might exist) purely because it had no population figure to check that
    # against. Both are None when not known -- the line is omitted rather
    # than shown as a suspicious "0" or "unknown".
    sample_size: int | None = None
    population_size: int | None = None
    # The actual selected items, from the SamplePopulationManifest. Without
    # these the model gets a bare COUNT and has to discover both the
    # sample_ids (by calling check_sample_coverage and reading `missing`)
    # and what each item even is (by fishing with searches). With one
    # sample that was survivable -- everything in CY support related to it.
    # With two it burned a real run's whole budget.
    samples: list[SampleItem] = field(default_factory=list)


def _render_evidence_item(item: EvidenceItem) -> str:
    body = item.extracted_text or (
        "\n".join(" | ".join(row) for row in item.extracted_table) if item.extracted_table else "[no text]"
    )
    return f"[{item.evidence_id}] {item.location}\n{body}"


# Every PY excerpt gets rendered into every turn of the conversation with no
# per-item cap of its own -- a real run showed exactly how bad that gets: one
# large PY testing file (or, before extraction chunked large tables, one
# huge extracted table) pushed a single test step to ~400K input tokens per
# API call. This is a blunt safety net, not a design target -- the real fix
# is still the doc's flagged "confirm/edit step" that would show a human
# only the PY excerpts relevant to this specific test step.
_MAX_PY_EXCERPT_CHARS = 20_000


def _render_py_excerpts(items: list[EvidenceItem]) -> str:
    if not items:
        return "(no PY support excerpts provided)"

    parts: list[str] = []
    total = 0
    for item in items:
        rendered = _render_evidence_item(item)
        if parts and total + len(rendered) > _MAX_PY_EXCERPT_CHARS:
            break
        parts.append(rendered)
        total += len(rendered)

    text = "\n\n".join(parts)
    if len(parts) < len(items):
        text += (
            f"\n\n...({len(items) - len(parts)} more PY excerpt(s) omitted for length -- "
            f"showing {len(parts)} of {len(items)})"
        )
    return text


_MAX_INVENTORY_ITEMS = 40


def _render_cy_inventory(items: list[EvidenceItem]) -> str:
    """The complete map of what CY evidence exists, shown up front. Without
    this the model can only discover the evidence pool by fishing with
    searches -- a real run burned turns querying for documents that simply
    weren't there (a delegation-of-authority matrix, org charts) with no
    way to know that short of failed search after failed search. With the
    map, one glance answers "what do I have?" and searches are for READING
    items, not discovering them. Capped: a huge pool falls back to a count
    plus the first _MAX_INVENTORY_ITEMS entries.
    """
    if not items:
        return "(no CY evidence extracted)"
    lines = []
    for item in items[:_MAX_INVENTORY_ITEMS]:
        body = item.extracted_text or (
            " | ".join(item.extracted_table[0]) if item.extracted_table else ""
        )
        preview = " ".join(body.split())[:110]
        lines.append(f"[{item.evidence_id}] ({item.source_type}) {item.location} -- {preview}")
    text = "\n".join(lines)
    if len(items) > _MAX_INVENTORY_ITEMS:
        text += f"\n...and {len(items) - _MAX_INVENTORY_ITEMS} more item(s) -- discover them via search_cy_support."
    return text


def build_user_turn(request: TestStepRequest, cy_evidence: list[EvidenceItem] | None = None) -> str:
    py_excerpts = _render_py_excerpts(request.py_support_excerpts)
    py_conclusion_line = (
        f"PY conclusion: {request.py_conclusion_text}\n\n"
        if request.py_conclusion_text
        else "PY conclusion: not separately provided -- read it from the PY support excerpts below if relevant.\n\n"
    )
    sample_line = _render_sample_line(request.sample_size, request.population_size, request.samples)
    inventory_block = ""
    if cy_evidence is not None:
        inventory_block = f"""\
CY evidence inventory -- this is the COMPLETE list of evidence extracted from
this control's CY support files. If a document type is not in this list (e.g.
a policy, an authority matrix, an org chart), it was not provided: do not
spend searches fishing for it -- note it via request_additional_support
instead. You must still retrieve an item with search_cy_support before citing
it; this inventory is a map, not retrieved evidence.
{_render_cy_inventory(cy_evidence)}

"""
    return f"""\
Control objective ({request.control_objective_ref}): {request.control_objective_text}

Test step ({request.test_step_id}): {request.test_step_text}

{sample_line}{inventory_block}{py_conclusion_line}PY support excerpts (format/approach precedent only -- not evidence for this
year's conclusion):
{py_excerpts}

Use search_cy_support to find this year's evidence for this test step before
concluding anything. Be efficient: searches return items nearly in full, so
a handful of searches should cover a small evidence pool -- do not re-query
the same document for different fields."""


_MAX_ROSTER_ITEMS = 25
# Budgeted across the WHOLE roster rather than a fixed cap per item. A real
# E1 sample row is ~560 chars over 21 columns, and a flat 240-char cap
# truncated it mid-record -- keeping "Discount Available: 0" and cutting
# off the invoice number, invoice date and business unit, i.e. exactly the
# fields a tickmark or a search would key on. Budgeting by total means the
# common small sample (1-3 items) gets its records in full, and only a
# genuinely large sample gets abbreviated.
_MAX_ROSTER_TOTAL_CHARS = 3_000
_MIN_ROSTER_DETAIL_CHARS = 120

# Values that carry no identifying information -- zero discounts, blank
# void dates, empty PO numbers. Dropped before budgeting so the characters
# go to fields that actually distinguish one sampled item from another.
_UNINFORMATIVE_VALUES = {"", "0", "0.0", "0.00", "none", "n/a"}


def _sample_detail(item: SampleItem) -> str:
    fields = item.key_fields
    if not fields:
        return " ".join(item.identifying_details.split())
    parts = [
        f"{k}: {v}"
        for k, v in fields.items()
        if str(v).strip().lower() not in _UNINFORMATIVE_VALUES
    ]
    return "; ".join(parts) if parts else " ".join(item.identifying_details.split())


def _render_sample_roster(samples: list[SampleItem]) -> str:
    """The selected items themselves -- sample_id plus the field values that
    identify each one. Without this the model knows only how MANY items were
    selected, so on a multi-item sample it has to work out both the
    sample_ids and what each item is by trial and error.
    """
    if not samples:
        return ""

    shown = samples[:_MAX_ROSTER_ITEMS]
    per_item = max(_MIN_ROSTER_DETAIL_CHARS, _MAX_ROSTER_TOTAL_CHARS // len(shown))

    lines = []
    for item in shown:
        detail = _sample_detail(item)
        if len(detail) > per_item:
            detail = detail[:per_item] + "..."
        lines.append(f"  sample_id {item.sample_id!r}: {detail}")
    if len(samples) > _MAX_ROSTER_ITEMS:
        lines.append(f"  ...and {len(samples) - _MAX_ROSTER_ITEMS} more item(s).")

    return (
        "The selected items, exactly as they appear in the sample listing. These "
        "sample_id values are what check_sample_coverage expects in "
        "found_evidence_ids:\n"
        + "\n".join(lines)
        + "\nFind CY support for EACH of these separately. Match evidence to an item on "
        "its field values (invoice number, vendor, amount, dates), NOT on a filename -- a "
        "support file may be named for a selection number, mis-named, or cover several "
        "items at once.\n\n"
    )


def _render_sample_line(
    sample_size: int | None, population_size: int | None, samples: list[SampleItem] | None = None
) -> str:
    if sample_size is None:
        return ""
    roster = _render_sample_roster(samples or [])
    if population_size is not None:
        header = (
            f"Sample: {sample_size} item(s) selected for testing from a population of "
            f"{population_size}. This is the correct, complete sample -- do not treat it as "
            "partial just because a support filename or document label suggests otherwise.\n\n"
        )
    else:
        header = (
            f"Sample: {sample_size} item(s) selected for testing. The full underlying population "
            "size was not provided to you -- if evidence in CY support implies the tested sample "
            "should be larger (e.g. a filename or document says \"selection 1\" of several), use "
            "request_additional_support to ask for the full sample selection listing rather than "
            "assuming either that more items exist or that they don't.\n\n"
        )
    return header + roster


def _tool_def(name: str, description: str, model) -> dict[str, Any]:
    schema = model.model_json_schema()
    schema.pop("title", None)
    return {"name": name, "description": description, "input_schema": schema}


TOOLS = [
    _tool_def(
        "search_cy_support",
        "Search this control's current-year (CY) support evidence. Use this "
        "before drawing any conclusion, and again whenever you need evidence "
        "for a specific sample item. Returns ranked snippets with evidence_ids "
        "you can cite in submit_conclusion.",
        SearchCySupportInput,
    ),
    _tool_def(
        "check_sample_coverage",
        "Check which of this test step's required samples you have evidence "
        "for. Pass the sample_ids (not evidence_ids) you believe you've found "
        "evidence for as found_evidence_ids. Always call this before "
        "submit_conclusion when the test step has a sample -- do not eyeball "
        "sample coverage yourself.",
        CheckSampleCoverageInput,
    ),
    _tool_def(
        "flag_exception",
        "Record a control exception found in the evidence, independent of "
        "your narrative. Use this whenever evidence shows the control did not "
        "operate as designed for a specific item.",
        FlagExceptionInput,
    ),
    _tool_def(
        "request_additional_support",
        "Request more support from the control owner when evidence is "
        "missing, illegible, ambiguous, or the sample isn't fully covered. "
        "This is a first-class action -- use it freely; it costs nothing to "
        "call and is exactly as valid an outcome as concluding satisfied.",
        RequestAdditionalSupportInput,
    ),
    _tool_def(
        "submit_conclusion",
        "Submit your final conclusion for this test step. This is the only "
        "way to end the conversation -- call it once you've gathered enough "
        "evidence (via search_cy_support) to conclude, including concluding "
        "insufficient_evidence. Note: when the test step has a sample, a "
        "'satisfied' conclusion is only accepted after a check_sample_coverage "
        "call showing complete coverage. Set each citation's sample_id to the "
        "sampled item it supports, so the workpaper can document each selection "
        "separately; leave it null only for evidence covering the step as a "
        "whole (a policy, the population extract, IPE parameters). Also fill "
        "sample_results with one entry per sampled item saying how the step "
        "came out FOR THAT ITEM -- an exception is always about a specific "
        "selection, and the step-level conclusion alone cannot say which one "
        "failed. Finally, break the test step into the individual attributes "
        "it actually requires (e.g. approver is independent, approver has "
        "authority, approval precedes payment, amount and coding agree) and "
        "report each in attribute_results per sampled item, with the value "
        "you observed and the evidence_ids proving it. A reviewer signs off "
        "attribute by attribute; do not make them mine the narrative to find "
        "where one was satisfied. State each finding ONCE, in the field that "
        "owns it: attribute_results carries the observed values and figures; "
        "exceptions and additional_support_requests are one concise line each "
        "and should not re-derive that detail; procedures_performed says what "
        "you DID, not what you found. A real workpaper repeated a single IPE "
        "break six times across these fields -- that is harder to review, not "
        "more thorough.",
        SubmitConclusionInput,
    ),
]

_INPUT_MODELS: dict[str, Any] = {
    "search_cy_support": SearchCySupportInput,
    "check_sample_coverage": CheckSampleCoverageInput,
    "flag_exception": FlagExceptionInput,
    "request_additional_support": RequestAdditionalSupportInput,
    "submit_conclusion": SubmitConclusionInput,
}


class AnthropicClientLike(Protocol):
    """The one method this module calls on the client -- narrow on purpose
    so a test double only has to implement this, not the whole SDK surface.
    """

    def create_message(self, **kwargs) -> Any: ...


def _call_model(client: "AnthropicClientLike", **kwargs) -> Any:
    # Real anthropic.Anthropic() exposes client.messages.create(...); fakes
    # used in tests can implement either shape.
    if hasattr(client, "messages"):
        return client.messages.create(**kwargs)
    return client.create_message(**kwargs)


def _mark_cache_breakpoint(messages: list[dict]) -> None:
    """Marks the last content block of the last message as a cache
    breakpoint, converting a bare string into block form if needed. By the
    time this runs, messages[-1] is always either the initial user turn
    (str) or a tool_results/nudge turn we appended ourselves (list[dict]) --
    never the assistant's raw SDK response blocks, since we always append a
    user turn after those before looping back to call the model again.

    Strips any breakpoint from earlier messages first. Caching still works
    via prefix-hash matching against what the server already cached from a
    previous turn's marker -- a marker doesn't need to still be physically
    present at that position in a later request for the server to find the
    match. Without stripping, markers would keep accumulating turn over
    turn and blow through the API's 4-breakpoint-per-request limit on any
    conversation longer than a few turns, which is the common case here.
    """
    for m in messages[:-1]:
        content = m.get("content")
        if isinstance(content, list):
            for block in content:
                if isinstance(block, dict):
                    block.pop("cache_control", None)

    last = messages[-1]
    content = last["content"]
    if isinstance(content, str):
        last["content"] = [{"type": "text", "text": content, "cache_control": {"type": "ephemeral"}}]
    elif isinstance(content, list) and content and isinstance(content[-1], dict):
        content[-1] = {**content[-1], "cache_control": {"type": "ephemeral"}}


@dataclass
class AuditLogEntry:
    turn: int
    tool_name: str
    tool_use_id: str
    input: dict
    output: dict
    is_error: bool
    timestamp: str


class IncompleteRunError(RuntimeError):
    """Raised when a test step is aborted before submit_conclusion -- either
    the iteration count or the token budget ran out first. Unlike a bare
    RuntimeError, this carries everything already paid for: the audit log of
    every tool call made so far, and the token/turn counts that triggered
    the abort. A caller can show that instead of a total loss (see
    agent/run_control.py's iter_control_results, which surfaces these on the
    error dict it yields).
    """

    def __init__(
        self,
        message: str,
        *,
        reason: str,
        audit_log: list[AuditLogEntry],
        tokens_used: int,
        turns_used: int,
    ) -> None:
        super().__init__(message)
        self.reason = reason
        self.audit_log = audit_log
        self.tokens_used = tokens_used
        self.turns_used = turns_used


def _usage_tokens(usage: Any) -> int:
    """Cost-weighted token units for the spending budget, in fresh-input-token
    equivalents. A raw sum of every usage field counted cache READS at full
    price -- but a cache read bills at roughly 1/10th of fresh input, and on
    a well-cached multi-turn run cache reads dominate the raw count. A real
    run "spent" 300K raw-counted tokens in 12 turns and got budget-aborted
    mid-investigation when its actual bill was a small fraction of what
    300K fresh tokens would cost. Weights (relative to fresh input = 1.0):
    cache write ~1.25x, cache read ~0.1x, output ~5x -- the standard
    Anthropic price ratios, stable across models.

    Missing fields (a minimal test double, or an SDK version without cache
    fields) default to 0 rather than raising. The ``or 0`` matters: on the
    real SDK the cache fields are Optional[int] -- the attribute EXISTS but
    is None when unpopulated, so getattr's default alone never applies and
    a bare sum would raise TypeError on the first real turn.
    """
    if usage is None:
        return 0
    return round(
        (getattr(usage, "input_tokens", 0) or 0)
        + 5.0 * (getattr(usage, "output_tokens", 0) or 0)
        + 1.25 * (getattr(usage, "cache_creation_input_tokens", 0) or 0)
        + 0.1 * (getattr(usage, "cache_read_input_tokens", 0) or 0)
    )


def _execute_tool(
    block: Any, ctx: ToolContext, turn: int, audit_log: list[AuditLogEntry]
) -> tuple[dict, bool, ConclusionOutput | None]:
    """Returns (content_for_claude, is_error, conclusion_if_this_was_a_successful_submit)."""
    # Count every call up front, before any early return -- a call that fails
    # basic input validation (e.g. the IPE status/evidence rule) is still a
    # real tool call and belongs in the count. This used to increment only at
    # the bottom of the function, so a rejected submit_conclusion attempt was
    # logged in the audit trail but silently missing from tool_call_count --
    # caught live when a real run showed model_metadata.tool_call_count=13
    # against an audit log of 14 entries.
    ctx.tool_call_count += 1

    model_cls = _INPUT_MODELS.get(block.name)
    if model_cls is None:
        content = {"error": f"unknown tool {block.name!r}"}
        audit_log.append(
            AuditLogEntry(turn, block.name, block.id, dict(block.input), content, True, _now())
        )
        return content, True, None

    try:
        parsed = model_cls.model_validate(block.input)
    except Exception as exc:  # noqa: BLE001 -- surfaced to the model, not swallowed
        content = {"error": f"invalid input: {exc}"}
        audit_log.append(
            AuditLogEntry(turn, block.name, block.id, dict(block.input), content, True, _now())
        )
        return content, True, None

    conclusion: ConclusionOutput | None = None

    if block.name == "search_cy_support":
        output = search_cy_support(parsed, ctx)
        content = output.model_dump()
        is_error = False

    elif block.name == "check_sample_coverage":
        # Server-side override: never trust the model's own required_sample_ids
        # echo, even though the tool's public schema (matching the design doc)
        # accepts one. ctx.required_sample_ids is the truth from the
        # SamplePopulationManifest for this test step.
        parsed = parsed.model_copy(update={"required_sample_ids": ctx.required_sample_ids})
        output = check_sample_coverage(parsed)
        content = output.model_dump()
        is_error = isinstance(output, CheckSampleCoverageError)
        if not is_error:
            ctx.last_sample_coverage = SampleCoverage(**output.model_dump())

    elif block.name == "flag_exception":
        output = flag_exception(parsed, ctx)
        content = output.model_dump()
        is_error = False

    elif block.name == "request_additional_support":
        output = request_additional_support(parsed, ctx)
        content = output.model_dump()
        is_error = False

    elif block.name == "submit_conclusion":
        gate_error: str | None = None
        try:
            validate_citations_against_transcript(parsed, ctx.evidence_ids_returned_by_search)
        except ValueError as exc:
            gate_error = str(exc)

        # Server-enforced, like the fabrication guard: the system prompt and
        # tool descriptions ASK for check_sample_coverage before concluding,
        # but a "satisfied" on a sampled test step with unverified or
        # incomplete coverage is exactly the conclusion an audit reviewer
        # can't accept, so the backend refuses it rather than trusting the
        # model to have followed the instruction.
        if gate_error is None and parsed.conclusion == "satisfied" and ctx.required_sample_ids:
            if ctx.last_sample_coverage is None:
                gate_error = (
                    "a 'satisfied' conclusion for a test step with a sample requires a "
                    "successful check_sample_coverage call first. Call it, then resubmit -- "
                    "or, if coverage can't be established, conclude not_satisfied or "
                    "insufficient_evidence and use request_additional_support for what's missing."
                )
            elif not ctx.last_sample_coverage.complete:
                sc = ctx.last_sample_coverage
                gate_error = (
                    f"sample coverage is incomplete ({sc.total_found}/{sc.total_required}; "
                    f"missing: {sc.missing}) -- 'satisfied' is not supportable. Search for "
                    "evidence covering the missing samples and re-run check_sample_coverage, "
                    "or conclude not_satisfied/insufficient_evidence and request additional support."
                )

        if gate_error is not None:
            content = {"error": gate_error}
            is_error = True
        else:
            conclusion = ConclusionOutput(
                **parsed.model_dump(),
                sample_coverage=ctx.last_sample_coverage,
                model_metadata=ModelMetadata(
                    model=ctx.model,
                    prompt_version=PROMPT_VERSION,
                    timestamp=_now(),
                    tool_call_count=ctx.tool_call_count,
                ),
            )
            content = {"recorded": True}
            is_error = False
    else:  # pragma: no cover -- guarded by the model_cls lookup above
        content = {"error": f"unhandled tool {block.name!r}"}
        is_error = True

    audit_log.append(AuditLogEntry(turn, block.name, block.id, dict(block.input), content, is_error, _now()))
    return content, is_error, conclusion


def _now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


_WRAPUP_WARNING = (
    "IMPORTANT: this test step is close to its turn/token limit. Do not start "
    "new lines of investigation. On your next turn, call submit_conclusion "
    "with your best-supported conclusion -- insufficient_evidence with "
    "additional_support_requests listing exactly what's missing is a fully "
    "valid close, and far better than running out of turns with no "
    "conclusion at all."
)


def _append_wrapup_warning(messages: list[dict]) -> None:
    """Appends the wrap-up warning as a text block on the newest user
    message (tool results or nudge), converting a bare-string message to
    block form if needed. The API allows text blocks after tool_result
    blocks in the same user message.
    """
    last = messages[-1]
    if isinstance(last["content"], str):
        last["content"] = [{"type": "text", "text": last["content"]}]
    last["content"].append({"type": "text", "text": _WRAPUP_WARNING})


def run_test_step(
    request: TestStepRequest,
    evidence_items: list[EvidenceItem],
    sample_manifest: SamplePopulationManifest | None,
    client: "AnthropicClientLike",
    model: str = DEFAULT_MODEL,
    max_iterations: int = MAX_TOOL_ITERATIONS,
    max_total_tokens: int = MAX_TOTAL_TOKENS,
    on_turn: "Callable[[int, int, list[AuditLogEntry]], None] | None" = None,
) -> tuple[ConclusionOutput, list[AuditLogEntry]]:
    """on_turn, if given, is called once per completed turn with
    (turn_number, cumulative_tokens_used, audit_log_so_far) -- purely for a
    caller (the Streamlit app) to show live progress while a step is still
    running. It sees the same audit_log list this function keeps appending
    to, so a caller that only reads it (rather than mutating it) is safe;
    nothing here depends on its return value.
    """
    ctx = ToolContext(
        evidence_items=evidence_items,
        model=model,
        required_sample_ids=[s.sample_id for s in sample_manifest.samples] if sample_manifest else [],
    )

    audit_log: list[AuditLogEntry] = []
    tokens_used = 0
    wrapup_warned = False
    system = [{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}]
    messages: list[dict] = [{"role": "user", "content": build_user_turn(request, cy_evidence=evidence_items)}]

    def _maybe_warn_wrapup(turn: int) -> None:
        # The $65 failure mode was the model grinding through every allowed
        # turn with no idea a limit existed, then aborting with nothing.
        # One warning, injected a turn before either cliff (last iteration,
        # or 80% of the token budget), gives it the chance to close with a
        # real conclusion instead. Fired at most once -- if the model
        # ignores it, the hard aborts below still cap the spend.
        nonlocal wrapup_warned
        if wrapup_warned:
            return
        if turn >= max_iterations - 1 or tokens_used >= 0.8 * max_total_tokens:
            _append_wrapup_warning(messages)
            wrapup_warned = True

    for turn in range(1, max_iterations + 1):
        # Without this, every turn re-sends (and re-bills at full price) the
        # ENTIRE growing conversation -- the API is stateless, so nothing
        # about system's cache_control above covers messages[]. This marks
        # only the newest last block each turn (not every prior one, which
        # would exceed the 4-breakpoint-per-request limit on a long-running
        # step) -- a fresh breakpoint still finds the server-side cache
        # written by the previous turn's marker, since matching is by prefix
        # content, not by which positions carry a literal marker right now.
        _mark_cache_breakpoint(messages)
        response = _call_model(
            client,
            model=model,
            max_tokens=MAX_OUTPUT_TOKENS,
            system=system,
            tools=TOOLS,
            messages=messages,
        )

        tokens_used += _usage_tokens(getattr(response, "usage", None))

        if getattr(response, "stop_reason", None) == "refusal":
            # Same preserve-paid-work rule as the budget/iteration aborts: a
            # refusal on turn 8 still had 7 turns of real, billed tool calls
            # behind it -- carry them out instead of discarding them.
            raise IncompleteRunError(
                f"model declined the request: {getattr(response, 'stop_details', None)}",
                reason="model_refusal",
                audit_log=audit_log,
                tokens_used=tokens_used,
                turns_used=turn,
            )

        was_truncated = getattr(response, "stop_reason", None) == "max_tokens"

        messages.append({"role": "assistant", "content": response.content})
        tool_use_blocks = [b for b in response.content if getattr(b, "type", None) == "tool_use"]

        if not tool_use_blocks:
            nudge = (
                "Every turn must end in a tool call. Use search_cy_support to "
                "keep gathering evidence, or submit_conclusion when ready."
            )
            # Both, not either: a truncated plain-text turn still made no
            # tool call, and dropping the forced-close instruction here let
            # the model reply with shorter prose and burn another turn
            # without ever calling a tool.
            messages.append(
                {
                    "role": "user",
                    "content": f"{_TRUNCATION_NOTICE}\n\n{nudge}" if was_truncated else nudge,
                }
            )
            if on_turn is not None:
                on_turn(turn, tokens_used, audit_log)
            if tokens_used >= max_total_tokens:
                raise IncompleteRunError(
                    f"test step {request.test_step_id!r} exceeded the {max_total_tokens:,}-token "
                    f"budget after {turn} turn(s) without reaching submit_conclusion",
                    reason="token_budget_exceeded",
                    audit_log=audit_log,
                    tokens_used=tokens_used,
                    turns_used=turn,
                )
            _maybe_warn_wrapup(turn)
            continue

        tool_results = []
        final_conclusion: ConclusionOutput | None = None
        for block in tool_use_blocks:
            content, is_error, conclusion = _execute_tool(block, ctx, turn, audit_log)
            tool_results.append(
                {
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": json.dumps(content),
                    "is_error": is_error,
                }
            )
            if conclusion is not None:
                final_conclusion = conclusion

        if was_truncated and final_conclusion is None:
            # The tool call itself was cut off mid-JSON, so the validation
            # error above names whichever schema fields fell off the end.
            # Say so explicitly -- left to interpret "Field required" alone,
            # a real run retried three times with an ever-LONGER narrative,
            # losing one more field each time.
            tool_results.append({"type": "text", "text": _TRUNCATION_NOTICE})

        messages.append({"role": "user", "content": tool_results})

        if on_turn is not None:
            on_turn(turn, tokens_used, audit_log)

        if final_conclusion is not None:
            return final_conclusion, audit_log

        if tokens_used >= max_total_tokens:
            raise IncompleteRunError(
                f"test step {request.test_step_id!r} exceeded the {max_total_tokens:,}-token "
                f"budget after {turn} turn(s) without reaching submit_conclusion",
                reason="token_budget_exceeded",
                audit_log=audit_log,
                tokens_used=tokens_used,
                turns_used=turn,
            )

        _maybe_warn_wrapup(turn)

    raise IncompleteRunError(
        f"test step {request.test_step_id!r} did not reach submit_conclusion "
        f"within {max_iterations} tool-loop iterations",
        reason="max_iterations",
        audit_log=audit_log,
        tokens_used=tokens_used,
        turns_used=max_iterations,
    )
