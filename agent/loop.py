"""Per-test-step tool loop from design doc sections 3, 6, and 7.

This is "the backend" the design doc keeps referring to: it's what appends
every tool call+result to the audit log automatically, what enforces the
forced-close-via-submit_conclusion rule, and what overrides
sample coverage server-side from the submitted conclusion rather than
trusting the model to report what it covered.

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

from agent.costs import TokenLedger
from agent.schemas import (
    ConclusionOutput,
    EvidenceItem,
    FlagExceptionInput,
    ModelMetadata,
    RequestAdditionalSupportInput,
    SampleCoverage,
    SampleItem,
    SamplePopulationManifest,
    LocateValueInput,
    SearchCySupportInput,
    SubmitConclusionInput,
)
from agent.tools import (
    ToolContext,
    _searchable_text,
    flag_exception,
    locate_value,
    request_additional_support,
    search_cy_support,
)
from agent.schemas import validate_citations_against_transcript

# Matches the deployment this is actually run against. It must name a model
# that exists on the Foundry resource -- a default that doesn't just fails
# the run at the first API call, so "a sensible general default" is worth
# less here than "the one that works".
DEFAULT_MODEL = "claude-opus-5"
PROMPT_VERSION = "v2"
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
- You test THIS YEAR's selections against THIS YEAR's support only. The PY
  workpaper shown to you is LAST YEAR's testing of LAST YEAR's items: its
  invoices, dates, amounts, names and documents are not this year's and
  must never be reported as observed, in any field except py_comparison.
  The backend rejects a conclusion that reports a value found only in the
  PY workpaper.
- Every turn must end in a tool call. Gather evidence with search_cy_support
  before concluding anything. When you have enough to conclude -- including
  concluding you don't -- call submit_conclusion. That is the only way this
  conversation ends.
- IPE (completeness and accuracy of the population) is NOT yours to test.
  A human performs it. Do not reconcile record counts, do not evaluate
  report parameters, and do not conclude on whether the population is
  complete -- there is no field for it and no credit for attempting it.
  If something about the population looks wrong while you are testing,
  raise it with flag_exception like any other finding.

How to reach a conclusion -- professional judgment, not literalism:
- Stay skeptical, but conclude on the TOTALITY of the evidence the way an
  experienced auditor would. Evidence rarely states every attribute in so
  many words. An attribute can be supported by reasonable inference from
  what IS there -- for example an email chain showing who approved and
  when, a system timestamp or workflow status showing a step happened, a
  sign-off or initials on the document, or one document corroborating
  another.
- For each attribute, decide which of three cases you are in:
  (a) The CY evidence supports it directly.
  (b) The CY evidence is incomplete, but a reasonable auditor would accept
      it on inference. Treat this as supported, and state the inference and
      why it is reasonable in judgment_notes.
  (c) The evidence is missing, or it contradicts the attribute. Only this
      is a failure: not_satisfied when the evidence contradicts it,
      insufficient_evidence (with request_additional_support) when it is
      missing.
  Do not put an attribute in (c) merely because the evidence is less
  explicit than you would like.
- Skepticism is not reduced. Red flags (an approver who is also the
  requester, approval dated after payment, altered or inconsistent
  documents, values that do not agree), contradictions between documents,
  and sampled items with no support at all are still raised with
  flag_exception, every time. An inference never explains away a red flag
  or a contradiction -- if the evidence points both ways, it is (c).

Using the prior year (PY) as a benchmark, for each test step:
- Identify the PY evidence that supported the PY conclusion (from the PY
  support excerpts and PY conclusion), and compare the CY evidence to it:
  the type of document, its source (system, report, person), and its level
  of detail.
- If the CY evidence is substantially the same and supports the step, reach
  the same conclusion the PY did and cite the comparison: set
  conclusion_basis to "PY comparison" and describe it in py_comparison. The
  CY evidence must still be retrieved and cited -- the PY tells you what
  sufficient evidence looked like, it is never evidence for this year.
- Do NOT rely on the PY, and test on CY evidence alone, when any of these
  is true: the control or process was redesigned; a system, report or
  control owner changed; the population changed in nature; the PY result
  for this step was itself an exception or not satisfied; or no usable PY
  evidence was provided. Say in py_comparison which of these applied.
- Whatever the basis, record it in conclusion_basis ("CY evidence direct",
  "CY evidence inferred" or "PY comparison") and keep py_comparison to one
  or two sentences. Fill py_comparison even when you did not rely on the
  PY, saying why. If a PY comparison also involved an inference, use
  "PY comparison" and state the inference in judgment_notes as well.

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
    # model's ONLY signal about sample size would be the roster length,
    # with zero context for why it's that size -- a
    # real run saw the model treat a correct, intentional 1-item sample as
    # suspicious (a CY support filename said "selection 1", implying more
    # might exist) purely because it had no population figure to check that
    # against. Both are None when not known -- the line is omitted rather
    # than shown as a suspicious "0" or "unknown".
    sample_size: int | None = None
    population_size: int | None = None
    # The actual selected items, from the SamplePopulationManifest. Without
    # these the model gets a bare COUNT and has to discover both the
    # sample_ids and what each item even is (by fishing with searches). With one
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


def user_turn_sections(
    request: TestStepRequest, cy_evidence: list[EvidenceItem] | None = None
) -> dict[str, str]:
    """The first user turn as its named pieces, in render order.

    build_user_turn() is just "".join() of these -- they are split out so
    the cost ledger can report which section of the prompt the turn-1 input
    bill actually went to (agent/costs.py's prompt_mix_rows). Turn 1 is the
    expensive turn of a step, and "why" is answerable only if the prompt's
    parts are separable rather than one f-string.
    """
    py_excerpts = _render_py_excerpts(request.py_support_excerpts)
    py_conclusion_line = (
        f"PY conclusion: {request.py_conclusion_text}\n\n"
        if request.py_conclusion_text
        else "PY conclusion: not separately provided -- read it from the PY support excerpts below if relevant.\n\n"
    )
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
    return {
        "Control objective + test step": (
            f"Control objective ({request.control_objective_ref}): {request.control_objective_text}\n"
            f"\n"
            f"Test step ({request.test_step_id}): {request.test_step_text}\n"
            f"\n"
        ),
        "Sample roster": _render_sample_line(
            request.sample_size, request.population_size, request.samples
        ),
        "CY evidence inventory": inventory_block,
        "PY conclusion": py_conclusion_line,
        "PY support excerpts": (
            "LAST YEAR's workpaper (PY) -- last year's items and last year's support. It is the\n"
            "benchmark for what sufficient evidence looked like; none of its items, dates,\n"
            "amounts or names are this year's, and none may be reported as observed:\n"
            f"{py_excerpts}\n"
            "\n"
        ),
        "Closing instructions": (
            "Use search_cy_support to find this year's evidence for this test step before\n"
            "concluding anything. Be efficient: searches return items nearly in full, so\n"
            "a handful of searches should cover a small evidence pool -- do not re-query\n"
            "the same document for different fields."
        ),
    }


def build_user_turn(request: TestStepRequest, cy_evidence: list[EvidenceItem] | None = None) -> str:
    return "".join(user_turn_sections(request, cy_evidence).values())


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
        "sample_id values are what every sample_result, attribute_result and "
        "citation must use:\n"
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
        "locate_value",
        "Check whether a value is literally PRINTED on one piece of evidence, "
        "before you rely on it as a tickmark anchor. The workpaper draws a red "
        "box around each attribute's value on the exhibit by searching the page "
        "for the text you report in attribute_results -- so if you report a date "
        "as '02/09/2026' and the document spells it 'February 9, 2026', the box "
        "cannot be drawn and the tickmark is lost. Use this whenever you are not "
        "certain of the exact printed form, especially for dates and names. A "
        "miss returns what IS on the page so you can anchor on that instead.",
        LocateValueInput,
    ),
    _tool_def(
        "flag_exception",
        "Record a control exception found in the evidence, independent of "
        "your narrative. Use this whenever evidence shows the control did not "
        "operate as designed for a specific item, and for every red flag, "
        "contradiction between documents, or sampled item with no support -- "
        "professional judgment about incomplete evidence never replaces this.",
        FlagExceptionInput,
    ),
    _tool_def(
        "request_additional_support",
        "Request more support from the control owner when evidence is "
        "missing, illegible, or contradictory, or the sample isn't fully "
        "covered -- case (c). Do not use it for evidence a reasonable auditor "
        "would accept on inference (case (b)); conclude on that and explain "
        "the inference in judgment_notes instead.",
        RequestAdditionalSupportInput,
    ),
    _tool_def(
        "submit_conclusion",
        "Submit your final conclusion for this test step. This is the only "
        "way to end the conversation -- call it once you've gathered enough "
        "evidence (via search_cy_support) to conclude, including concluding "
        "insufficient_evidence. Note: when the test step has a sample, a "
        "'satisfied' conclusion is only accepted if you conclude on EVERY "
        "sampled item -- the backend checks that against the sample listing "
        "and rejects a conclusion that silently skips one. Set each citation's sample_id to the "
        "sampled item it supports, so the workpaper can document each selection "
        "separately; leave it null only for evidence covering the step as a "
        "whole (a policy, the population extract). Also fill "
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
        "you DID, not what you found. A real workpaper repeated a single "
        "finding six times across these fields -- that is harder to review, "
        "not more thorough. The same rule governs the narrative, which is the "
        "first thing a reviewer reads in the finished workpaper: two or three "
        "sentences saying what you tested, against what, and how it came "
        "out. The attributes, values, dates, names and amounts do NOT belong "
        "there -- attribute_results already carries each of them beside the "
        "tickmark that points at it on the evidence. Ask yourself how few "
        "words can carry the key attributes, not how many you can fit. "
        "Set conclusion_basis to how you reached the conclusion, fill "
        "py_comparison with the PY benchmark comparison (or why you did not "
        "rely on the PY), and put every inference in judgment_notes.",
        SubmitConclusionInput,
    ),
]

_INPUT_MODELS: dict[str, Any] = {
    "search_cy_support": SearchCySupportInput,
    "locate_value": LocateValueInput,
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


def _last_year_facts_reported(conclusion: SubmitConclusionInput, ctx: ToolContext) -> list[str]:
    """Dates, numbers and names the conclusion reports as OBSERVED that come
    from last year's workpaper and appear nowhere in this year's evidence,
    sample listing or test step.

    The citation check only proves the cited evidence_ids are this year's;
    nothing checked the VALUES. A real run reported last year's items as
    if they had been tested this year. py_comparison is exempt -- it is
    where last year's facts legitimately belong.
    """
    if not ctx.py_text.strip():
        return []
    from agent.mirror.verify import _dates, _names, _numbers

    cy = "\n".join([_searchable_text(i) for i in ctx.evidence_items] + [ctx.cy_reference_text])
    cy_dates, cy_numbers, cy_lower = _dates(cy), _numbers(cy), cy.lower()
    py = ctx.py_text
    py_dates, py_numbers, py_lower = _dates(py), _numbers(py), py.lower()

    observed = (
        [a.value_observed for a in conclusion.attribute_results]
        + [r.note for r in conclusion.sample_results]
        + [c.quote_or_summary for c in conclusion.evidence_citations]
        + [conclusion.narrative]
    )
    found: set[str] = set()
    for text in observed:
        found |= {d for d in _dates(text) if d in py_dates and d not in cy_dates}
        found |= {n for n in _numbers(text) if n in py_numbers and n not in cy_numbers}
        found |= {
            name for name in _names(text)
            if " " in name and name.lower() in py_lower and name.lower() not in cy_lower
        }
    return sorted(found)


def _coverage_of(
    conclusion: SubmitConclusionInput, required_sample_ids: list[str]
) -> "SampleCoverage | None":
    """Which required samples this conclusion actually concludes on.

    "Addressed" means the conclusion says something about that selection:
    a sample_result, an attribute_result, or a citation pinned to it. A
    sample the model searched for and then said nothing about is NOT
    covered -- the workpaper would show an empty tab for it.

    None when the step has no sample list, which is a real case (a
    step-wide procedure with nothing to select) and must not be treated as
    zero coverage.
    """
    if not required_sample_ids:
        return None
    addressed = {r.sample_id for r in conclusion.sample_results}
    addressed |= {a.sample_id for a in conclusion.attribute_results if a.sample_id}
    addressed |= {c.sample_id for c in conclusion.evidence_citations if c.sample_id}
    missing = [sid for sid in required_sample_ids if sid not in addressed]
    found = len(required_sample_ids) - len(missing)
    return SampleCoverage(
        total_required=len(required_sample_ids),
        total_found=found,
        missing=missing,
        coverage_pct=round(100.0 * found / len(required_sample_ids), 1),
        complete=not missing,
    )


def _execute_tool(
    block: Any, ctx: ToolContext, turn: int, audit_log: list[AuditLogEntry]
) -> tuple[dict, bool, ConclusionOutput | None]:
    """Returns (content_for_claude, is_error, conclusion_if_this_was_a_successful_submit)."""
    # Count every call up front, before any early return -- a call that fails
    # basic input validation is still a
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

    elif block.name == "locate_value":
        output = locate_value(parsed, ctx)
        content = output.model_dump()
        is_error = False

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

        # Coverage is COMPUTED here, not asked for.
        #
        # This used to be a tool the model called, passing in the sample_ids
        # it believed it had covered -- so the check confirmed what it was
        # told, which is the model marking its own homework with an extra
        # API call in between. The backend already knows the required
        # sample_ids (from the sample listing) and can see exactly which
        # ones the submitted conclusion actually addresses, so the honest
        # check is free, needs no turn, and tests what was CONCLUDED rather
        # than what was claimed.
        if gate_error is None:
            leaked = _last_year_facts_reported(parsed, ctx)
            if leaked:
                gate_error = (
                    f"these values come from LAST YEAR's workpaper and appear nowhere in this year's "
                    f"evidence, sample listing or test step: {leaked}. You reported them as observed this "
                    "year. Test only this year's selections against this year's support: report what the "
                    "CY evidence shows, or conclude insufficient_evidence and request what is missing. "
                    "Last year's facts belong only in py_comparison."
                )

        coverage = _coverage_of(parsed, ctx.required_sample_ids)
        ctx.last_sample_coverage = coverage
        if (
            gate_error is None
            and parsed.conclusion == "satisfied"
            and coverage is not None
            and not coverage.complete
        ):
            gate_error = (
                f"this conclusion only addresses {coverage.total_found} of "
                f"{coverage.total_required} sampled item(s) -- missing {coverage.missing}. "
                "'satisfied' is not supportable for a sample you did not conclude on. "
                "Either add sample_results and attribute_results for the missing "
                "selection(s), or conclude not_satisfied/insufficient_evidence and use "
                "request_additional_support for what could not be obtained."
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
    ledger: "TokenLedger | None" = None,
) -> tuple[ConclusionOutput, list[AuditLogEntry]]:
    """on_turn, if given, is called once per completed turn with
    (turn_number, cumulative_tokens_used, audit_log_so_far) -- purely for a
    caller (the Streamlit app) to show live progress while a step is still
    running. It sees the same audit_log list this function keeps appending
    to, so a caller that only reads it (rather than mutating it) is safe;
    nothing here depends on its return value.

    ledger, if given, gets one entry per API call with the four token kinds
    kept separate, for the after-the-run cost breakdown. It is observation
    only: the spending cap runs off this function's own tokens_used, never
    off the ledger, so nothing about cost REPORTING can abort a run.
    """
    ctx = ToolContext(
        evidence_items=evidence_items,
        model=model,
        required_sample_ids=[s.sample_id for s in sample_manifest.samples] if sample_manifest else [],
        py_text="\n".join([request.py_conclusion_text] + [_searchable_text(i) for i in request.py_support_excerpts]),
        cy_reference_text="\n".join(
            [request.control_id, request.control_objective_text, request.test_step_text]
            + [f"{s.sample_id} {s.identifying_details} " + " ".join(f"{k} {v}" for k, v in (s.key_fields or {}).items())
               for s in (sample_manifest.samples if sample_manifest else [])]
        ),
    )

    audit_log: list[AuditLogEntry] = []
    tokens_used = 0
    wrapup_warned = False
    system = [{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}]
    sections = user_turn_sections(request, cy_evidence=evidence_items)
    messages: list[dict] = [{"role": "user", "content": "".join(sections.values())}]

    if ledger is not None:
        # Everything sent on turn 1, by piece -- the system prompt and the
        # tool schemas included, since both are billed and both are things
        # we chose to send. Recorded before the call so it is present even
        # if the very first request fails.
        ledger.note_prompt_mix(
            request.test_step_id,
            {
                **{k: len(v) for k, v in sections.items() if v},
                "System prompt": len(SYSTEM_PROMPT),
                "Tool schemas": len(json.dumps(TOOLS)),
            },
        )

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

        if ledger is not None:
            # Label the turn by what it actually did, so a turn-by-turn
            # readout says "searched twice" or "submitted the conclusion"
            # rather than just naming a number.
            called = [
                b.name for b in response.content if getattr(b, "type", None) == "tool_use"
            ]
            ledger.record(
                getattr(response, "usage", None),
                group=request.test_step_id,
                label=", ".join(called) if called else "no tool call (prose only)",
                turn=turn,
            )

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
