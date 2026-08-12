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
from typing import Any, Protocol

from agent.schemas import (
    CheckSampleCoverageError,
    CheckSampleCoverageInput,
    ConclusionOutput,
    EvidenceItem,
    FlagExceptionInput,
    ModelMetadata,
    RequestAdditionalSupportInput,
    SampleCoverage,
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

DEFAULT_MODEL = "claude-sonnet-5"
PROMPT_VERSION = "v1"
MAX_TOOL_ITERATIONS = 15

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

{COMPANY_CONTEXT}"""


@dataclass
class TestStepRequest:
    __test__ = False  # not a pytest test class -- name collision with the "Test" prefix convention

    test_step_id: str
    control_id: str
    control_objective_ref: str
    control_objective_text: str
    test_step_text: str
    py_conclusion_text: str
    py_support_excerpts: list[EvidenceItem] = field(default_factory=list)


def _render_evidence_item(item: EvidenceItem) -> str:
    body = item.extracted_text or (
        "\n".join(" | ".join(row) for row in item.extracted_table) if item.extracted_table else "[no text]"
    )
    return f"[{item.evidence_id}] {item.location}\n{body}"


def build_user_turn(request: TestStepRequest) -> str:
    py_excerpts = "\n\n".join(_render_evidence_item(item) for item in request.py_support_excerpts) or (
        "(no PY support excerpts provided)"
    )
    return f"""\
Control objective ({request.control_objective_ref}): {request.control_objective_text}

Test step ({request.test_step_id}): {request.test_step_text}

PY conclusion: {request.py_conclusion_text}

PY support excerpts (format/approach precedent only -- not evidence for this
year's conclusion):
{py_excerpts}

Use search_cy_support to find this year's evidence for this test step before
concluding anything."""


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
        "insufficient_evidence.",
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


@dataclass
class AuditLogEntry:
    turn: int
    tool_name: str
    tool_use_id: str
    input: dict
    output: dict
    is_error: bool
    timestamp: str


def _execute_tool(
    block: Any, ctx: ToolContext, turn: int, audit_log: list[AuditLogEntry]
) -> tuple[dict, bool, ConclusionOutput | None]:
    """Returns (content_for_claude, is_error, conclusion_if_this_was_a_successful_submit)."""
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
        try:
            validate_citations_against_transcript(parsed, ctx.evidence_ids_returned_by_search)
        except ValueError as exc:
            content = {"error": str(exc)}
            is_error = True
        else:
            conclusion = ConclusionOutput(
                **parsed.model_dump(),
                sample_coverage=ctx.last_sample_coverage,
                model_metadata=ModelMetadata(
                    model=ctx.model,
                    prompt_version=PROMPT_VERSION,
                    timestamp=_now(),
                    tool_call_count=ctx.tool_call_count + 1,
                ),
            )
            content = {"recorded": True}
            is_error = False
    else:  # pragma: no cover -- guarded by the model_cls lookup above
        content = {"error": f"unhandled tool {block.name!r}"}
        is_error = True

    ctx.tool_call_count += 1
    audit_log.append(AuditLogEntry(turn, block.name, block.id, dict(block.input), content, is_error, _now()))
    return content, is_error, conclusion


def _now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def run_test_step(
    request: TestStepRequest,
    evidence_items: list[EvidenceItem],
    sample_manifest: SamplePopulationManifest | None,
    client: "AnthropicClientLike",
    model: str = DEFAULT_MODEL,
    max_iterations: int = MAX_TOOL_ITERATIONS,
) -> tuple[ConclusionOutput, list[AuditLogEntry]]:
    ctx = ToolContext(
        evidence_items=evidence_items,
        model=model,
        required_sample_ids=[s.sample_id for s in sample_manifest.samples] if sample_manifest else [],
    )

    audit_log: list[AuditLogEntry] = []
    system = [{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}]
    messages: list[dict] = [{"role": "user", "content": build_user_turn(request)}]

    for turn in range(1, max_iterations + 1):
        response = _call_model(
            client,
            model=model,
            max_tokens=4096,
            system=system,
            tools=TOOLS,
            messages=messages,
        )

        if getattr(response, "stop_reason", None) == "refusal":
            raise RuntimeError(f"model declined the request: {getattr(response, 'stop_details', None)}")

        messages.append({"role": "assistant", "content": response.content})
        tool_use_blocks = [b for b in response.content if getattr(b, "type", None) == "tool_use"]

        if not tool_use_blocks:
            messages.append(
                {
                    "role": "user",
                    "content": (
                        "Every turn must end in a tool call. Use search_cy_support to "
                        "keep gathering evidence, or submit_conclusion when ready."
                    ),
                }
            )
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

        messages.append({"role": "user", "content": tool_results})

        if final_conclusion is not None:
            return final_conclusion, audit_log

    raise RuntimeError(
        f"test step {request.test_step_id!r} did not reach submit_conclusion "
        f"within {max_iterations} tool-loop iterations"
    )
