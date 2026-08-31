"""Typed schemas for the CY testing agent, matching docs/cy_testing_agent_design.md.

These are the data shapes that cross the extraction -> tool-loop -> review-UI
boundaries. Keep this file and the design doc in sync -- if a field changes
here, update the doc's JSON/pseudocode blocks (and vice versa).
"""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field, model_validator

# --------------------------------------------------------------------------
# Section 1 -- Intake requirements: CY sample list
# --------------------------------------------------------------------------


class SampleItem(BaseModel):
    sample_id: str  # "S01".."S25", stable within a test step
    test_step_id: str
    identifying_details: str  # e.g. "PO #48213, Branch 35210, $12,450, 3/14/2026"
    key_fields: dict[str, str] | None = None


class SamplePopulationManifest(BaseModel):
    test_step_id: str
    population_description: str = ""
    population_size: int | None = None
    sample_size: int
    # Optional: the intake UI no longer forces a human to type this (it was
    # never actually surfaced to the model or the workpaper -- see loop.py's
    # build_user_turn), so it's honest to make it match what's actually
    # collected rather than a required field with a fake-good value behind
    # it. Still validated against the known set when it IS given.
    selection_method: Literal["random", "haphazard", "judgmental", "all_items"] | None = None
    samples: list[SampleItem]

    @model_validator(mode="after")
    def _sample_count_matches(self) -> "SamplePopulationManifest":
        if len(self.samples) != self.sample_size:
            raise ValueError(
                f"sample_size={self.sample_size} does not match "
                f"len(samples)={len(self.samples)} for test_step_id={self.test_step_id!r}"
            )
        for item in self.samples:
            if item.test_step_id != self.test_step_id:
                raise ValueError(
                    f"sample {item.sample_id!r} has test_step_id={item.test_step_id!r}, "
                    f"expected {self.test_step_id!r}"
                )
        return self


# --------------------------------------------------------------------------
# Section 2 -- Document extraction
# --------------------------------------------------------------------------

EvidenceSourceType = Literal["excel_table", "excel_cell", "pdf_text", "pdf_table", "image_ocr"]


class EvidenceItem(BaseModel):
    evidence_id: str  # stable id, e.g. "ev_003"
    source_file: str  # original filename
    source_type: EvidenceSourceType
    location: str  # "Recon.xlsx!Sheet1!B12:F20" or "support.pdf p.3 (bbox 120,340,500,600)"
    extracted_text: str | None = None
    extracted_table: list[list[str]] | None = None
    extraction_confidence: float = Field(ge=0.0, le=1.0)
    preview_ref: str


# --------------------------------------------------------------------------
# Section 3 -- Tool schemas
# --------------------------------------------------------------------------


class SearchCySupportInput(BaseModel):
    query: str
    evidence_types: list[EvidenceSourceType] | None = None
    # Bounded: results stay in the conversation for the rest of the run, so
    # an unbounded top_k let one broad search park several thousand tokens
    # in history permanently. A real run asked for 8.
    top_k: int = Field(5, ge=1, le=10)


class SearchResult(BaseModel):
    evidence_id: str
    source_file: str
    location: str
    snippet: str
    score: float


class SearchCySupportOutput(BaseModel):
    results: list[SearchResult]


class CheckSampleCoverageInput(BaseModel):
    required_sample_ids: list[str]
    found_evidence_ids: list[str]


class CheckSampleCoverageOutput(BaseModel):
    total_required: int
    total_found: int
    missing: list[str]
    coverage_pct: float
    complete: bool


class CheckSampleCoverageError(BaseModel):
    error: Literal["no_sample_list_found"]


class FlagExceptionInput(BaseModel):
    test_step_id: str
    description: str
    evidence_ids: list[str]
    severity: Literal["low", "medium", "high"]


class FlagExceptionOutput(BaseModel):
    exception_id: str
    recorded: Literal[True] = True


class RequestAdditionalSupportInput(BaseModel):
    test_step_id: str
    description: str
    reason: Literal["missing", "illegible", "ambiguous", "insufficient_sample"]


class RequestAdditionalSupportOutput(BaseModel):
    request_id: str
    recorded: Literal[True] = True


# --------------------------------------------------------------------------
# Section 4 -- Conclusion output schema
# --------------------------------------------------------------------------


class EvidenceCitation(BaseModel):
    evidence_id: str
    source_file: str
    location: str
    quote_or_summary: str
    relevance: str
    # Which sampled item this evidence supports, when the test step has a
    # sample. This is what lets the workpaper document each selection on its
    # own sheet with its own tickmarks -- a reviewer clearing selection 2
    # wants selection 2's evidence, not one merged list. Left None for
    # evidence that applies to the step as a whole (a policy, the population
    # extract, IPE parameters) rather than to one item.
    sample_id: str | None = None


class AttributeResult(BaseModel):
    """One testable attribute of the test step, for one sampled item.

    A test step is rarely one assertion -- "approved by appropriate
    personnel prior to payment" is really four: the right approver, their
    authority, the approval itself, and its date against the payment date.
    The narrative explains all of it in prose, but a reviewer signing off
    has to read the whole paragraph to find where any single attribute was
    satisfied. This is that, one row at a time.
    """

    attribute: str  # what is being verified, e.g. "Approval prior to payment date"
    sample_id: str | None = None  # None = holds for the step as a whole
    result: Literal["satisfied", "not_satisfied", "not_tested"]
    value_observed: str  # what the evidence actually showed, e.g. "Approved 10/1/2025; paid 11/10/2025"
    # Which cited evidence proves it -- this is what ties an attribute to a
    # tickmark and, through it, to the boxed value on the exhibit page.
    evidence_ids: list[str] = []


class SampleResult(BaseModel):
    """How the test step came out for ONE sampled item."""

    sample_id: str
    conclusion: Literal["satisfied", "not_satisfied", "insufficient_evidence"]
    # Short -- why this item differs from the others, or what is missing for
    # it. The full reasoning belongs in the step's narrative.
    note: str = ""


class SampleCoverage(BaseModel):
    total_required: int
    total_found: int
    missing: list[str]
    coverage_pct: float
    complete: bool


class ModelMetadata(BaseModel):
    model: str
    prompt_version: str
    timestamp: str
    tool_call_count: int


class _ConclusionCore(BaseModel):
    """Fields the model itself decides. Shared by the submit_conclusion tool
    input (what Claude sends) and ConclusionOutput (what gets stored) so the
    validation rules live in exactly one place and apply to both.

    sample_coverage and model_metadata are deliberately NOT here -- they're
    backend-computed (sample_coverage from the last check_sample_coverage
    call, model_metadata from the loop's own bookkeeping) and appended after
    validation, not supplied by the model. Asking the model to restate them
    would just be another thing it could get wrong or fabricate.
    """

    test_step_id: str
    control_objective_ref: str
    conclusion: Literal["satisfied", "not_satisfied", "insufficient_evidence"]
    narrative: str
    evidence_citations: list[EvidenceCitation]
    procedures_performed: list[str]
    # Per-item verdicts. The step-level `conclusion` above is a roll-up and
    # cannot say WHICH selection failed -- but an exception is always about
    # a specific item, and that is the first thing a reviewer asks. A step
    # can be not_satisfied because one of five items failed while the other
    # four passed, and only this field can express that.
    sample_results: list["SampleResult"] = []
    # The test step broken into its individual attributes, per sampled item.
    # This is what lets a reviewer see WHERE the step is satisfied without
    # reading the whole narrative to find it.
    attribute_results: list["AttributeResult"] = []
    # Tri-state, not a bool + a maybe-empty list. A real run surfaced exactly
    # the case a bool can't express cleanly: the model relied on a
    # system-generated report (E1/BVE1/OneStream) but explicitly could NOT
    # obtain completeness/accuracy support for it. The old boolean forced
    # ipe_completeness_accuracy_evidence to be non-empty whenever reliance was
    # true, so the model was pushed into citing the report itself just to
    # satisfy the schema, then had to write a narrative disclaimer explaining
    # that citation wasn't real IPE validation. "not_validated" says that
    # directly, as a first-class, equally-valid answer -- same principle as
    # insufficient_evidence not being a second-class conclusion.
    ipe_completeness_accuracy_status: Literal["validated", "not_validated", "not_applicable"]
    ipe_completeness_accuracy_evidence: list[str]
    exceptions: list[str]
    additional_support_requests: list[str]
    confidence: Literal["high", "medium", "low"]
    confidence_rationale: str

    @model_validator(mode="after")
    def _citations_required_unless_insufficient(self) -> "_ConclusionCore":
        if self.conclusion != "insufficient_evidence" and not self.evidence_citations:
            raise ValueError(
                "evidence_citations must be non-empty unless conclusion == 'insufficient_evidence'"
            )
        return self

    @model_validator(mode="after")
    def _ipe_evidence_matches_status(self) -> "_ConclusionCore":
        if self.ipe_completeness_accuracy_status == "validated" and not self.ipe_completeness_accuracy_evidence:
            raise ValueError(
                "ipe_completeness_accuracy_evidence must be non-empty when "
                "ipe_completeness_accuracy_status == 'validated'"
            )
        if self.ipe_completeness_accuracy_status != "validated" and self.ipe_completeness_accuracy_evidence:
            raise ValueError(
                "ipe_completeness_accuracy_evidence must be empty when "
                "ipe_completeness_accuracy_status != 'validated' -- use 'not_validated' "
                "to say reliance exists but wasn't validated, not a citation of the report itself"
            )
        return self


class SubmitConclusionInput(_ConclusionCore):
    """The submit_conclusion tool's input_schema -- what Claude actually sends."""


class ConclusionOutput(_ConclusionCore):
    """The full stored/audited record: model fields + backend-computed fields."""

    sample_coverage: SampleCoverage | None
    model_metadata: ModelMetadata


def validate_citations_against_transcript(
    conclusion: SubmitConclusionInput | ConclusionOutput, evidence_ids_returned_by_search: set[str]
) -> None:
    """Reject a conclusion that cites evidence never returned by search_cy_support
    in this conversation -- the mechanical fabrication check from section 3.
    """
    cited = {c.evidence_id for c in conclusion.evidence_citations}
    cited |= set(conclusion.ipe_completeness_accuracy_evidence)
    fabricated = cited - evidence_ids_returned_by_search
    if fabricated:
        raise ValueError(
            f"conclusion cites evidence_id(s) never returned by search_cy_support "
            f"in this transcript: {sorted(fabricated)}"
        )
