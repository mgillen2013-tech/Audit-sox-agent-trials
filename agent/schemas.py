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
    population_description: str
    population_size: int | None = None
    sample_size: int
    selection_method: Literal["random", "haphazard", "judgmental", "all_items"]
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
    top_k: int = 5


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


class ConclusionOutput(BaseModel):
    test_step_id: str
    control_objective_ref: str
    conclusion: Literal["satisfied", "not_satisfied", "insufficient_evidence"]
    narrative: str
    evidence_citations: list[EvidenceCitation]
    procedures_performed: list[str]
    relies_on_system_generated_report: bool
    ipe_completeness_accuracy_evidence: list[str]
    exceptions: list[str]
    additional_support_requests: list[str]
    confidence: Literal["high", "medium", "low"]
    confidence_rationale: str
    sample_coverage: SampleCoverage | None
    model_metadata: ModelMetadata

    @model_validator(mode="after")
    def _citations_required_unless_insufficient(self) -> "ConclusionOutput":
        if self.conclusion != "insufficient_evidence" and not self.evidence_citations:
            raise ValueError(
                "evidence_citations must be non-empty unless conclusion == 'insufficient_evidence'"
            )
        return self

    @model_validator(mode="after")
    def _ipe_evidence_required_if_relied_on(self) -> "ConclusionOutput":
        if self.relies_on_system_generated_report and not self.ipe_completeness_accuracy_evidence:
            raise ValueError(
                "ipe_completeness_accuracy_evidence must be non-empty when "
                "relies_on_system_generated_report is true"
            )
        return self


def validate_citations_against_transcript(
    conclusion: ConclusionOutput, evidence_ids_returned_by_search: set[str]
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
