"""Validation-rule tests for the schemas -- these are the mechanical checks
the design doc calls "a hard, mechanical check rather than a hope": citation
requirements, the IPE-evidence requirement, the fabrication guard, and the
sample-list self-consistency check.
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from agent.schemas import (
    ConclusionOutput,
    EvidenceCitation,
    ModelMetadata,
    SampleCoverage,
    SampleItem,
    SamplePopulationManifest,
    validate_citations_against_transcript,
)


def _base_conclusion(**overrides) -> dict:
    base = dict(
        test_step_id="TS-4.2",
        control_objective_ref="CO-4",
        conclusion="satisfied",
        narrative="Recalculated accrual ties to GL with no variance.",
        evidence_citations=[
            EvidenceCitation(
                evidence_id="ev_012",
                source_file="Recon.xlsx",
                location="Sheet1!C4:C18",
                quote_or_summary="Accrual recalculation, ending balance $482,110",
                relevance="Recalculated figure used to test the accrual",
            )
        ],
        procedures_performed=["recalculation"],
        relies_on_system_generated_report=False,
        ipe_completeness_accuracy_evidence=[],
        exceptions=[],
        additional_support_requests=[],
        confidence="high",
        confidence_rationale="Native-text source, full coverage.",
        sample_coverage=SampleCoverage(
            total_required=25, total_found=25, missing=[], coverage_pct=100.0, complete=True
        ),
        model_metadata=ModelMetadata(
            model="claude-sonnet-5", prompt_version="v3", timestamp="2026-08-12T00:00:00Z", tool_call_count=6
        ),
    )
    base.update(overrides)
    return base


# --------------------------------------------------------------------------
# ConclusionOutput
# --------------------------------------------------------------------------


def test_satisfied_conclusion_requires_citations():
    with pytest.raises(ValidationError, match="evidence_citations must be non-empty"):
        ConclusionOutput(**_base_conclusion(evidence_citations=[]))


def test_insufficient_evidence_allows_empty_citations():
    conclusion = ConclusionOutput(
        **_base_conclusion(conclusion="insufficient_evidence", evidence_citations=[], sample_coverage=None)
    )
    assert conclusion.evidence_citations == []


def test_ipe_reliance_requires_evidence():
    with pytest.raises(ValidationError, match="ipe_completeness_accuracy_evidence must be non-empty"):
        ConclusionOutput(
            **_base_conclusion(relies_on_system_generated_report=True, ipe_completeness_accuracy_evidence=[])
        )


def test_ipe_reliance_with_evidence_is_valid():
    conclusion = ConclusionOutput(
        **_base_conclusion(
            relies_on_system_generated_report=True, ipe_completeness_accuracy_evidence=["ev_009"]
        )
    )
    assert conclusion.ipe_completeness_accuracy_evidence == ["ev_009"]


def test_fabricated_citation_rejected():
    conclusion = ConclusionOutput(**_base_conclusion())
    with pytest.raises(ValueError, match="never returned by search_cy_support"):
        validate_citations_against_transcript(conclusion, evidence_ids_returned_by_search=set())


def test_citation_returned_by_search_passes():
    conclusion = ConclusionOutput(**_base_conclusion())
    # Should not raise.
    validate_citations_against_transcript(conclusion, evidence_ids_returned_by_search={"ev_012"})


def test_fabricated_ipe_citation_rejected():
    conclusion = ConclusionOutput(
        **_base_conclusion(
            relies_on_system_generated_report=True, ipe_completeness_accuracy_evidence=["ev_999"]
        )
    )
    with pytest.raises(ValueError, match="ev_999"):
        validate_citations_against_transcript(conclusion, evidence_ids_returned_by_search={"ev_012"})


# --------------------------------------------------------------------------
# SamplePopulationManifest
# --------------------------------------------------------------------------


def test_sample_size_must_match_sample_list_length():
    with pytest.raises(ValidationError, match="does not match"):
        SamplePopulationManifest(
            test_step_id="TS-4.2",
            population_description="All POs > $5,000 issued Oct 2025-Sep 2026",
            sample_size=2,
            selection_method="random",
            samples=[
                SampleItem(sample_id="S01", test_step_id="TS-4.2", identifying_details="PO-48213"),
            ],
        )


def test_sample_item_test_step_id_must_match_manifest():
    with pytest.raises(ValidationError, match="expected 'TS-4.2'"):
        SamplePopulationManifest(
            test_step_id="TS-4.2",
            population_description="All POs > $5,000 issued Oct 2025-Sep 2026",
            sample_size=1,
            selection_method="random",
            samples=[
                SampleItem(sample_id="S01", test_step_id="TS-4.3", identifying_details="PO-48213"),
            ],
        )


def test_valid_manifest_round_trips():
    manifest = SamplePopulationManifest(
        test_step_id="TS-4.2",
        population_description="All POs > $5,000 issued Oct 2025-Sep 2026",
        population_size=340,
        sample_size=1,
        selection_method="random",
        samples=[
            SampleItem(
                sample_id="S01",
                test_step_id="TS-4.2",
                identifying_details="PO #48213, Branch 35210, $12,450, 3/14/2026",
                key_fields={"po_number": "48213", "branch": "35210"},
            ),
        ],
    )
    assert manifest.samples[0].key_fields["branch"] == "35210"
