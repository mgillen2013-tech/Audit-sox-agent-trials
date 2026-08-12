"""Smoke test: run one test step against the real Claude API.

This is NOT the app -- there's no intake form, no file upload, no review
UI yet. It's the minimal script that proves the tool loop in agent/loop.py
actually works end to end against a live model, using a couple of
hand-built EvidenceItems instead of running the extraction pipeline first.

To run this:
    1. pip install -r requirements.txt
    2. Set your API key: export ANTHROPIC_API_KEY=sk-ant-...
       (or your Microsoft Foundry credentials -- swap the client
       constructor below for AnthropicFoundry(...) if you're calling
       through Foundry instead of the first-party API)
    3. python3 -m agent.example_run
"""

from __future__ import annotations

import anthropic

from agent.loop import TestStepRequest, run_test_step
from agent.schemas import EvidenceItem, SampleItem, SamplePopulationManifest


def main() -> None:
    client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from the environment

    py_support = [
        EvidenceItem(
            evidence_id="py_ev_001",
            source_file="PY_Testing_C14.pdf",
            source_type="pdf_text",
            location="PY_Testing_C14.pdf p.2",
            extracted_text=(
                "Recalculated the October PY accrual: $455,200 recalculated vs. "
                "$455,200 per GL. No variance noted."
            ),
            extraction_confidence=1.0,
            preview_ref="PY_Testing_C14.pdf p.2",
        )
    ]

    cy_evidence = [
        EvidenceItem(
            evidence_id="ev_001",
            source_file="Recon_Oct2026.xlsx",
            source_type="excel_cell",
            location="Sheet1!C4",
            extracted_text=(
                "Accrual recalculation: $482,110 recalculated vs. $482,110 per "
                "GL export (E1 report FR-118). No variance."
            ),
            extraction_confidence=1.0,
            preview_ref="Recon_Oct2026.xlsx!Sheet1!C4",
        ),
        EvidenceItem(
            evidence_id="ev_002",
            source_file="GL_Export_Oct2026.pdf",
            source_type="pdf_text",
            location="GL_Export_Oct2026.pdf p.1",
            extracted_text="E1 report FR-118: Accrued interest balance, October 2026: $482,110.",
            extraction_confidence=1.0,
            preview_ref="GL_Export_Oct2026.pdf p.1",
        ),
    ]

    sample_manifest = SamplePopulationManifest(
        test_step_id="TS-4.2",
        population_description="Month-end accrual reconciliation, October 2026",
        sample_size=1,
        selection_method="all_items",
        samples=[
            SampleItem(
                sample_id="S01",
                test_step_id="TS-4.2",
                identifying_details="October 2026 accrued interest reconciliation",
            )
        ],
    )

    request = TestStepRequest(
        test_step_id="TS-4.2",
        control_id="C-14",
        control_objective_ref="CO-4",
        control_objective_text="Accruals are recorded completely and accurately each month-end.",
        test_step_text=(
            "Recalculate the month-end accrual and agree the recalculated amount to the GL."
        ),
        py_conclusion_text="Satisfied. Recalculation agreed to GL with no variance noted.",
        py_support_excerpts=py_support,
    )

    conclusion, audit_log = run_test_step(request, cy_evidence, sample_manifest, client)

    print("=== Conclusion ===")
    print(conclusion.model_dump_json(indent=2))
    print(f"\n=== Audit log ({len(audit_log)} tool calls) ===")
    for entry in audit_log:
        print(f"  [{entry.tool_name}] error={entry.is_error} input={entry.input}")


if __name__ == "__main__":
    main()
