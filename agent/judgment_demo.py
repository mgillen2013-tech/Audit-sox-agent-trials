"""Shows how the testing agent reaches a conclusion, on two made-up controls.

Both controls are invented for this demo -- no client data -- and are
chosen to exercise the two halves of the judgment rules in agent/loop.py:

  1. "CY matches PY": this year's support is the same kind of evidence, from
     the same source, as last year's, but one attribute is only shown
     indirectly (an email reply rather than the word "approved"). The agent
     should accept the reasonable inference, say so, and lean on the PY
     comparison.
  2. "CY does not match PY": the system and the owner changed, so the PY is
     no benchmark, and the CY evidence contains a red flag. The agent should
     test on CY alone, say why, and still flag the exception.

The agent decides; this script only feeds it the evidence and prints what it
decided. Run it with your Foundry settings (PowerShell):

    $env:ANTHROPIC_FOUNDRY_API_KEY="..."
    $env:ANTHROPIC_FOUNDRY_RESOURCE="your-resource"
    $env:ANTHROPIC_MODEL="your-deployment-name"
    python -m agent.judgment_demo
"""

from __future__ import annotations

import os
from dataclasses import dataclass

from agent.loop import DEFAULT_MODEL, TestStepRequest, run_test_step
from agent.schemas import EvidenceItem, SampleItem, SamplePopulationManifest


def _item(evidence_id: str, source_file: str, text: str) -> EvidenceItem:
    return EvidenceItem(
        evidence_id=evidence_id,
        source_file=source_file,
        source_type="pdf_text",
        location=f"{source_file} p.1",
        extracted_text=text,
        extraction_confidence=1.0,
        preview_ref=f"{source_file} p.1",
    )


@dataclass
class DemoControl:
    title: str
    request: TestStepRequest
    cy_evidence: list[EvidenceItem]
    manifest: SamplePopulationManifest


def _single_sample(step_id: str, sample_id: str, details: str) -> SamplePopulationManifest:
    return SamplePopulationManifest(
        test_step_id=step_id,
        population_description="Demo population",
        population_size=40,
        sample_size=1,
        samples=[SampleItem(sample_id=sample_id, test_step_id=step_id, identifying_details=details)],
    )


def demo_controls() -> list[DemoControl]:
    # ── 1. CY evidence is the same kind as PY; one attribute is inferred ──
    je_step = "1"
    je_manifest = _single_sample(je_step, "1", "Journal entry JE-2026-0388, accrual true-up, $18,250.00")
    je = DemoControl(
        title="Control 1 -- CY matches PY (manual journal entry approval)",
        request=TestStepRequest(
            test_step_id=je_step,
            control_id="DEMO.JE.1",
            control_objective_ref="DEMO.JE.1",
            control_objective_text=(
                "Manual journal entries are reviewed and approved by someone other than the "
                "preparer before they are posted."
            ),
            test_step_text=(
                "Verify the selected journal entry was approved by a reviewer independent of "
                "the preparer prior to posting."
            ),
            py_conclusion_text=(
                "Satisfied. The Controller approved the entry by email before it was posted; "
                "the approver is independent of the preparer. No exceptions noted."
            ),
            py_support_excerpts=[
                _item("py_001", "PY_JE_2025-0417.pdf",
                      "Journal entry JE-2025-0417. Prepared by: A. Rivera, Senior Accountant, "
                      "03/12/2025.\n"
                      "Email -- From: M. Chen (Controller) Sent: 03/13/2025 08:14 Subject: RE: "
                      "JE-2025-0417 for review. 'Reviewed backup, approved to post.'\n"
                      "ERP posting log: JE-2025-0417 posted 03/13/2025 10:02 by A. Rivera."),
            ],
            sample_size=je_manifest.sample_size,
            population_size=je_manifest.population_size,
            samples=list(je_manifest.samples),
        ),
        cy_evidence=[
            _item("ev_001", "JE_2026-0388_packet.pdf",
                  "Journal entry JE-2026-0388, accrual true-up, $18,250.00. Prepared by: "
                  "A. Rivera, Senior Accountant, 03/10/2026. Backup: accrual schedule attached."),
            _item("ev_002", "JE_2026-0388_email.pdf",
                  "From: A. Rivera Sent: 03/10/2026 16:20 To: M. Chen Subject: JE-2026-0388 for "
                  "review -- accrual true-up, backup attached.\n"
                  "From: M. Chen (Controller) Sent: 03/11/2026 07:55 To: A. Rivera Subject: RE: "
                  "JE-2026-0388 for review -- 'Looks good, go ahead.'"),
            _item("ev_003", "ERP_posting_log_Mar2026.pdf",
                  "ERP posting log: JE-2026-0388 posted 03/11/2026 09:40 by A. Rivera. "
                  "Amount $18,250.00."),
        ],
        manifest=je_manifest,
    )

    # ── 2. System and owner changed; the CY evidence holds a red flag ─────
    vm_step = "1"
    vm_manifest = _single_sample(vm_step, "1", "Supplier change SUP-2026-031, bank account update")
    vm = DemoControl(
        title="Control 2 -- CY does not match PY (vendor master change approval)",
        request=TestStepRequest(
            test_step_id=vm_step,
            control_id="DEMO.VM.2",
            control_objective_ref="DEMO.VM.2",
            control_objective_text=(
                "Changes to vendor master data are approved by AP management, independent of "
                "the requester, before they take effect."
            ),
            test_step_text=(
                "Verify the selected vendor master change was approved by AP management "
                "independent of the requester before it took effect."
            ),
            py_conclusion_text="Satisfied. Approved in the workflow tool by the AP Manager before the effective date.",
            py_support_excerpts=[
                _item("py_001", "PY_VendorFlow_report.pdf",
                      "Workflow tool approval report: change VM-2025-112 requested by J. Patel "
                      "(AP Specialist) 05/01/2025; approved by S. Okafor (AP Manager) "
                      "05/02/2025; effective 05/03/2025."),
            ],
            sample_size=vm_manifest.sample_size,
            population_size=vm_manifest.population_size,
            samples=list(vm_manifest.samples),
        ),
        cy_evidence=[
            _item("ev_001", "System_change_memo_2026.pdf",
                  "Effective 01/01/2026 vendor master changes are processed in the new ERP "
                  "supplier module. The previous workflow tool was retired on 12/31/2025. The "
                  "AP Manager role is now held by L. Brooks."),
            _item("ev_002", "Supplier_module_change_log.pdf",
                  "Supplier module change log: SUP-2026-031 bank account update. Requested by "
                  "L. Brooks 04/14/2026 09:12. Approved by L. Brooks 04/14/2026 09:15. "
                  "Effective 04/14/2026."),
        ],
        manifest=vm_manifest,
    )
    return [je, vm]


def report(control: DemoControl, conclusion, audit_log) -> str:
    flagged = [e.input.get("description", "") for e in audit_log if e.tool_name == "flag_exception"]
    lines = [
        "=" * 72,
        control.title,
        "=" * 72,
        f"Conclusion:       {conclusion.conclusion}  (confidence: {conclusion.confidence})",
        f"Basis:            {conclusion.conclusion_basis}",
        f"PY comparison:    {conclusion.py_comparison or '-'}",
        f"Judgment notes:   {conclusion.judgment_notes or '-'}",
        f"Narrative:        {conclusion.narrative}",
        "Exceptions flagged:" + ("" if flagged else " none"),
        *[f"  - {d}" for d in flagged],
        "Attributes:",
        *[f"  - {a.attribute}: {a.result} -- {a.value_observed}" for a in conclusion.attribute_results],
        f"Tool calls: {len(audit_log)}",
    ]
    return "\n".join(lines)


def main() -> None:
    from anthropic import AnthropicFoundry  # here so the module imports without Foundry installed

    client = AnthropicFoundry()
    model = os.environ.get("ANTHROPIC_MODEL", DEFAULT_MODEL)
    print(f"Using deployment: {model}\n")
    for control in demo_controls():
        conclusion, audit_log = run_test_step(
            control.request, control.cy_evidence, control.manifest, client, model=model
        )
        print(report(control, conclusion, audit_log))
        print()


if __name__ == "__main__":
    main()
