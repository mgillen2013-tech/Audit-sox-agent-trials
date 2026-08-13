"""Local web form for the CY testing agent -- fill in a control, upload
files, run test steps, see results, no JSON files or PowerShell commands.

Run with:
    streamlit run cy_testing_app.py
    (or, if that command isn't found: python3 -m streamlit run cy_testing_app.py)

This is a UI shell over already-tested code (agent/extraction,
agent/intake, agent/loop, agent/run_control) -- it stages uploaded files to
a temp folder, then calls the exact same iter_control_results() the
command-line runner uses, so the orchestration logic here is the same logic
proven by the CLI runs, not a second implementation of it.

This is still a shell, not the full review UI from the design doc (no
approve/edit/reject, no citation-card source previews) -- it's the intake
+ run + raw-results half, so you can drive a real control without hand-
building a control.json.
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

import streamlit as st
from anthropic import AnthropicFoundry

from agent.intake import parse_sample_list
from agent.loop import DEFAULT_MODEL
from agent.run_control import iter_control_results
from agent.schemas import ConclusionOutput

st.set_page_config(page_title="CY Testing Agent", page_icon="🧾", layout="wide")
st.title("🧾 CY Testing Agent")
st.caption("Upload PY testing + CY support for one control, run it, review draft conclusions.")

# ── Sidebar: Foundry connection ─────────────────────────────────────────────
with st.sidebar:
    st.header("Foundry connection")
    api_key = st.text_input(
        "API key",
        type="password",
        value=os.environ.get("ANTHROPIC_FOUNDRY_API_KEY", ""),
        help="From your Foundry resource's Keys and Endpoint page.",
    )
    resource = st.text_input(
        "Resource name",
        value=os.environ.get("ANTHROPIC_FOUNDRY_RESOURCE", ""),
        help='The subdomain in your endpoint URL, e.g. "example-resource".',
    )
    model = st.text_input(
        "Model",
        value=os.environ.get("ANTHROPIC_MODEL", DEFAULT_MODEL),
        help="Must match a model actually deployed on this resource.",
    )
    st.caption("These are only used for this session -- nothing is saved to disk.")

# ── 1. Control details ───────────────────────────────────────────────────────
st.header("1. Control details")
col1, col2 = st.columns(2)
with col1:
    control_id = st.text_input("Control ID", placeholder="C-14")
with col2:
    control_objective_ref = st.text_input("Control objective reference", placeholder="CO-4")
control_objective_text = st.text_area(
    "Control objective", placeholder="What this control is supposed to accomplish."
)

# ── 2. Test steps ────────────────────────────────────────────────────────────
st.header("2. Test steps")
num_steps = st.number_input("How many test steps does this control have?", min_value=1, max_value=10, value=1, step=1)

test_steps = []
for i in range(int(num_steps)):
    with st.expander(f"Test step {i + 1}", expanded=(i == 0)):
        tsid = st.text_input("Test step ID", key=f"tsid_{i}", placeholder=f"TS-{i + 1}")
        tstext = st.text_area("Test step text", key=f"tstext_{i}", placeholder="What this step requires you to test.")
        test_steps.append({"test_step_id": tsid, "test_step_text": tstext})

# ── 3. Files ──────────────────────────────────────────────────────────────
st.header("3. Upload files")
py_testing_file = st.file_uploader("PY testing workpaper", type=["pdf", "xlsx", "xls", "xlsm"])
cy_support_files = st.file_uploader(
    "CY support evidence (one or more)", type=["pdf", "xlsx", "xls", "xlsm"], accept_multiple_files=True
)
sample_list_file = st.file_uploader(
    "Sample list (Excel)",
    type=["xlsx", "xls", "xlsm"],
    help=(
        "Required columns: test_step_id, sample_id, identifying_details, "
        "population_description, selection_method (random/haphazard/judgmental/all_items). "
        "Optional: population_size, plus any other columns you want kept per sample."
    ),
)


def _render_conclusion(conclusion: ConclusionOutput) -> None:
    banner = {"satisfied": st.success, "not_satisfied": st.error, "insufficient_evidence": st.warning}
    banner[conclusion.conclusion](f"**{conclusion.conclusion.upper()}** — confidence: {conclusion.confidence}")
    st.write(conclusion.narrative)

    if conclusion.evidence_citations:
        st.markdown("**Evidence cited:**")
        for c in conclusion.evidence_citations:
            st.markdown(f"- `{c.evidence_id}` — {c.source_file} ({c.location}): {c.quote_or_summary}")

    if conclusion.exceptions:
        st.markdown(f"**Exceptions flagged:** {', '.join(conclusion.exceptions)}")
    if conclusion.additional_support_requests:
        st.markdown("**Additional support requested:**")
        for req in conclusion.additional_support_requests:
            st.markdown(f"- {req}")

    if conclusion.sample_coverage:
        sc = conclusion.sample_coverage
        st.caption(f"Sample coverage: {sc.total_found}/{sc.total_required} ({sc.coverage_pct}%)")
    st.caption(
        f"IPE status: {conclusion.ipe_completeness_accuracy_status} · "
        f"tool calls: {conclusion.model_metadata.tool_call_count}"
    )

    with st.expander("Raw JSON"):
        st.json(conclusion.model_dump())


# ── Run ───────────────────────────────────────────────────────────────────
st.header("4. Run")
if st.button("Run test steps", type="primary"):
    errors = []
    if not control_id or not control_objective_ref or not control_objective_text:
        errors.append("Fill in all of the control details.")
    step_ids = [s["test_step_id"] for s in test_steps]
    if any(not s["test_step_id"] or not s["test_step_text"] for s in test_steps):
        errors.append("Fill in every field for every test step.")
    if py_testing_file is None:
        errors.append("Upload a PY testing workpaper.")
    if not cy_support_files:
        errors.append("Upload at least one CY support file.")
    if sample_list_file is None:
        errors.append("Upload a sample list.")
    if not api_key or not resource:
        errors.append("Fill in the Foundry API key and resource name in the sidebar.")

    if errors:
        for e in errors:
            st.error(e)
    else:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            (tmp_dir / py_testing_file.name).write_bytes(py_testing_file.getvalue())
            for f in cy_support_files:
                (tmp_dir / f.name).write_bytes(f.getvalue())
            (tmp_dir / sample_list_file.name).write_bytes(sample_list_file.getvalue())

            try:
                manifests = parse_sample_list(tmp_dir / sample_list_file.name)
            except ValueError as exc:
                st.error(f"Couldn't read the sample list: {exc}")
                st.stop()

            unmatched = set(manifests) - set(step_ids)
            if unmatched:
                st.warning(
                    f"Sample list has test_step_id(s) not defined above (ignored): {sorted(unmatched)}"
                )
            missing = set(step_ids) - set(manifests)
            if missing:
                st.warning(
                    f"No sample list rows for test_step_id(s) {sorted(missing)} -- "
                    f"those steps will run with an empty sample."
                )

            spec = {
                "control_id": control_id,
                "control_objective_ref": control_objective_ref,
                "control_objective_text": control_objective_text,
                "py_testing_file": py_testing_file.name,
                "sample_list_file": sample_list_file.name,
                "cy_support_files": [f.name for f in cy_support_files],
                "test_steps": test_steps,
            }
            client = AnthropicFoundry(api_key=api_key, resource=resource)

            all_results = {}
            for test_step_id, result in iter_control_results(spec, tmp_dir, client, model):
                st.subheader(test_step_id)
                if "error" in result:
                    st.error(f"Failed: {result['error']}")
                else:
                    _render_conclusion(result["conclusion"])
                all_results[test_step_id] = result

        st.success("Done.")
