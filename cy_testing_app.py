"""Local web form for the CY testing agent -- fill in a control, upload
files, run test steps, see results, no JSON files or PowerShell commands.

Run with:
    streamlit run cy_testing_app.py

This is a UI shell over already-tested code (agent/extraction,
agent/intake, agent/loop, agent/run_control) -- it stages uploaded files to
a temp folder and a typed-in sample table to a temp Excel file, then calls
the exact same iter_control_results() the command-line runner uses, so the
orchestration logic here is the same logic proven by the CLI runs, not a
second implementation of it.

This is still a shell, not the full review UI from the design doc (no
approve/edit/reject, no citation-card source previews) -- it's the intake
+ run + raw-results half, so you can drive a real control without hand-
building a control.json.
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st
from anthropic import AnthropicFoundry

from agent.loop import DEFAULT_MODEL
from agent.run_control import iter_control_results
from agent.schemas import ConclusionOutput

SAMPLE_LIST_COLUMNS = [
    "test_step_id",
    "sample_id",
    "identifying_details",
    "population_description",
    "selection_method",
    "population_size",
]

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
        pyconc = st.text_area(
            "PY conclusion", key=f"pyconc_{i}", placeholder="What last year's workpaper concluded for this step."
        )
        test_steps.append({"test_step_id": tsid, "test_step_text": tstext, "py_conclusion_text": pyconc})

# ── 3. Files ──────────────────────────────────────────────────────────────
st.header("3. Upload files")
py_testing_file = st.file_uploader("PY testing workpaper", type=["pdf", "xlsx", "xls", "xlsm"])
cy_support_files = st.file_uploader(
    "CY support evidence (one or more)", type=["pdf", "xlsx", "xls", "xlsm"], accept_multiple_files=True
)

# ── 4. Sample list ───────────────────────────────────────────────────────────
st.header("4. Sample list")
st.caption(
    "One row per sampled item. test_step_id must match one of the Test step IDs above. "
    "Add rows with the + at the bottom of the table."
)
empty_samples = pd.DataFrame(columns=SAMPLE_LIST_COLUMNS)
sample_df = st.data_editor(
    empty_samples,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "selection_method": st.column_config.SelectboxColumn(
            options=["random", "haphazard", "judgmental", "all_items"]
        ),
        "population_size": st.column_config.NumberColumn(min_value=0, step=1),
    },
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


def _write_sample_list_xlsx(df: pd.DataFrame, path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))
    wb.save(path)


# ── Run ───────────────────────────────────────────────────────────────────
st.header("5. Run")
if st.button("Run test steps", type="primary"):
    errors = []
    if not control_id or not control_objective_ref or not control_objective_text:
        errors.append("Fill in all of the control details.")
    step_ids = [s["test_step_id"] for s in test_steps]
    if any(not s["test_step_id"] or not s["test_step_text"] or not s["py_conclusion_text"] for s in test_steps):
        errors.append("Fill in every field for every test step.")
    if py_testing_file is None:
        errors.append("Upload a PY testing workpaper.")
    if not cy_support_files:
        errors.append("Upload at least one CY support file.")
    if sample_df.empty:
        errors.append("Add at least one row to the sample list.")
    else:
        unmatched = set(sample_df["test_step_id"].dropna()) - set(step_ids)
        if unmatched:
            errors.append(f"Sample list references test_step_id(s) not defined above: {sorted(unmatched)}")
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
            _write_sample_list_xlsx(sample_df, tmp_dir / "samples.xlsx")

            spec = {
                "control_id": control_id,
                "control_objective_ref": control_objective_ref,
                "control_objective_text": control_objective_text,
                "py_testing_file": py_testing_file.name,
                "sample_list_file": "samples.xlsx",
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
