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

Sample list uploads are per test step and accept ANY spreadsheet columns
(see agent.intake.build_manifest_from_any_columns) -- a real population/
sample export straight out of E1 or wherever has columns like "invoice
number f0411.vinv", not "identifying_details". Population description,
selection method, and population size are entered directly on the form
instead of expected to live in the file, since they're audit judgments,
not something a source system export carries.

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

from agent.intake import build_manifest_from_any_columns, read_excel_rows
from agent.loop import DEFAULT_MODEL, MAX_TOOL_ITERATIONS, MAX_TOTAL_TOKENS
from agent.run_control import iter_control_results
from agent.schemas import ConclusionOutput
from agent.workpaper import build_workpaper

SELECTION_METHODS = ["random", "haphazard", "judgmental", "all_items"]

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

    st.header("Spending cap")
    max_total_tokens = st.number_input(
        "Max tokens per test step",
        min_value=10_000,
        max_value=2_000_000,
        value=MAX_TOTAL_TOKENS,
        step=10_000,
        help=(
            "Hard cutoff. If one test step's running total of input+output "
            "tokens crosses this, it's aborted rather than left to keep "
            "burning turns -- but everything it already found is still kept "
            "and shown, not thrown away. Lower this if you want to catch a "
            "runaway step (bad extraction, a huge PY file) even cheaper; "
            "raise it only if a legitimately complex step keeps hitting the "
            "cap with a conclusion still in reach."
        ),
    )

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

        st.markdown("**Sample**")
        pop_desc = st.text_input(
            "Population description",
            key=f"popdesc_{i}",
            placeholder="e.g. All AP payments issued in the test period",
        )
        sel_method = st.selectbox("Selection method", SELECTION_METHODS, key=f"selmethod_{i}")
        pop_size = st.number_input(
            "Population size (optional -- leave 0 if unknown)", key=f"popsize_{i}", min_value=0, step=1
        )
        sample_file = st.file_uploader(
            "Sample list for this step (any Excel export -- no specific columns required)",
            type=["xlsx", "xls", "xlsm"],
            key=f"samplefile_{i}",
        )

        test_steps.append(
            {
                "test_step_id": tsid,
                "test_step_text": tstext,
                "population_description": pop_desc,
                "selection_method": sel_method,
                "population_size": pop_size,
                "sample_file": sample_file,
            }
        )

# ── 3. Files ──────────────────────────────────────────────────────────────
st.header("3. Upload PY testing + CY support")
py_testing_file = st.file_uploader("PY testing workpaper", type=["pdf", "xlsx", "xls", "xlsm"])
cy_support_files = st.file_uploader(
    "CY support evidence (one or more)", type=["pdf", "xlsx", "xls", "xlsm"], accept_multiple_files=True
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


def _render_result(test_step_id: str, result: dict) -> None:
    """One test step's outcome -- success or preserved failure. Used both
    for live rendering during a run and for re-rendering from
    st.session_state afterward: Streamlit reruns this whole script on ANY
    widget click (including the workpaper download button), and anything
    rendered only inside the run-click block vanishes on that rerun.
    """
    st.subheader(test_step_id)
    if "error" in result:
        st.error(f"Failed: {result['error']}")
        if "audit_log" in result:
            # A token-budget or max-iterations abort still made real,
            # billed API calls -- show what they actually found instead of
            # a dead end with nothing to show for the spend.
            st.caption(
                f"Stopped after {result['turns_used']} turn(s), "
                f"~{result['tokens_used']:,} tokens used, "
                f"{len(result['audit_log'])} tool call(s) made before the abort."
            )
            with st.expander("Tool calls made before failure"):
                for entry in result["audit_log"]:
                    st.markdown(f"**Turn {entry.turn} — `{entry.tool_name}`**" + (" ⚠️ error" if entry.is_error else ""))
                    st.json({"input": entry.input, "output": entry.output})
    else:
        _render_conclusion(result["conclusion"])


# ── Run ───────────────────────────────────────────────────────────────────
st.header("4. Run")
run_clicked = st.button("Run test steps", type="primary")
if run_clicked:
    errors = []
    if not control_id or not control_objective_ref or not control_objective_text:
        errors.append("Fill in all of the control details.")
    for s in test_steps:
        if not s["test_step_id"] or not s["test_step_text"] or not s["population_description"]:
            errors.append(f"Fill in every field for test step {s['test_step_id'] or '(unnamed)'}.")
        if s["sample_file"] is None:
            errors.append(f"Upload a sample list for test step {s['test_step_id'] or '(unnamed)'}.")
    if py_testing_file is None:
        errors.append("Upload a PY testing workpaper.")
    if not cy_support_files:
        errors.append("Upload at least one CY support file.")
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

            sample_manifests = {}
            build_failed = False
            for s in test_steps:
                sample_path = tmp_dir / s["sample_file"].name
                sample_path.write_bytes(s["sample_file"].getvalue())
                try:
                    rows = read_excel_rows(sample_path)
                    sample_manifests[s["test_step_id"]] = build_manifest_from_any_columns(
                        rows,
                        test_step_id=s["test_step_id"],
                        population_description=s["population_description"],
                        selection_method=s["selection_method"],
                        population_size=int(s["population_size"]) or None,
                    )
                except ValueError as exc:
                    st.error(f"Sample list for {s['test_step_id']}: {exc}")
                    build_failed = True

            if build_failed:
                st.stop()

            spec = {
                "control_id": control_id,
                "control_objective_ref": control_objective_ref,
                "control_objective_text": control_objective_text,
                "py_testing_file": py_testing_file.name,
                "cy_support_files": [f.name for f in cy_support_files],
                "test_steps": [{"test_step_id": s["test_step_id"], "test_step_text": s["test_step_text"]} for s in test_steps],
            }
            client = AnthropicFoundry(api_key=api_key, resource=resource)

            # Live progress -- a step can run for a couple of minutes across
            # several tool calls, and iter_control_results only yields once a
            # step is fully done or aborted. Without this there is no
            # visibility (and no cost signal) while the money is being spent.
            progress = st.empty()

            def _show_progress(step_id: str, turn: int, tokens: int, log: list) -> None:
                progress.info(
                    f"Running **{step_id}** — turn {turn}/{MAX_TOOL_ITERATIONS}, ~{tokens:,} tokens used so far, "
                    f"{len(log)} tool call(s) made. (Cap: {int(max_total_tokens):,} tokens)"
                )

            all_results = {}
            for test_step_id, result in iter_control_results(
                spec,
                tmp_dir,
                client,
                model,
                sample_manifests=sample_manifests,
                max_total_tokens=int(max_total_tokens),
                on_turn=_show_progress,
            ):
                progress.empty()
                _render_result(test_step_id, result)
                all_results[test_step_id] = result

            # Generate the CY workpaper file (same file type as the PY
            # upload) while the temp dir still exists, and keep only the
            # BYTES -- the temp dir is gone the moment this with-block
            # closes, and the download button needs to survive reruns.
            wp_bytes = wp_name = None
            try:
                wp_path = build_workpaper(spec, all_results, py_testing_file.name, tmp_dir)
                wp_bytes, wp_name = wp_path.read_bytes(), wp_path.name
            except Exception as exc:  # noqa: BLE001 -- the run's results must still show even if the file build breaks
                st.warning(f"Couldn't generate the workpaper file: {exc}")

        st.session_state["run_output"] = {"results": all_results, "wp_bytes": wp_bytes, "wp_name": wp_name}
        st.success("Done.")

# Rendered OUTSIDE the run-click block so results and the download button
# survive Streamlit's script reruns (every widget click is a rerun -- without
# this, clicking Download would blank the whole results view).
_out = st.session_state.get("run_output")
if _out:
    if not run_clicked:  # freshly-run results were already rendered live above
        st.header("Results (last run)")
        for _tsid, _result in _out["results"].items():
            _render_result(_tsid, _result)
    if _out["wp_bytes"]:
        st.download_button(
            "⬇️ Download CY workpaper (DRAFT)",
            data=_out["wp_bytes"],
            file_name=_out["wp_name"],
            mime=(
                "application/pdf"
                if _out["wp_name"].endswith(".pdf")
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
        st.caption(
            "Same file type as the PY workpaper you uploaded. Stamped DRAFT -- "
            "review and approve before it goes anywhere."
        )
