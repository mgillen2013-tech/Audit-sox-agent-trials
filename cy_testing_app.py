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

One workbook covers both the population and the sample selections for a
control -- the common real shape (one file, one tab per concept) rather
than two separate uploads. Its sheets accept ANY columns (see
agent.intake.build_manifest_from_any_columns): a real export straight out
of E1 has columns like "invoice number f0411.vinv", not
"identifying_details". The whole workbook is also added to CY support
evidence, so the population tab is searchable and can back an IPE
completeness/accuracy conclusion, not just the sample tab.

Selection method and population description are no longer collected as
manual fields -- they were never actually surfaced to the model or the
workpaper, and forcing an answer just produced noise. Population size,
when a population tab is picked, is computed by counting that tab's rows
rather than typed -- and it IS passed to the model now, specifically so it
can tell a genuinely complete small sample from a partial one instead of
guessing off a support filename (see agent.loop.build_user_turn).

This is still a shell, not the full review UI from the design doc (no
approve/edit/reject, no citation-card source previews) -- it's the intake
+ run + raw-results half, so you can drive a real control without hand-
building a control.json.
"""

from __future__ import annotations

import os
import tempfile
from io import BytesIO
from pathlib import Path

import openpyxl
import streamlit as st
from anthropic import AnthropicFoundry

from agent.costs import TokenLedger, prices_for
from agent.intake import build_manifest_from_any_columns, read_excel_rows
from agent.loop import DEFAULT_MODEL, MAX_TOOL_ITERATIONS, MAX_TOTAL_TOKENS
from agent.run_control import iter_control_results
from agent.schemas import ConclusionOutput
from agent.build_cy_workpaper import build_cy_workpaper

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
        "Max cost-weighted tokens per test step",
        min_value=10_000,
        max_value=2_000_000,
        value=MAX_TOTAL_TOKENS,
        step=10_000,
        help=(
            "Hard cutoff, in cost-weighted token units that track actual "
            "dollars: cached re-reads count at ~1/10th (they bill ~10x "
            "cheaper), model output counts 5x (it bills ~5x more). This is "
            "why the number here reads lower than Foundry's raw token "
            "dashboard -- Foundry counts every cached token at face value. "
            "If a step crosses the cap it's aborted, but everything it "
            "already found is kept and shown. Lower this to catch a runaway "
            "step cheaper; raise it only if a legitimately complex step "
            "keeps hitting the cap with a conclusion still in reach."
        ),
    )
    max_run_tokens = st.number_input(
        "Max cost-weighted tokens for the whole control",
        min_value=10_000,
        max_value=8_000_000,
        value=int(max_total_tokens) * 4,
        step=50_000,
        help=(
            "Ceiling across every test step plus OCR. The per-step cap "
            "resets each step, so a 5-step control could spend 5x it with "
            "nothing watching the total -- this is what watches the total. "
            "Steps already finished keep their conclusions; only the "
            "remaining ones are skipped."
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

# ── 2. Population & sample workbook ─────────────────────────────────────────
st.header("2. Population & sample workbook")
st.caption(
    "One Excel file for the whole control -- one tab holds the full population "
    "(used to check IPE completeness/accuracy), another tab holds the items "
    "selected for testing. Only one tab? Just pick it for both, or leave "
    "population unset if you don't have one."
)
pop_sample_file = st.file_uploader(
    "Population & sample workbook", type=["xlsx", "xls", "xlsm"], key="pop_sample_file"
)

sheet_names: list[str] = []
if pop_sample_file is not None:
    try:
        _wb = openpyxl.load_workbook(BytesIO(pop_sample_file.getvalue()), read_only=True)
        sheet_names = _wb.sheetnames
        _wb.close()
    except Exception as exc:  # noqa: BLE001 -- surfaced inline, not a crash
        st.error(f"Couldn't read tabs from {pop_sample_file.name}: {exc}")

population_tab = None
if sheet_names:
    population_tab = st.selectbox(
        "Which tab is the full population?",
        ["(none)"] + sheet_names,
        key="population_tab",
        help="Used for IPE completeness/accuracy -- e.g. record counts, report parameters, tie-outs.",
    )
    population_tab = None if population_tab == "(none)" else population_tab

# ── 3. Test steps ────────────────────────────────────────────────────────────
st.header("3. Test steps")
num_steps = st.number_input("How many test steps does this control have?", min_value=1, max_value=10, value=1, step=1)

test_steps = []
for i in range(int(num_steps)):
    with st.expander(f"Test step {i + 1}", expanded=(i == 0)):
        tsid = st.text_input("Test step ID", key=f"tsid_{i}", placeholder=f"TS-{i + 1}")
        tstext = st.text_area("Test step text", key=f"tstext_{i}", placeholder="What this step requires you to test.")

        sample_tab = None
        if sheet_names:
            sample_tab = st.selectbox(
                "Which tab holds this step's sample selections?",
                sheet_names,
                key=f"sampletab_{i}",
            )
        else:
            st.info("Upload the population & sample workbook above to pick this step's sample tab.")

        test_steps.append(
            {
                "test_step_id": tsid,
                "test_step_text": tstext,
                "sample_tab": sample_tab,
            }
        )

# ── 4. Files ──────────────────────────────────────────────────────────────
st.header("4. Upload PY testing + CY support")
py_testing_file = st.file_uploader("PY testing workpaper", type=["pdf", "xlsx", "xls", "xlsm"])
cy_support_files = st.file_uploader(
    "CY support evidence (one or more)", type=["pdf", "xlsx", "xls", "xlsm"], accept_multiple_files=True
)
ocr_scanned_pages = st.checkbox(
    "Read scanned/image-only pages with vision OCR",
    value=True,
    help=(
        "Scanned PDFs (a photographed invoice, an E1 screenshot saved to PDF) "
        "carry no text at all -- without this they're flagged as unreadable and "
        "the conclusion rests on whatever else was provided. This transcribes "
        "them once per run, before testing starts. Costs roughly 1-2K tokens "
        "per scanned page; pages that already have text are never re-read."
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


def _usd(amount: float) -> str:
    """Cents are the natural unit for a single turn but the wrong one for a
    run, so scale the precision instead of rounding a real cost to $0.00 --
    a per-turn row reading zero is exactly the kind of thing that makes a
    breakdown untrustworthy.
    """
    if amount and abs(amount) < 0.10:
        return f"${amount:,.4f}"
    return f"${amount:,.2f}"


def _render_token_burn(ledger: TokenLedger) -> None:
    """What the run cost and where it went.

    The spending caps run on ONE cost-weighted number, which is the right
    shape for a circuit breaker and the wrong shape for understanding a
    bill: it cannot say whether OCR or testing spent it, which step ran
    away, or whether the context is being re-read at full price every turn.
    This is that same spend, itemised.
    """
    st.header("6. Token burn")
    if not ledger.records:
        # Every step failed before its first API call (bad credentials, an
        # unreadable upload). Nothing was billed, and empty tables would
        # read as a reporting bug rather than as "this cost nothing".
        st.info("No API calls were billed on this run.")
        if ledger.ocr_status:
            st.caption(f"OCR: {ledger.ocr_status}")
        return

    family, fresh, cw, cr, out = prices_for(ledger.model)
    total = ledger.dollars

    a, b, c = st.columns(3)
    a.metric("Estimated cost", _usd(total))
    b.metric("Tokens billed", f"{ledger.raw_tokens:,}")
    c.metric("API calls", f"{len(ledger.records):,}")
    st.caption(
        f"Priced as **{family}** at ${fresh:.2f} in / ${out:.2f} out per million tokens, "
        f"cache writes ${cw:.2f}, cache reads ${cr:.2f}. List prices — Foundry bills Claude "
        "at standard API rates, but treat this as an estimate, not an invoice."
    )

    st.markdown("**Where it went**")
    st.dataframe(
        [
            {
                "Phase": g.group,
                "Cost": _usd(g.dollars),
                "% of run": f"{(100.0 * g.dollars / total if total else 0):.1f}%",
                "API calls": g.calls,
                "Tokens": f"{g.raw_tokens:,}",
            }
            for g in ledger.group_rows()
        ],
        hide_index=True,
    )
    ocr_note = (
        "OCR is its own row because it runs before the tool loop starts — the "
        "per-step cap can neither see it nor stop it."
    )
    if ledger.ocr_status:
        ocr_note += f" **This run: {ledger.ocr_status}.**"
    st.caption(ocr_note)

    st.markdown("**What kind of tokens**")
    st.dataframe(
        [
            {"Token kind": name, "Tokens": f"{tok:,}", "Cost": _usd(usd), "% of run": f"{pct:.1f}%"}
            for name, tok, usd, pct in ledger.kind_rows()
        ],
        hide_index=True,
    )
    st.caption(
        "This is the cache-health check. On a healthy multi-turn run, **cache reads "
        "are the largest token count and one of the smallest costs** — they bill at "
        "1/10th. If fresh input is large on every turn instead, the cache is being "
        "invalidated and the run is paying full price to re-read its own context."
    )

    for g in ledger.group_rows():
        with st.expander(f"{g.group} — turn by turn ({_usd(g.dollars)})"):
            st.dataframe(
                [
                    {
                        "Turn": rec.turn or "—",
                        "What it did": rec.label,
                        "Fresh in": f"{rec.input_tokens:,}",
                        "Cache write": f"{rec.cache_write_tokens:,}",
                        "Cache read": f"{rec.cache_read_tokens:,}",
                        "Output": f"{rec.output_tokens:,}",
                        "Cost": _usd(usd),
                    }
                    for rec, usd in ledger.call_rows(g.group)
                ],
                hide_index=True,
            )
            mix = ledger.prompt_mix_rows(g.group)
            if mix:
                st.markdown("**What filled this step's first prompt**")
                st.dataframe(
                    [
                        {
                            "Section": name,
                            "Characters": f"{chars:,}",
                            "≈ Tokens": f"{tokens:,}",
                            "≈ Cost": _usd(usd),
                        }
                        for name, chars, tokens, usd in mix
                    ],
                    hide_index=True,
                )
                st.caption(
                    "Turn 1 is the expensive turn, and this is what we chose to put "
                    "in it. Apportioned across the measured turn-1 input, so the rows "
                    "sum to what was actually billed; the split between them is "
                    "proportional to length and therefore approximate."
                )

    # Deliberately NOT a tab in the workpaper: that file goes to external
    # audit, and what the draft cost to produce is internal. It is still
    # worth keeping, so it downloads on its own.
    st.download_button(
        "⬇️ Download this cost breakdown (.txt)",
        data="\n".join(ledger.report_lines()) + "\n",
        file_name="token_burn.txt",
        mime="text/plain",
    )


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
                f"~{result['tokens_used']:,} cost-weighted tokens used, "
                f"{len(result['audit_log'])} tool call(s) made before the abort."
            )
            with st.expander("Tool calls made before failure"):
                for entry in result["audit_log"]:
                    st.markdown(f"**Turn {entry.turn} — `{entry.tool_name}`**" + (" ⚠️ error" if entry.is_error else ""))
                    st.json({"input": entry.input, "output": entry.output})
    else:
        _render_conclusion(result["conclusion"])


# ── Run ───────────────────────────────────────────────────────────────────
st.header("5. Run")
run_clicked = st.button("Run test steps", type="primary")
if run_clicked:
    errors = []
    if not control_id or not control_objective_ref or not control_objective_text:
        errors.append("Fill in all of the control details.")
    if pop_sample_file is None:
        errors.append("Upload the population & sample workbook.")
    for s in test_steps:
        if not s["test_step_id"] or not s["test_step_text"]:
            errors.append(f"Fill in every field for test step {s['test_step_id'] or '(unnamed)'}.")
        if pop_sample_file is not None and not s["sample_tab"]:
            errors.append(f"Pick a sample tab for test step {s['test_step_id'] or '(unnamed)'}.")
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
            pop_sample_path = tmp_dir / pop_sample_file.name
            pop_sample_path.write_bytes(pop_sample_file.getvalue())

            population_size = None
            if population_tab:
                try:
                    population_size = len(read_excel_rows(pop_sample_path, sheet_name=population_tab))
                except Exception as exc:  # noqa: BLE001 -- surfaced inline, run still attempted
                    st.warning(f"Couldn't count rows in population tab {population_tab!r}: {exc}")

            sample_manifests = {}
            build_failed = False
            for s in test_steps:
                try:
                    rows = read_excel_rows(pop_sample_path, sheet_name=s["sample_tab"])
                    pop_desc = (
                        f"Population per '{population_tab}' tab of {pop_sample_file.name}" if population_tab else ""
                    )
                    sample_manifests[s["test_step_id"]] = build_manifest_from_any_columns(
                        rows,
                        test_step_id=s["test_step_id"],
                        population_description=pop_desc,
                        population_size=population_size,
                    )
                except ValueError as exc:
                    st.error(f"Sample tab for {s['test_step_id']}: {exc}")
                    build_failed = True

            if build_failed:
                st.stop()

            spec = {
                "control_id": control_id,
                "control_objective_ref": control_objective_ref,
                "control_objective_text": control_objective_text,
                "py_testing_file": py_testing_file.name,
                "cy_support_files": [f.name for f in cy_support_files] + [pop_sample_file.name],
                "test_steps": [{"test_step_id": s["test_step_id"], "test_step_text": s["test_step_text"]} for s in test_steps],
            }
            client = AnthropicFoundry(api_key=api_key, resource=resource)

            # Live progress -- a step can run for a couple of minutes across
            # several tool calls, and iter_control_results only yields once a
            # step is fully done or aborted. Without this there is no
            # visibility (and no cost signal) while the money is being spent.
            progress = st.empty()

            def _show_ocr(source_file: str, page_num: int, ok: bool, tokens: int) -> None:
                # OCR runs before the tool loop, so its spend sits OUTSIDE
                # the per-step cap and can't trip it -- showing the running
                # total is what keeps it from being invisible money.
                progress.info(
                    f"{'Transcribed' if ok else 'Could not transcribe'} scanned page — "
                    f"{source_file} p.{page_num} · ~{tokens:,} cost-weighted tokens on OCR so far"
                )

            def _show_progress(step_id: str, turn: int, tokens: int, log: list) -> None:
                progress.info(
                    f"Running **{step_id}** — turn {turn}/{MAX_TOOL_ITERATIONS}, ~{tokens:,} cost-weighted "
                    f"tokens used so far, {len(log)} tool call(s) made. (Cap: {int(max_total_tokens):,})"
                )

            ledger = TokenLedger(model)
            all_results = {}
            for test_step_id, result in iter_control_results(
                spec,
                tmp_dir,
                client,
                model,
                sample_manifests=sample_manifests,
                max_total_tokens=int(max_total_tokens),
                on_turn=_show_progress,
                ocr_scanned_pages=ocr_scanned_pages,
                on_ocr=_show_ocr,
                max_run_tokens=int(max_run_tokens),
                ledger=ledger,
            ):
                progress.empty()
                _render_result(test_step_id, result)
                all_results[test_step_id] = result

            # Generate the CY workpaper (Excel) while the temp dir still
            # exists -- exhibits re-read the support files from it -- and
            # keep only the BYTES: the temp dir is gone the moment this
            # with-block closes, and the download button must survive
            # Streamlit's reruns.
            wp_bytes = wp_name = None
            wp_mirrored = False
            wp_warnings: list[str] = []
            try:
                outcome = build_cy_workpaper(
                    spec,
                    all_results,
                    py_testing_file.name,
                    tmp_dir,
                    support_dir=tmp_dir,
                    sample_manifests=sample_manifests,
                    population_workbook=pop_sample_file.name,
                    population_sheet=population_tab,
                )
                wp_bytes, wp_name = outcome.path.read_bytes(), outcome.path.name
                wp_mirrored = outcome.mirrored_py
                wp_warnings = list(outcome.warnings) + [
                    f"Tickmark {letter} on {name}: {anchor!r} could not be located on the page, "
                    f"so no box was drawn (the legend still lists it)."
                    for name, letter, anchor, _reason in outcome.unplaced_callouts
                ]
            except Exception as exc:  # noqa: BLE001 -- the run's results must still show even if the file build breaks
                st.warning(f"Couldn't generate the workpaper file: {exc}")

        st.session_state["run_output"] = {
            "results": all_results,
            "wp_bytes": wp_bytes,
            "wp_name": wp_name,
            "ledger": ledger,
            "wp_mirrored": wp_mirrored,
            "wp_warnings": wp_warnings,
        }
        st.success("Done.")
        _render_token_burn(ledger)

# Rendered OUTSIDE the run-click block so results and the download button
# survive Streamlit's script reruns (every widget click is a rerun -- without
# this, clicking Download would blank the whole results view).
_out = st.session_state.get("run_output")
if _out:
    if not run_clicked:  # freshly-run results were already rendered live above
        st.header("Results (last run)")
        for _tsid, _result in _out["results"].items():
            _render_result(_tsid, _result)
        if _out.get("ledger") is not None:
            _render_token_burn(_out["ledger"])
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
        if _out.get("wp_mirrored"):
            st.caption(
                "Built on the prior-year workpaper: same tabs and layout, with this "
                "year's sample, evidence and tickmarks written in, plus a Summary tab "
                "in front. Stamped DRAFT -- review and approve before it goes anywhere."
            )
        else:
            st.caption(
                "Built with the standalone layout -- the prior-year workpaper could "
                "not be used as a template, so this does NOT match last year's format. "
                "See the notes below. Stamped DRAFT."
            )
        # Everything the finished file cannot show: evidence that was not
        # attached, and tickmarks whose anchor was never found. A legend entry
        # pointing at a box that does not exist sends a reviewer hunting.
        if _out.get("wp_warnings"):
            with st.expander(f"⚠️ {len(_out['wp_warnings'])} note(s) about what is NOT in this workpaper"):
                for w in _out["wp_warnings"]:
                    st.markdown(f"- {w}")
