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

One upload: the sample listing. Its sheets accept ANY columns (see
agent.intake.build_manifest_from_any_columns), because a real export
straight out of E1 has columns like "invoice number f0411.vinv", not
"identifying_details". The whole workbook is added to CY support evidence,
so its tabs are searchable rather than only parsed.

The population is OPTIONAL and behind a toggle. It used to be the headline
of this section because IPE completeness/accuracy was tested here; a human
performs that now, so what is left is worth having and not worth asking
for every run: the population SIZE reaches the model, which is how it
tells a deliberate 1-item sample from a partial upload (a real run treated
a correct sample as suspicious without it), and the rows are carried into
the workpaper's Population tab.

Selection method and population description are not collected at all --
they were never surfaced to the model or the workpaper, and forcing an
answer just produced noise.

This is still a shell, not the full review UI from the design doc (no
approve/edit/reject, no citation-card source previews) -- it's the intake
+ run + raw-results half, so you can drive a real control without hand-
building a control.json.
"""

from __future__ import annotations

import contextlib
import json
import os
import shutil
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
from agent.build_cy_workpaper import build_cy_workpaper, check_py_template

# Plain-English next steps for the ways a test step commonly fails. Keyed by
# the runner's "reason" code where it has one, and matched on the error
# text otherwise (API errors carry no code of ours).
_FAILURE_HINTS = {
    "token_budget_exceeded": (
        "it hit the per-step token cap before finishing. Raise "
        "\"Max cost-weighted tokens per test step\" in the sidebar and re-run."
    ),
    "run_budget_exceeded": (
        "the whole control hit its token cap. Raise \"Max cost-weighted tokens "
        "for the whole control\" in the sidebar and re-run."
    ),
    "model_refusal": (
        "the model declined to answer. Re-run; if it repeats, send the error text."
    ),
    "max_iterations": (
        "it used every allowed turn without submitting a conclusion. Re-run; if "
        "it repeats, the support may not contain what the step asks for."
    ),
}


def _hint_for(error: str) -> str:
    text = error.lower()
    if "deploymentnotfound" in text or "404" in text:
        return (
            "Foundry has no deployment with the name in the sidebar's Model box. "
            "Copy the Name column from your resource's Deployments page exactly -- "
            "it can differ from the model's own name -- and check Resource name "
            "is the resource that deployment lives on."
        )
    if any(k in text for k in ("connection error", "getaddrinfo", "name or service", "nodename", "failed to resolve")):
        return (
            "couldn't reach the resource. Check Resource name: it's the part "
            "before .services.ai.azure.com in your endpoint URL."
        )
    if "401" in text or "authentication" in text or "api key" in text:
        return "the API key was rejected. Re-enter it in the sidebar."
    if "429" in text or "rate limit" in text:
        return "Foundry rate-limited the run. Wait a minute and re-run."
    if "overloaded" in text or "529" in text or "timeout" in text or "timed out" in text:
        return "the model service was busy or slow. Re-run."
    return ""


# The connection settings this computer last used (resource and model), so
# they do not have to be retyped every session. Control details are NOT
# remembered: each run is usually a different control, and a box that
# arrives pre-filled with the last one has to be cleared before every use. Stored in the user's home folder, NOT the app folder: the
# code arrives as a fresh ZIP each time and anything beside it would be lost.
# The API key is never written here.
_REMEMBER_PATH = Path.home() / ".cy_testing_agent" / "last_inputs.json"


def _remembered() -> dict:
    try:
        return json.loads(_REMEMBER_PATH.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001 -- missing or unreadable just means no memory
        return {}


def _remember(values: dict) -> None:
    try:
        _REMEMBER_PATH.parent.mkdir(parents=True, exist_ok=True)
        _REMEMBER_PATH.write_text(json.dumps(values, indent=2), encoding="utf-8")
    except Exception:  # noqa: BLE001 -- convenience only; never break a run over it
        pass


_last = _remembered()


def _prefill(key: str, fallback: str = ""):
    return _last.get(key) or fallback


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
        value=os.environ.get("ANTHROPIC_FOUNDRY_RESOURCE") or _prefill("resource"),
        help='The subdomain in your endpoint URL, e.g. "example-resource".',
    )
    model = st.text_input(
        "Model",
        value=os.environ.get("ANTHROPIC_MODEL") or _prefill("model", DEFAULT_MODEL),
        help=(
            "The deployment NAME from your resource's Deployments page in "
            "Foundry -- any Claude model works (Sonnet, Opus, ...), but the "
            "name must match exactly."
        ),
    )
    st.caption("Resource and model are remembered on this computer. The API key is never saved.")

    # One tiny request (1 output token, a fraction of a cent) that proves
    # key + resource + deployment name before a run is spent finding out.
    if st.button("Test connection"):
        if not api_key or not resource or not model:
            st.error("Fill in the API key, resource name and model first.")
        else:
            try:
                AnthropicFoundry(api_key=api_key, resource=resource).messages.create(
                    model=model, max_tokens=1, messages=[{"role": "user", "content": "ping"}]
                )
                st.success(f"✅ Connected: '{model}' answered on {resource}.")
            except Exception as exc:  # noqa: BLE001 -- shown to the user, not raised
                _hint = _hint_for(str(exc))
                st.error(f"Couldn't connect: {exc}" + (f"\n\n**What to do:** {_hint}" if _hint else ""))

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
    control_id = st.text_input("Control ID", placeholder="e.g. SS.PTP.AP.156")
with col2:
    control_objective_ref = st.text_input(
        "Control objective reference", placeholder="e.g. Non-PO Invoice Review and Approval"
    )
control_objective_text = st.text_area(
    "Control objective",
    placeholder="What this control is supposed to accomplish.",
)

# ── 2. Sample selections ────────────────────────────────────────────────────
st.header("2. Sample selections")
st.caption(
    "The Excel file listing the items selected for testing. One tab per test "
    "step, with whatever columns your extract happens to have."
)
pop_sample_file = st.file_uploader(
    "Sample listing", type=["xlsx", "xls", "xlsm"], key="pop_sample_file"
)

sheet_names: list[str] = []
if pop_sample_file is not None:
    try:
        _wb = openpyxl.load_workbook(BytesIO(pop_sample_file.getvalue()), read_only=True)
        sheet_names = _wb.sheetnames
        _wb.close()
    except Exception as exc:  # noqa: BLE001 -- surfaced inline, not a crash
        st.error(f"Couldn't read tabs from {pop_sample_file.name}: {exc}")

# Optional, and behind a toggle so the common case is one upload and one
# tab pick. IPE is performed by a human now, which is what the population
# tab mainly existed for; what remains is worth having but is not worth
# asking for every run:
#   - the population SIZE reaches the model, so it can tell a deliberate
#     1-item sample from a partial upload (a real run treated a correct
#     sample as suspicious without it)
#   - the rows are carried into the workpaper's Population tab
population_tab = None
if sheet_names and st.checkbox(
    "This file also contains the full population",
    key="has_population",
    help="Optional. Adds the population to the workpaper and tells the agent "
         "how large the population was, so it can tell a small-but-complete "
         "sample from a partial one.",
):
    population_tab = st.selectbox("Which tab is the full population?", sheet_names, key="population_tab")

# ── 3. Test steps ────────────────────────────────────────────────────────────
st.header("3. Test steps")
num_steps = st.number_input("How many test steps does this control have?", min_value=1, max_value=10, value=1, step=1)

test_steps = []
for i in range(int(num_steps)):
    with st.expander(f"Test step {i + 1}", expanded=(i == 0)):
        # The id is derived from position, not typed. It was a required
        # field whose value the user never chose anything meaningful for --
        # everything downstream only needs it to be unique and stable
        # within the run, and "TS-1" is what everyone typed anyway.
        tsid = f"TS-{i + 1}"
        tstext = st.text_area(
            "Test step text",
            key=f"tstext_{i}",
            placeholder="What this step requires you to test.",
        )

        sample_tab = None
        if sheet_names:
            sample_tab = st.selectbox(
                "Which tab holds this step's sample selections?",
                sheet_names,
                key=f"sampletab_{i}",
            )
        else:
            st.info("Upload the sample listing above to pick this step's tab.")

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


def _py_template_verdict(name: str, data: bytes) -> tuple[bool, str, list[str]]:
    """check_py_template on the uploaded bytes, once per distinct file.

    Cached in session state: Streamlit reruns this whole script on every
    click, and unzipping the workbook each time would make the form lag.
    """
    import hashlib

    key = ("py_verdict", name, hashlib.sha256(data).hexdigest())
    cached = st.session_state.get("_py_verdict")
    if cached and cached[0] == key:
        return cached[1]
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / name
        path.write_bytes(data)
        verdict = check_py_template(path)
    st.session_state["_py_verdict"] = (key, verdict)
    return verdict


# Said BEFORE the run, beside the upload, because finding out afterwards
# costs a paid run. A real run used a file this check would have rejected,
# got a workpaper in a completely different layout, and the only notice was
# small grey text under the download button.
if py_testing_file is not None:
    _ok, _why, _notes = _py_template_verdict(py_testing_file.name, py_testing_file.getvalue())
    if _ok:
        st.success("✅ The workpaper will be built on this file's layout (tabs, tickmarks, format).")
        for _n in _notes:
            st.caption(f"ℹ️ {_n[0].upper()}{_n[1:]}.")
    else:
        st.error(
            "**This file can't be used as the layout for the workpaper.** The run will "
            "still work, but the workpaper will come out in the plain standalone format "
            "-- no prior-year tabs, no tickmark boxes.\n\n"
            f"Why: {_why}\n\n"
            "Fix: upload the original prior-year workpaper (.xlsx) instead."
        )
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
    st.caption(f"Tool calls: {conclusion.model_metadata.tool_call_count}")

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


def _build_workpaper(spec, results, py_name, run_dir, manifests, pop_name, pop_tab, ledger) -> dict:
    """Build the workpaper from whatever results exist; return what the page shows.

    Used at the end of a run AND to recover an interrupted one, so a run
    that was cut off still produces a file from the steps that finished.
    Keeps only the BYTES: the download button must survive Streamlit's
    reruns, and the staged files are deleted afterwards.
    """
    out = {
        "results": results, "wp_bytes": None, "wp_name": None, "ledger": ledger,
        "wp_mirrored": False, "wp_fallback_reason": "", "wp_nothing_concluded": False,
        "wp_warnings": [],
    }
    try:
        outcome = build_cy_workpaper(
            spec,
            results,
            py_name,
            Path(run_dir),
            support_dir=Path(run_dir),
            sample_manifests=manifests,
            population_workbook=pop_name,
            population_sheet=pop_tab,
        )
        out.update(
            wp_bytes=outcome.path.read_bytes(),
            wp_name=outcome.path.name,
            wp_mirrored=outcome.mirrored_py,
            wp_fallback_reason=outcome.fallback_reason,
            wp_nothing_concluded=outcome.nothing_concluded,
            wp_warnings=list(outcome.warnings) + [
                f"Tickmark {letter} on {name}: {anchor!r} could not be located on the page, "
                f"so no box was drawn (the legend still lists it)."
                for name, letter, anchor, _reason in outcome.unplaced_callouts
            ],
        )
    except Exception as exc:  # noqa: BLE001 -- the run's results must still show even if the file build breaks
        st.warning(f"Couldn't generate the workpaper file: {exc}")
    return out


# ── Run ───────────────────────────────────────────────────────────────────
st.header("5. Run")
run_clicked = st.button("Run test steps", type="primary")
st.caption(
    "While it runs, don't click or change anything else on this page -- Streamlit "
    "restarts the page on any click, which stops the run. If that happens, the "
    "steps that already finished are kept and you can still build the workpaper."
)

# A run that was cut off part-way. Streamlit restarts the whole script on
# any click, which silently abandons a run in progress: no message, no
# workpaper, and the finished steps -- already paid for -- vanished with it.
# Each finished step is now saved as it completes, so it can be recovered.
_partial = st.session_state.get("partial_run")
if _partial and not run_clicked:
    _done = [t for t in _partial["step_ids"] if t in _partial["results"]]
    _missing = [t for t in _partial["step_ids"] if t not in _partial["results"]]
    st.warning(
        "**The last run stopped before it finished** -- usually because something on "
        "the page was clicked while it ran.\n\n"
        f"Finished: {', '.join(_done) or 'none'}  ·  Not run: {', '.join(_missing) or 'none'}"
    )
    if _done and st.button("Build the workpaper from the finished steps"):
        st.session_state["run_output"] = _build_workpaper(
            _partial["spec"], _partial["results"], _partial["py_name"], _partial["run_dir"],
            _partial["manifests"], _partial["pop_name"], _partial["pop_tab"], _partial["ledger"],
        )
        st.session_state.pop("partial_run", None)
        st.session_state.pop("run_dir", None)
        shutil.rmtree(_partial["run_dir"], ignore_errors=True)
        st.rerun()
if run_clicked:
    _remember({"resource": resource, "model": model})
    errors = []
    if not control_id or not control_objective_ref or not control_objective_text:
        errors.append("Fill in all of the control details.")
    if pop_sample_file is None:
        errors.append("Upload the sample listing.")
    for s in test_steps:
        if not s["test_step_text"]:
            errors.append(f"Fill in the test step text for test step {s['test_step_id']}.")
        if pop_sample_file is not None and not s["sample_tab"]:
            errors.append(f"Pick a sample tab for test step {s['test_step_id']}.")
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
        # A staging folder that survives the page restarting mid-run, so an
        # interrupted run can still be built into a workpaper. The previous
        # run's folder is removed first; a finished run removes its own.
        shutil.rmtree(st.session_state.pop("run_dir", ""), ignore_errors=True)
        st.session_state.pop("partial_run", None)
        with contextlib.nullcontext(tempfile.mkdtemp(prefix="cy_run_")) as tmp:
            st.session_state["run_dir"] = tmp
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
            # Saved BEFORE the first step, and updated as each finishes (the
            # dict is shared, so filling all_results fills this too).
            st.session_state["partial_run"] = {
                "spec": spec, "results": all_results, "manifests": sample_manifests,
                "py_name": py_testing_file.name, "pop_name": pop_sample_file.name,
                "pop_tab": population_tab, "run_dir": tmp, "ledger": ledger,
                "step_ids": [s["test_step_id"] for s in test_steps],
            }
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
                # Record the result BEFORE trying to display it, and never
                # let a display failure escape. This loop sits between a
                # finished agent run and the workpaper build, so an
                # exception here throws away work that has already been
                # paid for -- which is exactly what happened when a field
                # removed from the schema was still referenced below.
                all_results[test_step_id] = result
                try:
                    _render_result(test_step_id, result)
                except Exception as exc:  # noqa: BLE001
                    st.warning(
                        f"Couldn't display {test_step_id} ({exc}). The result is still "
                        f"included in the workpaper below."
                    )

            st.session_state["run_output"] = _build_workpaper(
                spec, all_results, py_testing_file.name, tmp_dir, sample_manifests,
                pop_sample_file.name, population_tab, ledger,
            )
            # Finished normally: nothing to recover, and the staged copies of
            # the evidence are no longer needed on disk.
            st.session_state.pop("partial_run", None)
            shutil.rmtree(st.session_state.pop("run_dir", ""), ignore_errors=True)

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
        elif _out.get("wp_nothing_concluded"):
            # Not a PY problem at all: every test step failed, so there was
            # nothing to lay out in any format. Show the steps' own errors
            # HERE, with what to do -- the first version of this notice
            # blamed the PY file and sent the user to re-upload it.
            _failures = "\n".join(
                f"- **{_tsid}:** {_r.get('error', 'no conclusion returned')}"
                + (f"\n  - What to do: {_FAILURE_HINTS[_r['reason']]}" if _r.get("reason") in _FAILURE_HINTS else "")
                + (f"\n  - What to do: {_hint_for(_r.get('error', ''))}" if _r.get("reason") not in _FAILURE_HINTS and _hint_for(_r.get("error", "")) else "")
                for _tsid, _r in _out["results"].items()
                if "conclusion" not in _r
            )
            st.error(
                "**The test step didn't finish, so there's no result to put in a workpaper.** "
                "Your prior-year file wasn't the problem.\n\n"
                f"{_failures}\n\n"
                "The downloadable file below only records the failure."
            )
        else:
            # Red, on screen, with the reason -- not a caption. This file
            # looks nothing like last year's, and a reviewer who does not
            # know why will take it for a broken build.
            st.error(
                "**This workpaper is NOT in last year's format.** The prior-year file "
                "couldn't be used as the layout, so it was built in the plain standalone "
                "format instead (no prior-year tabs, no tickmark boxes on the exhibits).\n\n"
                f"Why: {_out.get('wp_fallback_reason') or 'see the notes below'}\n\n"
                "Fix: re-run with the original prior-year workpaper (.xlsx)."
            )
        # Everything the finished file cannot show: evidence that was not
        # attached, and tickmarks whose anchor was never found. A legend entry
        # pointing at a box that does not exist sends a reviewer hunting.
        if _out.get("wp_warnings"):
            with st.expander(
                f"⚠️ {len(_out['wp_warnings'])} note(s) about what is NOT in this workpaper",
                expanded=not _out.get("wp_mirrored"),
            ):
                for w in _out["wp_warnings"]:
                    st.markdown(f"- {w}")
