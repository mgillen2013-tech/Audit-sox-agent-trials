"""The joint: the agent's conclusions -> a CY workpaper that mirrors PY.

Both halves of this repo already existed and never met. agent/loop.py and
run_control.py decide what is true; agent/ooxml/ reproduces the prior-year
workpaper. This is the single call that runs one into the other, and the
only place that knows both exist.

TWO RENDERERS, ON PURPOSE
-------------------------
The PY-mirroring renderer needs the prior-year workpaper as a real .xlsx
with the tabs it writes into. That is the normal case and the good one:
the output inherits last year's layout instead of inventing one.

But a PY workpaper is not always an xlsx -- plenty are scanned or exported
to PDF -- and a template can be shaped differently enough that writing
into it would produce something wrong. In those cases this falls back to
agent/workpaper.py, the self-contained builder that owes nothing to a
template. The fallback output does NOT look like last year's; that is the
whole point of preferring the other one, and the caller is told which it
got rather than left to guess from the file.

The rule the fallback exists to serve: never refuse to produce a workpaper
because the layout could not be mirrored. A run that reached real
conclusions has already spent real money and real evidence review, and a
worse-looking file that carries them is worth far more than an error.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from agent.ooxml import ooxml_utils as ox
from agent.ooxml.build_workpaper import adopt_prior_output
from agent.ooxml.build_workpaper import build_workpaper as build_ooxml_workpaper
from agent.ooxml.models import PopulationRow, WorkpaperRequest
from agent.narrative_styles import NarrativeStyle
from agent.render_bridge import build_workpaper_request, summary_rows
from agent.schemas import ConclusionOutput, SamplePopulationManifest

# Column letters, matching the bridge's own extract-order convention.
_LETTERS = [chr(ord("A") + i) for i in range(26)]


@dataclass
class WorkpaperOutcome:
    """What was produced, and everything the caller must be told about it.

    `warnings` is not decoration. It carries the two things that are
    invisible in the finished file and matter most: evidence the bridge
    could not attach, and tickmarks whose anchor could not be located on
    the page. A legend entry pointing at a box that was never drawn sends
    a reviewer hunting for a mark that does not exist, so a caller that
    drops these is shipping a quietly wrong workpaper.
    """

    path: Path
    mirrored_py: bool
    warnings: list[str] = field(default_factory=list)
    unplaced_callouts: list[tuple[str, str, str, str]] = field(default_factory=list)
    # Why the prior-year layout was NOT used, in a sentence. Empty when it
    # was. Kept separate from `warnings` so a caller can put it front and
    # centre: a run that fell back produced a workpaper that looks nothing
    # like last year's, and the first version reported that only in small
    # grey text -- a reviewer reasonably took it for a regression.
    fallback_reason: str = ""
    # True when the fallback happened because NO test step concluded -- the
    # PY file was never the problem. The first loud notice blamed the PY
    # file for this case too and told the user to re-upload it, which sent
    # them after the wrong thing; the step's own error is the real cause.
    nothing_concluded: bool = False

    @property
    def summary_line(self) -> str:
        how = (
            "mirroring the prior-year workpaper's layout"
            if self.mirrored_py
            else f"using the standalone layout, NOT last year's format -- {self.fallback_reason}"
        )
        note = ""
        if self.unplaced_callouts:
            note = f"; {len(self.unplaced_callouts)} tickmark(s) could not be placed"
        return f"Wrote {self.path.name} {how}{note}"


def _population_rows(
    population_path: Path | None, sheet_name: str | None
) -> list[PopulationRow]:
    """The population tab, as rows keyed by template column letter.

    Returns [] when there is nothing to read, and the OOXML builder treats
    an empty list as "leave the template's own population alone" rather
    than clearing it -- deleting the population removes the basis for the
    sample, and a real output shipped that way once.
    """
    if not population_path or not sheet_name or not population_path.exists():
        return []
    import openpyxl

    wb = openpyxl.load_workbook(population_path, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return []
        ws = wb[sheet_name]
        rows: list[PopulationRow] = []
        for raw in ws.iter_rows(min_row=2, values_only=True):
            if not any(v is not None for v in raw):
                continue
            rows.append(
                PopulationRow(
                    values={
                        _LETTERS[i]: v
                        for i, v in enumerate(raw)
                        if i < len(_LETTERS) and v is not None
                    }
                )
            )
        return rows
    finally:
        wb.close()


def check_py_template(
    py_path: Path, sample_sheet: str = "Sample", population_sheet: str = "Population"
) -> tuple[bool, str, list[str]]:
    """(usable, why not, notes) for a prior-year file, before anything runs.

    Public so the app can call it the moment the file is uploaded. Finding
    out AFTER a paid run that the PY file was unusable is the expensive
    order; a red message beside the upload box costs nothing.

    Checked by actually unpacking and inspecting it rather than by
    extension. A file can be named .xlsx and still be the wrong shape.
    The probe applies adopt_prior_output exactly as the real build does,
    so a previous CY workpaper is judged as the builder will see it.
    """
    if py_path.suffix.lower() not in (".xlsx", ".xlsm"):
        return False, (
            f"the PY workpaper is a {py_path.suffix or 'file with no extension'}, "
            "not an .xlsx/.xlsm workbook"
        ), []
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        probe = str(Path(tmp) / "probe")
        try:
            ox.extract_template(str(py_path), probe)
            notes = adopt_prior_output(probe, sample_sheet)
            ox.check_template(
                probe,
                required_sheets=[sample_sheet, population_sheet],
                max_style_index=0,  # formats are copied from the template, not addressed by position
            )
        except ox.TemplateMismatch as exc:
            return False, _plain_mismatch(str(exc), sample_sheet, population_sheet), []
        except Exception as exc:  # noqa: BLE001 -- a corrupt or exotic file is a fallback, not a crash
            return False, f"the PY workpaper could not be opened as a template: {exc}", []
    return True, "", notes


def _plain_mismatch(detail: str, sample_sheet: str, population_sheet: str) -> str:
    """The template error, led by what to DO about it."""
    if "no sheet named" in detail:
        return (
            f"it needs tabs named {sample_sheet!r} and {population_sheet!r} "
            f"(or numbered '{sample_sheet} 1', '{sample_sheet} 2'... from a previous "
            f"CY workpaper). Details: {detail}"
        )
    return detail


# Kept under its old name for callers and tests that use the two-value form.
def _can_mirror(py_path: Path, sample_sheet: str, population_sheet: str) -> tuple[bool, str]:
    ok, why, _notes = check_py_template(py_path, sample_sheet, population_sheet)
    return ok, why


def build_cy_workpaper(
    spec: dict[str, Any],
    results: dict[str, dict],
    py_testing_filename: str,
    out_dir: Path,
    *,
    support_dir: Path | None = None,
    sample_manifests: dict[str, SamplePopulationManifest] | None = None,
    population_workbook: str | None = None,
    population_sheet: str | None = None,
    sample_sheet_name: str = "Sample",
    population_sheet_name: str = "Population",
    narrative_style: "NarrativeStyle" = "model",
    review_client: Any = None,
    review_model: str = "",
    evidence_items: "list[Any] | None" = None,
) -> WorkpaperOutcome:
    """Render one control's results, mirroring PY where that is possible.

    Two switches, both experiments, both off by default:

    narrative_style="template" gives every sample tab the SAME sentences
    with only its own values substituted, and moves the model's prose to
    the Summary. See agent/narrative_styles.py for why that might read
    better across five tabs, and what it costs.

    review_client + review_model turn on a second model pass that checks
    each tickmark line against the text of the evidence it cites, and
    writes the result into a "Reviewer check" column. It costs one API
    call per sampled item and never gates the workpaper -- see
    agent/reviewer.py.
    """
    support = Path(support_dir or out_dir)
    py_path = support / py_testing_filename
    out_dir.mkdir(parents=True, exist_ok=True)

    conclusions: list[tuple[str, ConclusionOutput]] = [
        (step_id, r["conclusion"]) for step_id, r in results.items() if "conclusion" in r
    ]
    if not conclusions:
        # Nothing concluded: fall straight through to the legacy builder,
        # which knows how to render a control whose steps all failed.
        outcome = _legacy(spec, results, py_testing_filename, out_dir, support, population_workbook,
                          ["no test step reached a conclusion, so there was nothing to mirror"])
        outcome.nothing_concluded = True
        return outcome

    can, why, adopted = check_py_template(py_path, sample_sheet_name, population_sheet_name)
    if not can:
        return _legacy(spec, results, py_testing_filename, out_dir, support, population_workbook,
                       [f"could not mirror the prior-year layout: {why}"], reason=why)

    # One Sample tab holds every sampled item across every test step, which
    # is how the prior-year workpapers are built -- the sample listing is
    # per control, not per step.
    warnings: list[str] = list(adopted)
    sample_items = []
    step_labels = []
    for step_id, conclusion in conclusions:
        manifest = (sample_manifests or {}).get(step_id)
        text = next(
            (s.get("test_step_text", "") for s in spec.get("test_steps", []) if s.get("test_step_id") == step_id),
            "",
        )
        request, step_warnings = build_workpaper_request(
            conclusion,
            manifest,
            control_id=spec["control_id"],
            template_path=py_path,
            output_path=out_dir / _mirrored_filename(spec["control_id"]),
            support_dir=support,
            narrative_style=narrative_style,
            step_label=f"{step_id}: {text}" if text else step_id,
        )
        sample_items.extend(request.sample_items)
        warnings.extend(f"[{step_id}] {w}" for w in step_warnings)
        step_labels.append(f"{step_id}: {text}" if text else step_id)

    output_path = out_dir / _mirrored_filename(spec["control_id"])
    request = WorkpaperRequest(
        template_path=str(py_path),
        output_path=str(output_path),
        sample_items=sample_items,
        population_rows=_population_rows(
            support / population_workbook if population_workbook else None, population_sheet
        ),
    )

    # Optional second opinion, per sampled item. Failures come back as
    # "not reviewed" rather than raising -- the workpaper it checks is
    # already built and valid, and losing it because an optional review
    # failed would be absurd.
    reviews: dict[str, str] = {}
    if review_client is not None and evidence_items:
        from agent.reviewer import review_sample

        for step_id, conclusion in conclusions:
            cited = {c.evidence_id for c in conclusion.evidence_citations}
            by_sample: dict[str, list[tuple[str, str, str]]] = {}
            for a in conclusion.attribute_results:
                if a.sample_id:
                    entries = by_sample.setdefault(a.sample_id, [])
                    entries.append((_LETTERS[len(entries)], a.attribute, a.value_observed))
            for sample_id, legend in by_sample.items():
                result = review_sample(
                    review_client,
                    review_model or "claude-opus-5",
                    sample_id=sample_id,
                    legend=legend,
                    evidence_items=evidence_items,
                    evidence_ids=cited,
                )
                reviews[sample_id] = result.summary
                if not result.all_supported and not result.error:
                    warnings.append(f"[{step_id}] reviewer: {result.summary}")

    unplaced: list[tuple[str, str, str, str]] = []
    build_ooxml_workpaper(
        request,
        sample_sheet=sample_sheet_name,
        population_sheet=population_sheet_name,
        on_unplaced=unplaced.extend,
        summary_rows=summary_rows(
            list(zip(step_labels, [c for _s, c in conclusions])), reviews=reviews
        ),
    )
    return WorkpaperOutcome(output_path, mirrored_py=True, warnings=warnings, unplaced_callouts=unplaced)


def _mirrored_filename(control_id: str) -> str:
    """DRAFT in the name, because the name is what travels: attached to an
    email, listed in a folder, nobody sees the banner inside. "_wp_" keeps
    it distinct from the standalone builder's "..._CY_Testing_DRAFT.xlsx",
    so which layout a file got is visible without opening it.
    """
    safe = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in str(control_id))
    return f"{safe}_CY_Testing_wp_DRAFT.xlsx"


def _legacy(
    spec: dict[str, Any],
    results: dict[str, dict],
    py_testing_filename: str,
    out_dir: Path,
    support: Path,
    population_workbook: str | None,
    warnings: list[str],
    *,
    reason: str = "",
) -> WorkpaperOutcome:
    """The standalone builder, for when PY cannot be mirrored."""
    from agent.workpaper import build_workpaper as build_legacy_workpaper

    path = build_legacy_workpaper(
        spec,
        results,
        py_testing_filename,
        out_dir,
        support_dir=support,
        source_workbook=population_workbook,
    )
    return WorkpaperOutcome(
        Path(path), mirrored_py=False, warnings=warnings,
        fallback_reason=reason or (warnings[0] if warnings else "reason not recorded"),
    )
