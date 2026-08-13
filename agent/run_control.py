"""Run every test step of one control against real files.

This is what section 1 of the design doc's intake describes, wired up:
a PY testing file, a CY sample list, and one or more CY support files go in;
one draft conclusion per test step comes out.

Still NOT built: an upload form (you hand it file paths in a JSON spec
instead), and per-test-step slicing of the PY testing file -- the PY
excerpts shown to the model are the entire extracted PY testing file for
every test step, not just the portion relevant to that step. That's a
known simplification, not an oversight: splitting a workpaper into
per-step sections is a real sub-problem on its own (see the design doc's
"confirm/edit step" note in section 1), and passing the whole thing is
harmless (just extra tokens) given the model already treats PY as
format/approach precedent, not evidence.

Usage:
    python3 -m agent.run_control path/to/control.json

control.json shape -- see control.example.json alongside this file. File
paths inside it are resolved relative to the JSON file's own directory, so
a whole control's files can live together in one folder and move as a unit.

{
  "control_id": "C-14",
  "control_objective_ref": "CO-4",
  "control_objective_text": "...",
  "py_testing_file": "PY_Testing_C14.pdf",
  "sample_list_file": "samples_C14.xlsx",
  "cy_support_files": ["Recon_Oct2026.xlsx", "GL_Export_Oct2026.pdf"],
  "test_steps": [
    {"test_step_id": "TS-4.2", "test_step_text": "..."}
  ]
}

test_steps[].py_conclusion_text is optional -- if omitted, the model reads
the PY conclusion from py_testing_file's extracted content itself instead
of a human retyping it.

Requires the same environment variables as agent/example_run.py
(ANTHROPIC_FOUNDRY_API_KEY, ANTHROPIC_FOUNDRY_RESOURCE, optionally
ANTHROPIC_MODEL).
"""

from __future__ import annotations

import json
import os
import sys
from collections.abc import Callable, Iterator
from pathlib import Path
from typing import Any

from agent.extraction import extract, extract_many
from agent.intake import parse_sample_list
from agent.loop import DEFAULT_MODEL, MAX_TOTAL_TOKENS, IncompleteRunError, TestStepRequest, run_test_step
from agent.schemas import ConclusionOutput, SamplePopulationManifest


def iter_control_results(
    spec: dict[str, Any],
    base_dir: Path,
    client: Any,
    model: str,
    sample_manifests: dict[str, SamplePopulationManifest] | None = None,
    max_total_tokens: int = MAX_TOTAL_TOKENS,
    on_turn: "Callable[[str, int, int, list], None] | None" = None,
) -> Iterator[tuple[str, dict]]:
    """The testable core -- no file writing, no env var reads. Yields
    (test_step_id, {"conclusion": ConclusionOutput, "audit_log": [...]})
    on success, or (test_step_id, {"error": str, ...}) on failure, as each
    test step finishes, rather than blocking silently until the whole
    control is done -- a control with several steps can take minutes, and a
    caller (the Streamlit app) that wants to show progress per step needs
    results as they land, not a batch dump at the end.

    A failure that aborted via run_test_step's token-budget or
    max-iterations circuit breaker (agent.loop.IncompleteRunError) still
    yields everything already paid for: the error dict also carries
    "audit_log" (every tool call made before the abort), "reason"
    ("token_budget_exceeded" or "max_iterations"), "tokens_used", and
    "turns_used". A run that already spent real money on API calls before
    failing should never come back as a bare error string with nothing to
    show for it -- see the design doc's note on this circuit breaker.

    sample_manifests: pass already-built manifests to skip
    spec["sample_list_file"] / parse_sample_list() entirely -- used by the
    Streamlit app, which builds one manifest per test step directly from
    whatever columns that step's uploaded file actually has (see
    agent.intake.build_manifest_from_any_columns) rather than requiring one
    combined file in the clean fixed-column format. Falls back to the
    file-based path when omitted, for the CLI's control.json workflow.

    on_turn: forwarded to run_test_step for each test step, called with
    (test_step_id, turn_number, cumulative_tokens_used, audit_log_so_far) so
    a caller can show live progress -- a step can run for a couple of
    minutes with several tool calls, and iter_control_results only yields
    once a step is fully done or aborted.
    """
    py_evidence = extract(base_dir / spec["py_testing_file"])
    cy_evidence = extract_many([base_dir / f for f in spec["cy_support_files"]])
    if sample_manifests is None:
        sample_manifests = parse_sample_list(base_dir / spec["sample_list_file"])

    for step in spec["test_steps"]:
        test_step_id = step["test_step_id"]
        manifest = sample_manifests.get(test_step_id)
        if manifest is None:
            print(
                f"  WARNING: no sample list rows found for {test_step_id!r} -- "
                f"running with an empty sample (check_sample_coverage will error)"
            )

        request = TestStepRequest(
            test_step_id=test_step_id,
            control_id=spec["control_id"],
            control_objective_ref=spec["control_objective_ref"],
            control_objective_text=spec["control_objective_text"],
            test_step_text=step["test_step_text"],
            py_conclusion_text=step.get("py_conclusion_text", ""),
            py_support_excerpts=py_evidence,
        )

        step_on_turn = (
            (lambda turn, tokens, log, _id=test_step_id: on_turn(_id, turn, tokens, log))
            if on_turn is not None
            else None
        )
        try:
            conclusion, audit_log = run_test_step(
                request,
                cy_evidence,
                manifest,
                client,
                model=model,
                max_total_tokens=max_total_tokens,
                on_turn=step_on_turn,
            )
        except IncompleteRunError as exc:
            yield test_step_id, {
                "error": str(exc),
                "audit_log": exc.audit_log,
                "reason": exc.reason,
                "tokens_used": exc.tokens_used,
                "turns_used": exc.turns_used,
            }
            continue
        except Exception as exc:  # noqa: BLE001 -- one step's failure shouldn't kill the run
            yield test_step_id, {"error": str(exc)}
            continue

        yield test_step_id, {"conclusion": conclusion, "audit_log": audit_log}


def run_control(
    spec: dict[str, Any],
    base_dir: Path,
    client: Any,
    model: str,
    sample_manifests: dict[str, SamplePopulationManifest] | None = None,
    max_total_tokens: int = MAX_TOTAL_TOKENS,
) -> dict[str, dict]:
    """Batch wrapper over iter_control_results() for callers (the CLI, tests)
    that just want the whole set of results at the end.
    """
    return dict(
        iter_control_results(
            spec, base_dir, client, model, sample_manifests=sample_manifests, max_total_tokens=max_total_tokens
        )
    )


def main() -> None:
    if len(sys.argv) != 2:
        print("Usage: python3 -m agent.run_control path/to/control.json")
        sys.exit(1)

    spec_path = Path(sys.argv[1])
    base_dir = spec_path.parent
    spec = json.loads(spec_path.read_text())

    from anthropic import AnthropicFoundry  # imported here so tests don't need Foundry creds to import this module

    client = AnthropicFoundry()
    model = os.environ.get("ANTHROPIC_MODEL", DEFAULT_MODEL)
    print(f"Using model: {model}\n")

    results = run_control(spec, base_dir, client, model)

    output_dir = base_dir / "output" / spec["control_id"]
    output_dir.mkdir(parents=True, exist_ok=True)

    for test_step_id, result in results.items():
        out_path = output_dir / f"{test_step_id}.json"
        if "error" in result:
            print(f"[{test_step_id}] FAILED: {result['error']}")
            error_payload: dict[str, Any] = {"error": result["error"]}
            if "audit_log" in result:
                error_payload["reason"] = result["reason"]
                error_payload["tokens_used"] = result["tokens_used"]
                error_payload["turns_used"] = result["turns_used"]
                error_payload["audit_log"] = [entry.__dict__ for entry in result["audit_log"]]
                print(
                    f"  ({result['reason']}, {result['tokens_used']:,} tokens used, "
                    f"{len(result['audit_log'])} tool call(s) preserved)"
                )
            out_path.write_text(json.dumps(error_payload, indent=2))
            continue

        conclusion: ConclusionOutput = result["conclusion"]
        print(
            f"[{test_step_id}] {conclusion.conclusion} "
            f"(confidence: {conclusion.confidence}, tool calls: {conclusion.model_metadata.tool_call_count})"
        )
        out_path.write_text(
            json.dumps(
                {
                    "conclusion": conclusion.model_dump(),
                    "audit_log": [entry.__dict__ for entry in result["audit_log"]],
                },
                indent=2,
            )
        )

    print(f"\nWrote results to {output_dir}")


if __name__ == "__main__":
    main()
