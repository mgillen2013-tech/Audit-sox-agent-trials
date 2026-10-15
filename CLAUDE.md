# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**Read `docs/STATUS.md` first.** It holds the agreed direction, the current
state, what is not yet proven, and the method for fixing bugs. This file
covers commands and architecture only.
# Project goalWe are building a SOX Testing Preparer Agent]: .It must work for ANY valid input. The examples we use are only for testing. # Who I amI'm not a programmer. Explain changes in plain language and keep code simple. # Core rules- NEVER hardcode names, values, file paths, or logic taken from an example.- Everything specific to one dataset belongs in a config file or input, not in the code.- Before writing code, state in one sentence what general behavior it adds.- If a change only makes the current example work, stop and ask me. # Examples- Current test example: [e.g. "Control 156"]. It is a test case, not the product.- Keep examples in a separate `examples/` or `tests/` folder.- Before saying something is done, run it on at least two different examples. # Before finishing any task1. Check: would this still work if the example changed completely?2. List any assumptions you made.3. Remind me of the project goal in one line.
## Commands

```
pip install -r requirements-dev.txt          # app deps + pytest
python -m pytest -q                          # full suite (~1 min, no network, no API key)
python -m pytest -q agent/tests/test_mirror.py
python -m pytest -q agent/tests/test_mirror.py -k tickmark   # single test / subset
python -m streamlit run cy_testing_app.py --server.address localhost   # the app
python -m agent.run_control path/to/control.json   # CLI runner (spec shape: agent/control.example.json)
```

- There is no build step and no linter config.
- Always `python -m streamlit` (plain `streamlit` is often not on PATH on
  Windows, which is where the user runs it), and always
  `--server.address localhost` (Streamlit otherwise binds to every interface).
- Model access is Microsoft Foundry: `ANTHROPIC_FOUNDRY_API_KEY`,
  `ANTHROPIC_FOUNDRY_RESOURCE`, optional `ANTHROPIC_MODEL` (a **deployment
  name**, not a model name), or the app sidebar. Tests never need these.
- The user downloads the branch as a ZIP and runs it on Windows; they do not
  use git. Instructions to them must work from a fresh unzip in PowerShell.

## Architecture

A deterministic pipeline with model calls only where judgment is needed:

```
agent/extraction/   files -> evidence items (pdf/excel text; OCR of scanned pages
                    is a vision model call in extraction/ocr.py)
agent/intake.py     sample listing (any columns) -> SamplePopulationManifest
agent/loop.py       the testing agent: manual tool-use loop, ends in submit_conclusion
agent/run_control.py  per-step orchestration (iter_control_results), used by app AND CLI
agent/build_cy_workpaper.py  the single joint from conclusions to an .xlsx
```

`build_cy_workpaper` tries three renderers in order and records which one
was used and why (`WorkpaperOutcome.mirrored_py`, `fallback_reason`):

1. **`agent/mirror/`: the preferred path.** It edits a copy of the
   prior-year (PY) workbook (design in `docs/mirror_engine.md`):
   - `structure.py`: the model reads how the PY is organised.
   - `inventory.py`: lists every element as it physically is.
   - `rewrite.py`: the model decides each element and compares support.
   - `apply.py`: carries out those decisions.
   - `verify.py`: writes the Mirror Check and Support Check tabs.

   It needs a client, so it runs only when one is passed.
2. **`agent/ooxml/` + `render_bridge.py`:** the older layout builder, which
   writes into the PY template by tab and heading.
3. **`agent/workpaper.py`:** the standalone layout, used when nothing can be
   mirrored.

All `.xlsx` writing is raw OOXML surgery on an unzipped package
(`agent/ooxml/ooxml_utils.py`). Never round-trip a workbook through openpyxl
to write it, because openpyxl silently drops drawings: text boxes, pictures
and tickmarks. openpyxl is fine for reading and for building test fixtures.
Every package goes through `ox.package_problems()` before it is zipped. A
problem it finds raises, and the build falls back rather than shipping a file
Excel cannot open.

The README's "only three modules call the model" predates the mirror engine.
`agent/mirror/` makes the structure, edit, second-look and check calls, all
through `rewrite.call_model`. That function never raises; a failed call keeps
the PY content and reports it. Token use is recorded in `agent/costs.py`
(`TokenLedger`), which only observes. The run caps in `loop.py` are in
cost-weighted tokens, not dollars.

## Invariants that are easy to break

- **The repo is public.** Workpapers, support and listings are live
  BrightView records. `.gitignore` blocks xlsx/xlsm/xls/pdf/csv, and test
  fixtures are built with openpyxl/reportlab/PIL, never copied. Real PYs used
  for trials live only in the session scratchpad.
- **No rules about what a workpaper contains.** Code must not decide what a
  tab, shape or cell means (no name patterns, word lists, or "nearest box"
  pairing). The model reads it from the PY and code verifies the reading.
  Unaddressed elements are kept and flagged, never deleted.
- **Fix the class of bug, not the instance** (see `docs/STATUS.md`). A fix
  that names a specific control, tab, word or layout is the wrong fix. Add a
  test built to the general shape and confirm it fails on the old code.
- **Test doubles must be at least as hostile as the real SDK.** For example,
  usage cache fields are `None`, not `0` (see `agent/tests/conftest.py`).
  App tests drive `cy_testing_app.py` through Streamlit `AppTest` with
  populated `session_state`. A render path no test executes has broken a paid
  run before.
- **A paid run must always produce a workpaper.** Results are recorded before
  display, failures fall back down the renderer chain, and an interrupted run
  can be rebuilt from `st.session_state["partial_run"]`.
- **OOXML details that have each caused an unopenable file:**
  - Relationship targets may be relative or absolute (`ox._part_path`).
  - Drawings may not declare the `xdr:`/`a:`/`r:` prefixes.
  - Copied sheets must take their r:id targets with them (`_detach_clone`).
  - Media names must be unique workbook-wide.
  - Inserting a tab must shift `localSheetId`/`activeTab`
    (`shift_tab_positions`).
  - Only one tab may be `tabSelected`.
 
  - 
