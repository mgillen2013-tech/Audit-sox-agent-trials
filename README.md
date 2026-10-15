# CY SOX control testing agent

Tests a SOX control against this year's evidence and produces a workpaper
in the prior year's format.

Give it the prior-year workpaper, the population and sample workbook, and
the current-year support (invoices, approvals, payment screens). It reads
the evidence — including scanned pages, via OCR — finds what supports each
sampled item, concludes on each test step, and writes an Excel workpaper
with the evidence pasted in, the tested values boxed in red, and tickmarks
tied to a legend.

## Running it

On Windows: double-click **`Start CY Testing Agent.bat`** in the unzipped
folder. The first time, it installs what the app needs (a few minutes);
after that it just starts the app in your browser. It runs on your
computer only. Close the black window to stop it.

Or by hand:

```
pip install -r requirements.txt
python -m streamlit run cy_testing_app.py --server.address localhost
```

`python -m streamlit` rather than plain `streamlit`: on Windows, pip
installs the `streamlit` launcher into a Scripts directory that is often
not on PATH, so `streamlit run ...` fails with "not recognized as a name
of a cmdlet" even though it is correctly installed. The `python -m` form
does not depend on PATH and works everywhere.

`--server.address localhost` keeps it on your machine. Streamlit
otherwise binds to every network interface, which on corporate wifi means
others could reach it while it runs.

Needs a Claude deployment on Microsoft Foundry. Set `ANTHROPIC_FOUNDRY_API_KEY`
and `ANTHROPIC_FOUNDRY_RESOURCE`, or type them into the sidebar.

## How it is put together

A deterministic Python pipeline with an agent in the middle:

```
extract the files      ← ordinary code
run the agent          ← the model decides what to look for
build the workpaper    ← ordinary code
```

Only three modules call the model at all: `agent/loop.py` (the judgment),
`agent/extraction/ocr.py` (reading scanned pages) and `agent/reviewer.py`
(an optional second opinion). Everything else is deterministic, which is
what keeps the cost down and the output reviewable.

The agent has four tools — search the evidence, check whether a value is
printed on a page, flag an exception, request missing support — and ends
by submitting a conclusion. Sample coverage and the fabrication guard are
enforced in code, not asked for in the prompt.

## Documentation

- `docs/cy_testing_agent_design.md` — the design, and why each guard exists
- `docs/layout_must_be_mirrored_not_computed.md` — why layout is read from
  the prior-year file rather than computed
- `docs/ooxml_workpaper_methodology.md` — how the workpaper is written
- `docs/enhancement_notes.md` — agreed but unbuilt

## Notes

IPE is not tested here; a human performs it. The workpaper shows the row
as outstanding rather than omitting it.

Evidence files are gitignored. This repository is public — do not commit
client data.
