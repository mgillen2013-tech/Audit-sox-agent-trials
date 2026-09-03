# PY Testing Workpaper Builder — Methodology

This documents how the BrightView PTP.AP.156 workpaper was reconstructed, refactored
into a repeatable pipeline. It's split so a "critical thinking" LLM agent only has to
make judgment calls — everything mechanical is deterministic Python.

## Why not just use openpyxl end-to-end?

openpyxl **cannot round-trip shapes**. This workpaper's evidentiary value lives in
floating shapes drawn on top of pasted images: red-outlined callout boxes with a
letter (A, B, C…) that visually tie a value in the raw data extract to the same value
on the source document. If you load this file with openpyxl and save it again,
every one of those shapes silently disappears — openpyxl's drawing support only
understands images and charts, not arbitrary `<xdr:sp>` textboxes/rectangles.

So the pipeline works **below** openpyxl, directly on the OOXML parts inside the
`.xlsx` zip:

- `xl/worksheets/sheetN.xml` — cell values (raw data row, IA Calculation box)
- `xl/sharedStrings.xml` — every text cell value is an index into this table
- `xl/drawings/drawingN.xml` — images + textboxes + rectangles, positioned in EMU
  (914,400 EMU = 1 inch)
- `xl/media/imageN.png` — the actual evidence images
- `[Content_Types].xml`, `*.rels` — plumbing that has to stay consistent

openpyxl is still used, but only as a **read-only inspector** (to understand a
template) and never to write the final file.

## The four-tab template pattern

Every one of these SOX PTP.AP control workpapers follows the same shape:

1. **IA Leadsheet** — a single static textbox: Control ID, Control Description,
   Test Steps, Test Reference. This does not change per sample; it's control
   metadata, not evidence.
2. **Sample** — one raw-data row (pulled from the JDE/Data Access Studio extract)
   plus:
   - a narrative textbox ("Test Step 1: IA verified...")
   - an **IA Calculation** box (plain cells with borders, not a shape) if the
     sample needs a supporting tie-out sum (e.g., multiple invoice lines summing
     to the payment)
   - the evidence images (invoice, approval, payment confirmation, etc.)
   - **callout rectangles** with a single bold red letter, positioned on top of
     the raw data row AND on top of each evidence image, at the exact pixel
     location of the value being tied out
3. **Population** — the full query extract the sample was drawn from (no shapes,
   just data — safe to touch with openpyxl if you ever needed to).
4. **Parameters** — a screenshot of the Data Access Studio query used to pull the
   population, a narrative textbox describing completeness/accuracy procedures,
   and 1-2 callout rectangles highlighting the record count and date-range filter.

Because the *shape* of every workpaper is identical, only the **content** changes
per instance. That's the seam the pipeline is built around.

## The callout-positioning problem (and how it's solved deterministically)

The hard part isn't drawing a red box — it's knowing *where* to draw it. A human
auditor eyeballs the screen. The pipeline uses OCR instead:

1. Render each evidence PDF page to a PNG (`pdf2image`, 150 DPI).
2. Autocrop whitespace so images aren't full 8.5x11 blank margins.
3. Run `pytesseract.image_to_data` to get a bounding box for every word.
4. For each tie-out attribute, search for its anchor text (e.g. `"21022452"`,
   `"30,000.00"`, `"11/10/2025"`) and take the union of the matching word boxes,
   padded a few px.
5. Convert that pixel box to EMU using the image's actual DPI
   (`EMU = px * 914400 / dpi`), then apply the same *display scale factor* used to
   place the image on the sheet (images are placed at a chosen EMU width, not
   their native size, so any overlay math must scale by
   `chosen_width_emu / native_width_px`, not by the raw DPI conversion alone).
6. Draw the rectangle using `xdr:absoluteAnchor` — absolute x/y/width/height in
   EMU from the sheet's top-left corner. This sidesteps Excel's column/row grid
   entirely, which is what makes pixel-accurate overlays tractable. (A
   `oneCellAnchor`/`twoCellAnchor`, by contrast, requires you to know the drawing
   sheet's exact column widths and row heights to compute an offset, which is far
   more fragile.)

The raw-data row gets the same treatment, but using the *cell grid* instead of an
image: column pixel widths are computed from the sheet's `<col width=.../>`
definitions using Excel's documented approximation
(`px ≈ round(width_chars * 7 + 5)` for the default Calibri 11 font), then converted
with the *fixed* 96 DPI cell-grid constant (`9525 EMU/px`) — this is different from
the image DPI constant above and is a common source of misalignment if conflated.

## What's deterministic vs. what needs judgment

This split is the actual point of this package — it's built so a Claude-powered
agent only has to fill in a structured `SampleItem` object; it never touches XML.

| Deterministic (code, `ooxml_utils.py` / `drawing_builder.py`) | Judgment (LLM agent, produces `models.py` objects) |
|---|---|
| Unzipping/rezipping the xlsx | Which raw sample was selected and its field values |
| Shared-string dedup and index assignment | Which attributes need a tie-out letter (A, B, C…) and what each one means (Payee Name/Number, Amount, Date, Approver...) |
| EMU/pixel conversion math | The OCR anchor text to search for, per attribute, per evidence image |
| OCR word-box lookup | Whether the control was satisfied / whether there's an exception, and the narrative sentence saying so |
| Drawing XML generation (images, rectangles, textboxes) | Whether an IA Calculation tie-out is needed, and what its line items are |
| Cell XML generation (raw data row, IA Calc rows) | Population-tab completeness/accuracy commentary (doc types excluded, period tested) |
| Re-zipping and content-type consistency | Overall pass/fail conclusion |

In an Anthropic-API pipeline this maps cleanly onto your existing multi-model
design (per `/areas/sox-testing-agent.md`):

- **Opus** (planning): given the raw sample extract + evidence PDFs, decide the
  test attributes, draft the narrative, decide pass/fail.
- **Sonnet** (production reasoning): fill in the structured `SampleItem` /
  `TieOutField` objects (including picking OCR anchor text) and validate against
  the pydantic schema — this is exactly the "deterministic pre-LLM extraction
  with pydantic-validated structured output" pattern already planned.
- **Haiku** (high volume): run this per-sample across every PK/PT item in a
  period; each sample is independent and cheap to parallelize.
- The **Python code in this package** is the tool the agent calls
  (`check_sample_coverage`-style) — never asked to "write XML," only to
  populate structured objects that `build_workpaper()` turns into the final file.

## Files in this package

- `models.py` — pydantic schemas for everything the LLM agent needs to produce.
- `ooxml_utils.py` — template loading, shared-string management, sheet cell XML,
  EMU/column-width math.
- `ocr_utils.py` — PDF→PNG rendering, autocrop, OCR anchor-text bounding boxes.
- `drawing_builder.py` — builds `drawingN.xml` (images, callouts, narrative
  textboxes) from structured input.
- `build_workpaper.py` — top-level `build_workpaper()` orchestration function,
  plus the exact call used to regenerate the Premiere Onboard sample as a
  worked example.

## Known simplifications vs. a from-scratch original

- IA Calculation box is built as plain bordered/filled cells (matching what the
  *original* template actually did — it is not a floating shape).
- Callout rectangles carry their letter as text *inside* the rectangle shape,
  rather than as a separate adjacent textbox (the original occasionally used
  two shapes per callout). Visually equivalent, half the shape count.
- Column-width→pixel conversion uses Excel's standard approximation, not the
  exact MDW (max digit width) calculation Excel itself uses internally. This is
  accurate to a few pixels — fine for a highlight box, not for anything load-bearing.
