# The mirror engine: the PY is the document, the model only edits it

`agent/mirror/`. Preferred path for every CY workpaper; the earlier
builders remain as fallbacks only.

## Why

Every earlier renderer GENERATED a workpaper and then tried to make it look
like the PY. Each new control exposed another way it did not -- a tab per
test step instead of per selection, values under the wrong headings, the
builder's own wording instead of the PY's -- and every fix was specific to
the control that exposed it. Mirroring cannot be bolted onto generation.

## What every PY seen so far has in common

Measured on CI.9, MAR.17, RR.134 and AP.156:

| | CI.9 | MAR.17 | AP.156 | RR.134 |
|---|---|---|---|---|
| Lead tab: one text box -- description, steps, references | yes | yes | yes | -- |
| One tab per selection, headings + the item's row | yes | yes | yes | table |
| Exhibit screenshots in slots, red frames around them | yes | yes | yes | -- |
| Tickmark label + red box pairs ("TM 1", "TM2", "A") | yes | yes | yes | -- |
| One conclusion box per test step, fixed house wording | yes | yes | yes | -- |
| Note box, IA recalculation cells, embedded email | yes | yes | -- | -- |

RR.134 is the selection-table variant: no per-item tabs and no testing
documentation; the selections are rows on one tab.

## How

1. **inventory.py** (plain code) lists every PY item with its exact XML:
   text boxes (runs grouped into the box's own formats, `s0`, `s1`...),
   tickmark labels and boxes, frames, picture slots, embedded objects,
   cells. Tab roles are inferred, never configured.
2. **rewrite.py** (one model call per selection tab, one for the lead
   tab) shows the model each PY item -- what every PY screenshot showed and
   what every PY tickmark marked, both read from the PY by OCR -- beside the
   CY facts. It returns keep / replace / remove per item: same wording,
   paragraphs, labels and formats; only facts change.
3. **apply.py** (plain code) writes those edits into the SAME shapes and
   cells of a copy of the PY: text boxes keep their anchor, fill, outline,
   fonts and paragraph settings; CY exhibits are fitted into the PY slots
   and the frames resized; the PY's own tickmark shapes are moved onto the
   CY values (OCR) with the label on the same side as in the PY. Cells take
   the KIND of value the PY cell held. Embedded files (PY evidence) are
   removed; formulas recalculate on open.
4. **verify.py** -- the Mirror Check tab, right after the Summary:
   - structure: every PY box accounted for, paragraph counts and labels kept
   - leftovers: a PY number, date, amount or name in CY text that is not a
     CY fact (high) -- the failure a reader cannot spot
   - placement: exhibits or tickmarks that could not be placed
   - PY data: tabs still holding last year's population / parameters / reports
   - a final model pass comparing PY and CY item by item

Tabs: one per SELECTION (never per test step), in the PY's naming
convention; the PY tab at the same position is the template, the first PY
tab for selections beyond the PY's count; surplus PY selection tabs (last
year's items) are removed.

Failure is never a lost workpaper: a failed model call keeps the PY text
and says so; an engine failure falls back to the layout builder, then the
standalone layout.
