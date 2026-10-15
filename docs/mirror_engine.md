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

## Principle: no rules about what a workpaper contains

Two questions only: **does CY mirror PY**, and **does the CY support mirror
the PY support**. Code copies, edits and checks; it never decides what a
tab, shape or cell MEANS. That is read from the PY by the model and checked
against the workbook by code. A first version did decide things -- a tab
named "1" is one selection, a "TM 1" text is a tickmark, a box with no
label is deleted -- and each was wrong for some PY (two tabs per selection,
boxes without labels, a lead sheet in cells).

## How

1. **structure.py** -- one model call: how is this PY organised? Which tabs
   document which PY selection (one tab or several), which tabs carry
   control-wide narrative to update (lead sheets, in boxes OR cells), which
   hold last year's population / parameters / other evidence, the heading
   and record rows of each selection tab, and CY tab names in the PY's own
   convention. Code verifies every name and grouping against the workbook;
   if the reading cannot be trusted, tab-name patterns are used as a
   FLAGGED fallback.
2. **inventory.py** -- plain code: every element with its exact XML, as it
   physically is (text shape, outline, frame, connector, group, picture,
   embedded object) and every cell. No meanings.
3. **rewrite.py** -- one model call per tab that needs CY content. The
   model sees every element (its text with formats marked, what it sits on
   -- read from the PY by OCR -- and its nearest neighbour), every picture
   and what it showed, the cells, the CY facts and ALL uploaded CY support
   pages. Per element it decides keep / replace / remove / place (on a CY
   exhibit value, on a data-row heading, or beside another element) /
   connect (re-draw an arrow). Per picture: which CY exhibit, and a crop if
   the PY showed only part of a page. It also compares support: each PY
   document against its CY counterpart (same / partial / missing).
   Anything it does not address is KEPT and flagged, never deleted.
4. **apply.py** -- plain code carrying out those decisions in a copy of the
   PY: text boxes keep anchor, fill, outline, fonts and paragraph settings;
   exhibits are fitted (and cropped) into PY slots with frames resized;
   shapes are moved; arrows re-drawn; cells take the kind of value the PY
   cell held. Each CY selection gets the same group of tabs a PY selection
   had. An uploaded CY population is written under the PY's headings.
5. **verify.py** -- the **mirror check** and **support check** (shown in the app
   after a run; not written into the workbook, at the reviewer's request):
   - structure: every PY text element accounted for; lead-ins the PY set
     apart by FORMAT (bold "Test Step 1:") kept
   - leftovers, the safety net: every number and date (any format) from the
     PY's working tabs, and every value of its record rows, looked for in
     every cell and shape of the CY working tabs; present and not a CY fact
     = left over, wherever it is. Capitalised phrases that also appear on
     the PY's own documents are candidates the final model pass judges
     (person vs job title is a judgment)
   - placement: anything that could not be placed or was not decided on
   - tickmarks: across all of a selection's tabs, every mark CITED in CY
     text ("See TM2", "per A") must be PLACED on the CY evidence, and a mark
     the PY cited must still be cited. The PY's marks are read from the PY
     itself -- a short shape whose exact text is cited inside other PY text
     -- so any convention works ("TM 2", "A", "^"). The edit call declares
     each tickmark, the value it marks and what cites it; when a value is
     not found on the page, a short second-look call reads the page lines
     and picks the exact text; the final model pass checks each placed
     value is what the citing sentence says it marks
   - support: the model's PY-vs-CY document comparison
   - PY data: tabs still holding last year's population/parameters/reports
   - a final model pass comparing PY and CY text item by item

Failure is never a lost workpaper: a failed call keeps the PY tab as it was
and says so; an engine failure falls back to the layout builder, then the
standalone layout.
