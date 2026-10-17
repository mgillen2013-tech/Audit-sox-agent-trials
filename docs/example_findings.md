# Findings from examples

What each real example taught us, kept apart from the code. Every entry is
written as a GENERAL lesson -- what it says about PY workpapers or support
packs in general, and what changed in the base because of it. Nothing here
is a rule aimed at one control, and nothing here is client data: no vendor
names, amounts or document numbers. The examples themselves never enter the
repo (it is public); they are used from the session scratchpad only.

Format per finding: **what we saw** -> **what it means in general** ->
**what changed in the base**.

## AP.156 (non-PO invoice approval)

- A blank column in this year's listing (a void date) shifted every later
  value one column left of its heading -> listings drop or reorder columns
  from year to year -> values are matched to the PY's headings by NAME, never
  by position (`agent/ooxml/sample_grid.py`).
- The same listing had two columns with the same heading -> duplicate
  headings are normal in system extracts -> intake keeps both ("Name (2)").
- Our own earlier output was uploaded as the next year's PY -> a previous CY
  workpaper is a valid prior year -> previous generated tabs are recognised
  by their content and replaced, not treated as PY content.
- Opening and saving in Excel renumbered the file's formats -> nothing may be
  looked up by position in styles.xml -> formats are copied from the PY cell
  being replaced.

## CI.9 (extra work invoicing)

- One tab per selection named "1", no population tab, and the selection's
  conclusion boxes for ALL test steps on that one tab -> the unit is the
  selection, not the test step -> one group of tabs per selection, as the PY
  has it.
- The tab carried printer settings; copying it for a second selection
  produced a file Excel refused to open -> copies must take every linked part
  with them -> self-contained copies, plus a final integrity check on every
  file.
- An embedded email was attached as evidence -> PYs embed files, not only
  screenshots -> embedded files are reported as PY support needing a CY
  counterpart.
- Tickmarks were labelled "TM1"/"TM2" and cited in the conclusion text ->
  tickmark styles vary by preparer -> the PY's cited marks are read from the
  PY itself, and every mark cited in CY text must be placed on CY evidence.

## MAR.17 (customer billing review)

- An IA recalculation written in cells beside the data row, with formulas
  -> PYs keep working calculations in cells -> cells other than the record
  are offered to the editor; formulas are never edited and recalculate on
  open.
- Separate Population and Parameters tabs holding last year's data, and a
  divider tab -> PY workbooks contain last-year evidence tabs -> such tabs are
  flagged for replacement (population written back when uploaded).
- A short annotation shape ("Total Amount") beside a tickmark -> not every
  short shape is a tickmark -> marks are recognised only where other text
  cites them, matched with their exact capitals.

## RR.134 (revenue recognition support)

- No testing boxes, no screenshots, no per-selection tabs: two report
  extracts and a selection table -> some PYs are a selection table plus
  reports -> the selection table is updated in place, conclusions go on the
  Summary, and the report tabs are flagged as last year's data.
- The trial's listing matched none of the PY's headings -> when nothing
  matches, the PY's headings must not be overwritten -> the row is left and a
  high flag raised instead.

## From real runs of the mirror engine

- The whole support pack was pasted onto the lead sheet as well as the
  selection tab -> a model offered every exhibit will use them anywhere it
  is allowed to -> evidence is added only to tabs where the PY had
  evidence, each CY document is shown once across the workbook, and
  anything on more than one tab is flagged on the Mirror Check.
- Extra CY support pages (more than the PY had pictures) came out in one
  long column under the PY's pictures, while the PY had its evidence side
  by side -> the same bug had been fixed in the older layout builder but
  came back through the mirror engine's own overflow code -> extra pages
  now continue the PY's picture grid (its column positions, slot size and
  rows, filling its last row first, never on another element), at least two
  across (`continue_grid` in `agent/mirror/apply.py`).

- A fillable PDF form could not be read -> a form keeps its answers in
  fields, not in the page text, and pages render with those fields blank
  unless forms are switched on -> field values are read with their printed
  labels (dynamic XFA forms from their data packet), and every page render
  goes through one form-aware renderer (`agent/ooxml/imaging.py`).
- Approvals arrive as emails -> support is not only PDFs and workbooks ->
  each .msg/.eml is turned into a PDF at upload (headers + body) and its
  readable attachments become support files of their own
  (`agent/extraction/email_convert.py`); unreadable attachments are named
  in a note so the reviewer knows evidence was not seen.

- A run tested last year's support: the conclusion reported last year's
  items as observed -> the only guard checked that cited evidence_ids were
  this year's, never the VALUES reported, and the PY-benchmark prompt drew
  the agent's attention to last year's evidence -> at submit, any date,
  number or name reported as observed that is in the PY workpaper but in
  none of this year's evidence, listing or test step is sent back
  (`_last_year_facts_reported` in `agent/loop.py`); py_comparison is the
  only field where last year's facts may appear.

- A PY with five selection tabs came out with CY support on three; the
  other two were last year's testing, unchanged -> "keep the PY content"
  was the fallback for a failed or partial edit call and for any tab the
  structure reading missed, which keeps the file valid but lets last
  year's work pass for this year's, and after the check tabs were removed
  nothing in the workbook said so -> a failed edit is retried once with
  more room; any tab still showing last year's content (failed edit, PY
  picture not replaced, PY tab with support that no edit touched) is
  coloured red, listed first on the Summary as NOT UPDATED, and named in a
  red message in the app.

- Tabs whose CY support was an Excel workbook kept last year's pictures
  -> the agent read Excel support, but only PDFs and images were offered
  to the workpaper as pictures -> cited cells are drawn as a table picture
  (header row, bold and fills kept) and pasted screenshots are taken out
  as they are (`agent/mirror/excel_exhibits.py`); macro-enabled .xlsm reads
  like .xlsx, and old .xls is refused at upload with "save as .xlsx".

## Across all four

- The first versions decided structure from tab names and shape text; each
  example broke a different guess -> structure, element meaning and tickmark
  roles are read from the PY by the model and verified in code
  (`docs/mirror_engine.md`).
