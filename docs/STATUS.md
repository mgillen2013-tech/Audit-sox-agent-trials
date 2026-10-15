# Where things stand -- read this first

Handoff note for the next working session. Last updated after the first
real-model run of the mirror engine (code at `6e08231`).

## The direction (agreed with the reviewer, do not drift from it)

- **Mirror the PY. No rules.** The prior-year workpaper is the document;
  CY is made by editing a copy of it. Code copies, edits and checks; it
  never decides what a tab, shape or cell *means* -- the model reads that
  from the PY, and code verifies the reading against the workbook.
- **Two questions only:** (1) does CY mirror PY -- conclusion wording,
  spacing, layout, everything; (2) does the CY support mirror the PY
  support.
- **Never fix for one control.** 172 controls; every fix must hold for a PY
  we have not seen. If a fix names a specific tab, word or layout, it is
  the wrong fix.
- **Never delete what has not been decided; flag it.** Never lose a paid
  run's workpaper; fall back and say why.

## How bugs are fixed: improve the base, not the instance

We are building a BASE that works for any PY and any support, not a list
of patches. Every bug a run exposes is treated as one example of a CLASS
of bug, and is fixed at the class:

- **Ask "what assumption made this possible?"**, not "what is wrong with
  this file?". A value under the wrong heading was not a 156 problem; it
  was "the code assumed column order" -- so columns are now matched by
  heading for every PY. A copied tab breaking Excel was not a CI.9 problem;
  it was "copies keep links they cannot resolve" -- so every copy is now
  made self-contained, and a final integrity check refuses any file with a
  broken link, whatever caused it.
- **Remove the assumption, or move the decision to the PY.** If code had
  to guess something about the workpaper (a tab's role, what a shape is, a
  heading row, a cell's type), the fix is to read it from the PY -- the
  model reads, code verifies -- not to add a better guess.
- **Add a net that catches the whole class**, so the next variant shows up
  even if the fix missed it: the integrity check (any unopenable file), the
  leftover scan (any PY fact left anywhere), the Support Check (any PY
  document without a CY counterpart), "not addressed" flags (anything the
  model skipped).
- **Test the class, not the file**: fixtures are built to the general shape
  (two tabs per selection, headings below a title, absolute link paths,
  no pictures), and each new test is shown to fail on the old code.
- **The test for any fix:** would it still hold for a PY or a support pack
  we have never seen? If it names a specific tab, word, control or layout,
  it is the wrong fix.

## Findings from examples

Kept separately in `docs/example_findings.md` -- what each real example
taught us, as general lessons, never as code aimed at that example. Add to
it after every run on a new example.

## How it works now

`agent/mirror/` (design: `docs/mirror_engine.md`):
structure call (model reads the PY's organisation) -> tabs laid out the
PY's way, one group of tabs per selection -> per-tab edit call (model
decides every element; code applies) -> Mirror Check + Support Check tabs
(leftover scan over every cell and shape, support comparison, final model
pass). Fallbacks: layout builder (`agent/ooxml/`), then standalone layout.

## Model calls: we have room to use more

Run costs have been kept low (Sonnet deployment, prompt caching, and the
token breakdown after each run shows where every token goes). That leaves
headroom to lean on the model more where JUDGMENT is the bottleneck,
instead of writing code to approximate it -- which is exactly where the
rule-based versions went wrong. Candidates, in the spirit of "read it from
the PY, verify in code":

- a second look when a check flags something (re-place a tickmark that was
  not found, re-read a support page the first pass missed) rather than
  leaving it for the reviewer
- reading support pages directly (vision) to match them to PY documents
  and to choose crops, instead of relying on OCR text alone
- a per-selection Support Check that also confirms each CY document
  actually evidences the attribute the PY used it for
- a stronger final review pass over the whole CY workpaper against the PY

Baseline from a real mirror-engine run (Sonnet): about **$0.27 for the
whole control** -- output 15,885 tokens ($0.16, 60% of cost), fresh input
24,367 ($0.05), cache writes 18,848 ($0.05), cache reads 54,512 ($0.01).
Roughly 130K cost-weighted units against a 300K cap PER STEP. Dollar
figures use the price table in agent/costs.py, an estimate -- Foundry
billing is authoritative. Shape of the spend: output is ~5x the price of
input, so calls that READ more (images, more of the PY) are cheap; calls
that WRITE more are what cost. Prefer extra calls with short answers.

Guardrails still apply: every model answer is validated by code before it
is used, nothing is deleted on a model's say-so without being reported,
and the per-run token caps stay. Watch the token breakdown as calls are
added -- spend should go where it buys accuracy.

## Status

- **Proven with a real model:** the reviewer ran the mirror engine end to
  end with a real model and reported the output as much improved over the
  earlier builders. This is now the base to build on, not replace.
- Before that it was trialled with a scripted stand-in on four real PYs
  (CI.9, MAR.17, RR.134, AP.156 -- kept out of the repo).
- Next: work from the real output's **Mirror Check** and **Support Check**
  tabs. Each flag that recurs is a class of problem -- fix it in the base
  (see "How bugs are fixed"), never for the one control.
- Tickmark alignment added: a mark cited in text must be on the CY
  evidence (and on the value the text means); see docs/mirror_engine.md.
  Proven on the scripted stand-in only -- check it on the next real run.
- Not yet seen: a PY with its lead sheet in cells, or tickmarks typed into
  cells (Wingdings ticks). Ask for one.
- Visual check in Excel is the user's -- this container cannot open xlsx.

## Practicalities

- User runs on Windows from a downloaded ZIP:
  `pip install -r requirements.txt` then
  `python -m streamlit run cy_testing_app.py --server.address localhost`.
- Foundry: Model box = the deployment NAME; use **Test connection** first.
- Repo is PUBLIC: never commit workpapers, support or listings
  (`.gitignore` blocks xlsx/pdf/csv). Test fixtures are built, not copied.
