# Where things stand -- read this first

Handoff note for the next working session. Last updated at commit `6e08231`.

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

## How it works now

`agent/mirror/` (design: `docs/mirror_engine.md`):
structure call (model reads the PY's organisation) -> tabs laid out the
PY's way, one group of tabs per selection -> per-tab edit call (model
decides every element; code applies) -> Mirror Check + Support Check tabs
(leftover scan over every cell and shape, support comparison, final model
pass). Fallbacks: layout builder (`agent/ooxml/`), then standalone layout.

## Not yet proven

- **No real model call has run through the mirror engine.** Everything was
  trialled with a scripted stand-in on four real PYs (CI.9, MAR.17, RR.134,
  AP.156 -- kept out of the repo). The first real output, with its Mirror
  Check and Support Check tabs, is the next thing to look at.
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
