"""The PY workpaper is the document; the model only edits it.

Every earlier renderer GENERATED a workpaper and then tried to make it look
like the prior year. Each new control exposed another way it did not --
tab per test step instead of per selection, values under the wrong
headings, the builder's own wording in place of the PY's -- and each fix
was specific to the control that exposed it.

This package turns that around:

  inventory.py  lists everything in the PY that belongs to last year: each
                text box and its exact wording and formatting, each data
                cell, each exhibit slot, each tickmark. Plain code.
  rewrite.py    one model call per selection tab: "here is the PY item and
                the CY facts -- return the CY version with the same wording,
                structure and spacing, changing only the facts."
  apply.py      writes those edits back into the SAME boxes and cells of a
                copy of the PY, puts CY exhibits into the PY's slots and
                moves the PY's own tickmarks onto the CY values. Plain code.
  verify.py     checks the result against the PY: structure (plain code),
                PY facts left behind (plain code), and a final model pass
                comparing PY and CY item by item.

Nothing about a control's layout is assumed. It all arrives with its PY.
"""
