# Layout is read from the PY workpaper, never computed

## The rule

The prior-year workpaper is not just a style reference and not just a
source of cell formats. **Its drawing XML is the layout specification.**
Where the narrative box sits, where each exhibit sits, how large each one
is — all of it already exists, per control, in the file we are handed.

Any code that decides *where to put something* is a bug in waiting. The
correct move is always: find where the PY put the equivalent object, and
put ours there.

## Why this keeps mattering

Three separate complaints on generated output turned out to be the same
root cause — a builder that computes positions instead of mirroring them:

- "the output is messy" — it did not look like last year's
- large blank regions on the Sample tab, with exhibits stacked in a
  column down the left
- exhibits appearing in an order and at sizes nobody chose

`build_sample_drawing` stacks objects vertically from `start_y_emu` using
`gap_emu` / `image_gap_emu` constants. The real PY workpaper does nothing
of the sort. Measured from `T0.SS.PTP.AP.156.01`'s own drawing2.xml:

```
TextBox 1  (narrative)   cols  0-6    rows  3-14     top-left
Picture 2                cols  6-19   rows  3-24     top-middle
Picture 5                cols 19-36   rows  2-40     top-right
Picture 3                cols  0-6    rows 26-70     bottom-left
Picture 4                cols  7-19   rows 25-68     bottom-middle
Picture 6                cols 21-34   rows 40-68     bottom-right
```

That is a grid, roughly three across and two down. A vertical stack leaves
the entire right-hand half of the sheet empty, which is exactly what a
reviewer circled on a real output and labelled "so much blank space".

## What this means in code

1. Parse the template's `drawing<N>.xml` for the Sample sheet on load.
2. Derive an ordered list of SLOTS: every `<xdr:pic>` anchor becomes an
   image slot (position + extent), the narrative `<xdr:sp>` becomes the
   narrative slot.
3. Place this year's evidence into those slots in order. Do not compute a
   position for anything that has a slot.
4. Only when CY has MORE exhibits than PY had slots does anything get
   derived — and even then by continuing the template's own grid rhythm
   (same widths, next row band), never by a fresh algorithm.

## Why this is also the answer for 172 controls

Every control has its own PY workpaper with its own layout. A computed
layout needs per-control tuning that nobody will do 172 times. A mirrored
layout needs none: the spec arrives with the input, and a control whose PY
looks nothing like this one still comes out right.

## Missing evidence is a finding, not a layout event

Slots only line up when this year's evidence matches last year's in count
and kind. It usually will -- the same control tested the same way gathers
the same documents -- but when it does not, the two cases are NOT
symmetric:

**More exhibits than PY had slots** is a layout question, and the layout
answers it: continue the template's own grid rhythm (same widths, next row
band). No judgment required.

**Fewer** is not a layout question at all. If the prior year had an
approval email and this year does not, that is a gap in the evidence, and
the agent should say so -- via `request_additional_support`, or as an
exception if the control cannot be concluded without it. It then appears
on the Summary tab where a reviewer is already looking.

An earlier draft of this document proposed leaving the slot visibly empty
"so a reviewer can see something is missing". That was wrong, for two
reasons:

1. **A blank rectangle is a weak and ambiguous signal.** A reviewer
   reasonably reads whitespace as a rendering glitch, not as a finding.
   Missing support is one of the most consequential things a test can
   surface and it should not be communicated by an absence.
2. **It puts judgment in the renderer.** The whole architecture rests on
   the agent deciding and the builder drawing. Having the builder
   communicate "evidence is missing" through empty space is the renderer
   making an audit assertion by implication -- and it would do so even
   when the agent had a perfectly good reason for the difference (the
   control changed, the vendor moved to a portal, the document was merged
   into another).

So: the layout mirrors PY, and any DIFFERENCE between what PY had and what
CY produced is reported by the agent. There is a useful consequence -- the
slot inventory is a list of the document types the prior year relied on,
which is exactly the checklist an agent should test its own evidence
against before concluding. Handing it that list turns a layout detail into
a real completeness procedure.
