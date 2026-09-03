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

## The one thing to watch

Slots are only meaningful if this year's evidence maps onto last year's in
count and kind. It usually will — the same control tested the same way
gathers the same documents — but when it does not, the extra exhibits are
the case above, and MISSING exhibits should leave a slot visibly empty
rather than reflowing everything, so a reviewer can see that something the
prior year had is absent this year. An empty slot is information.
