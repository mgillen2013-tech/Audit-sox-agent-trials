# Enhancement notes

Ideas agreed but not built. Each entry says what, why, and what to decide
before writing any code — so the decision does not have to be rediscovered
when someone picks it up.

---

## Per-control prompt considerations, from the control library

**Status:** agreed, not built. Nothing in the code has changed for this.

**What:** add a `prompt_considerations` column to the control reference
library (the AuditBoard export / control_library.xlsx described in the
library design). Its contents get injected into the agent's prompt for
that control, and only that control.

**Why it is needed at all.** The system prompt is shared by all 172
controls, so anything written there has to be true of every one of them.
That leaves nowhere to say the things that are true of exactly one:

- this vendor bills through a portal, so the approval is a screenshot and
  not an email
- the payment screen for this control settles several invoices per check,
  so an IA Calculation is normally required
- the void date column is always blank for this population; its being
  blank is not an exception
- amounts on this control are in CAD, and the extract does not say so

None of that can go in the shared prompt without being wrong elsewhere,
and none of it can be inferred from the test step text. Today it is
discovered per run, forgotten, and rediscovered.

### Decide these before building

**1. Guidance about EVIDENCE, never about the CONCLUSION.**
The same line as the memory discussion, and the reason to be careful
here. A note reading "the approval is a screenshot, not an email" tells
the agent where to look and is exactly the point. A note reading "this
control always passes" or "do not worry about the missing DOA matrix"
reaches into the judgment and quietly narrows the test — and it would do
so invisibly, because the output still looks complete.

The column should be documented as answering "what would a colleague tell
someone testing this control for the first time about the DOCUMENTS", and
reviewed on that basis. Consider validating it on load and refusing, or
at least warning on, text that reads as a conclusion.

**2. A reviewer has to be able to see it.**
If a control-specific instruction shaped how the testing was performed,
the workpaper should say so. An instruction that influenced the work but
appears nowhere in the file is not reviewable, and "why did it not test
X" becomes unanswerable after the fact. Cheapest option: reproduce the
text on the Summary tab, or in the run's audit log.

**3. Length cap, and it should be small.**
Loaded per control, so it costs tokens only on that control's runs —
negligible. But an uncapped free-text field in a spreadsheet becomes an
essay, and a long note competes with the actual evidence for the model's
attention. A couple of hundred characters is probably right; enforce it
rather than hoping.

**4. It is a snapshot, like the rest of the library.**
Exported from AuditBoard (or maintained alongside it) at a point in time.
That is a feature for a testing cycle -- the guidance should not shift
mid-cycle -- but it means someone has to own keeping it current, and the
export date should be visible.

### Depends on

The control library reader, which is also not built. This column is part
of that work rather than a separate piece: the normalized shape it loads
into (control id -> objective -> steps -> attributes -> expected
evidence) simply gains one more field.
