# CY Testing Agent — Design Notes

Draft design for the per-test-step CY (current year) testing agent: upload PY
(prior year) testing + CY support → deterministic extraction → per-test-step
Claude call using PY as precedent → tool-based reasoning → structured draft
conclusion → human review → workpaper.

This doc focuses on the two pieces called out as needing the most work
(tool schemas, conclusion output schema), and records the surrounding
decisions so they don't have to be re-derived later.

## 0. Enterprise & control environment context (feeds the system prompt)

This isn't background for developers to skim — it's a fixed block that goes
into the system prompt, in the cacheable prefix from section 6. The model
should have this context on every call so it can correctly read
company-specific terminology and structure in the evidence, instead of
guessing at abbreviations or misreading what's normal. Because it's
identical across every test step and every engagement, it caches almost for
free after the first call — there's no real cost argument for leaving it
out.

Draft text for that block (tune the specifics against the actual filed
10-Ks and internal system documentation before finalizing — the company
scale/segment facts below were pulled from public filings via search, not
read directly from the source document, and the system names are as
provided, not independently verified):

```
COMPANY CONTEXT

BrightView Holdings (NYSE: BV) is the largest provider of commercial
landscaping services in the United States, operating through two reportable
segments:
- Maintenance Services: recurring, largely seasonal/evergreen landscaping
  work (mowing, irrigation, tree care, snow removal) under ongoing
  contracts.
- Development Services: project-based landscape architecture and
  construction for new facilities and major redesigns.

Operating model: 280+ branches perform the work locally (crews, equipment,
customer relationships); accounting and financial reporting are centralized
at corporate. Most SOX controls sit at corporate; a smaller set are
branch-level/transactional (e.g. purchase approvals, timekeeping, system
access, asset/vehicle controls) and are tested via samples drawn across
branches.

Financial systems referenced in evidence:
- E1 (JD Edwards EnterpriseOne) — the core financial/ERP system.
- BVE1 — a separate branch-facing instance of E1 used to reconcile branch
  activity back to corporate E1.
- OneStream — corporate consolidation and financial reporting.

When evidence references one of these systems, a report name, or a branch
number, treat that as normal company vocabulary, not something to flag as
unusual on its own.
```

## 1. Intake requirements

The original sketch of this ("simple form: 2 uploads + your fields") was
never actually pinned down. It has to be, because the tool loop in section
4 assumes a sample population already exists to check evidence against —
and nothing upstream of that said where it comes from.

**Intake is per control**, not per engagement — PY workpapers are organized
by control, and a control's 2-5 test steps share one PY precedent. Three
inputs, plus a small metadata form:

1. **PY testing workpaper** (PDF/Excel) — control objective, test step
   language, PY sample/support, PY conclusion. Parsed once, reused as
   precedent for every test step under this control (section 2).
2. **CY sample list** — this is the piece that was missing. See schema
   below. Keeps sample *selection* a human decision (the auditor already
   picked "these 25 of 340") — the tool's job is evaluating evidence against
   a known sample, not deciding what to sample. That's a deliberate scope
   boundary, not a placeholder for automating selection later.
3. **CY support evidence** — PDFs/Excels/screenshots/emails covering the
   sampled items, extracted per section 2.
4. **Metadata**: `control_id`, `period_under_review`, preparer name, and
   `control_objective` / test-step text *only if* it can't be reliably
   pulled from the PY workpaper — extraction should attempt this first
   rather than making the auditor re-type it, with a confirm/edit step
   before the control moves into the tool loop.

### CY sample list schema

One file per control, one row per sampled item, tagged to the test step it
belongs to:

```
SampleItem = {
  sample_id: str                    # "S01".."S25", stable within a test step
  test_step_id: str
  identifying_details: str          # free text, e.g. "PO #48213, Branch 35210, $12,450, 3/14/2026"
  key_fields: dict[str, str] | null # optional structured columns if the sample list has real
                                     # columns (po_number, branch, date, amount, ...) instead of
                                     # one free-text description — lets search_cy_support query
                                     # on a specific field rather than a fuzzy description match
}

SamplePopulationManifest = {
  test_step_id: str
  population_description: str       # e.g. "All POs > $5,000 issued Oct 2025-Sep 2026" -- may be "" (see below)
  population_size: int | null
  sample_size: int
  selection_method: "random" | "haphazard" | "judgmental" | "all_items" | null
  samples: [SampleItem]
}
```

`required_sample_ids` in `check_sample_coverage` (section 3) is just the
`sample_id` list from this manifest for the given test step — this is what
resolves the open question of where that input comes from.

**Built, two paths.** `agent/intake.py` has both:
- `parse_sample_list` / `build_manifests_from_rows`: the clean fixed-column
  format above (`test_step_id` / `sample_id` / `identifying_details` /
  `population_description` / `selection_method` required, `population_size`
  optional). Used by the CLI's `control.json` workflow. `sample_size` is
  computed from row count, never read from a column, so it can't drift out
  of sync with the actual list.
- `build_manifest_from_any_columns` / `read_excel_rows`: **what the
  Streamlit app actually uses**, after a real run surfaced that this
  assumption was wrong — a real sample/population export (an E1 AP payment
  extract, in the case that broke it) has columns like `invoice number
  f0411.vinv`, not `identifying_details`, and has no
  `population_description`/`selection_method` concept at all, because
  those are audit judgments, not something the source system tracks.
  Forcing a rename before upload was exactly the friction this app exists
  to remove. `sample_id` is auto-detected from a likely id-ish column
  (falling back to row position); `identifying_details` and `key_fields`
  are built from every column present.

**One workbook, two tabs (built):** a further real-run correction — the
app originally collected `population_description` / `selection_method` /
`population_size` as manual per-step form fields, and a separate "sample
list" file upload per step. A real run showed the actual cost of that: the
model's only signal about sample size was `check_sample_coverage`'s bare
count, with zero context for *why* it was that size, so a correct
1-item judgmental sample got treated as suspicious purely because a CY
support filename said "selection 1" (implying more might exist). The fix
wasn't a better prompt — it was giving the model a real number to check
against. The intake is now: **one Excel workbook per control**, one tab
holding the full population, another holding the sample selections
(`read_excel_rows(path, sheet_name=...)` targets either tab of the same
file) — matching how these files actually arrive in practice, per the
user directly. `population_size` is computed by counting the population
tab's rows, never typed. `selection_method` is dropped from the UI
entirely (it was collected but never actually reached the model or the
workpaper — see `_render_sample_line` below); the schema field is now
optional (`| null`) rather than forced to a fake-good value. The whole
workbook is also added to CY support evidence, so the population tab is
searchable and can back an IPE completeness/accuracy conclusion.

`agent/loop.py`'s `build_user_turn` now renders a `Sample:` line from
`TestStepRequest.sample_size` / `population_size` (populated in
`run_control.py` from the manifest): when population size is known, it
tells the model outright that the sample is correct and complete, so it
stops treating a small-but-intentional sample as a red flag; when it's
unknown, it explicitly tells the model to use
`request_additional_support` instead of guessing.

`agent/run_control.py` wires this together with real PY/CY file extraction
(`agent/extraction`) into one test-step run per step in a small
`control.json` spec (see `agent/control.example.json`) for the CLI, or
directly from the Streamlit form's uploads for the app (`iter_control_results`
takes pre-built manifests via a `sample_manifests` argument, skipping the
file-based path entirely). PY excerpts still aren't sliced per test step
(every step sees the whole extracted PY file) — real files now drive the
whole loop end to end either way.

## 2. Document extraction strategy (PDFs + Excel)

Everything downstream reasons over a normalized `EvidenceItem`, never over raw
bytes. Claude is never handed a PDF or workbook directly.

```
EvidenceItem = {
  evidence_id: str            # stable id, e.g. "ev_003"
  source_file: str            # original filename
  source_type: "excel_table" | "excel_cell" | "pdf_text" | "pdf_table" | "image_ocr"
  location: str                # "Recon.xlsx!Sheet1!B12:F20" or "support.pdf p.3 (bbox 120,340,500,600)"
  extracted_text: str | null
  extracted_table: list[list[str]] | null
  extraction_confidence: float # 1.0 for native text/cells, OCR confidence for scans
  preview_ref: str             # pointer to a cell-range screenshot / page image / crop, for the review UI
}
```

Per file type:

- **Excel (openpyxl, already a dependency):** walk every sheet, detect header
  rows heuristically, emit one `EvidenceItem` per detected table plus the raw
  cell dump with coordinates (so a citation can point at `Sheet1!B15`, not
  just "the spreadsheet"). Also pull cell fill color and comments —
  preparers color-code exceptions and leave comments, and that's real signal
  extraction shouldn't discard. `extraction_confidence = 1.0` always (it's
  structured data, not inferred).

  **Built (fixed after a real run):** annotated-cell extraction originally
  emitted one `excel_cell` EvidenceItem per colored/commented cell,
  unconditionally. A real population tab had 3 sampled rows highlighted
  across every column (a normal preparer convention: mark the whole
  selected row within the population listing, not one cell) — that turned
  into 57 individually-meaningless fragments (`value: 'PK'; fill_color:
  FFFFFF00`, no column label, no relation to the other 19 pieces of the
  same row), on top of the row's real content already sitting in the
  table extraction. The model burned its entire tool-call/token budget
  (200K tokens, 20 calls, no conclusion) trying to chase these down
  individually — this was the actual root cause, not a prompting or
  budget-size problem. Fixed: when most of a row's populated cells share
  the *same* highlight color (≥60%, ≥3 cells — `_WHOLE_ROW_FRACTION` /
  `_WHOLE_ROW_MIN_CELLS` in `agent/extraction/excel.py`), it collapses to
  ONE row-level EvidenceItem ("Row 3: entire row highlighted... likely
  marks a selected/sampled item... see the table extraction for full
  values") instead of one per cell. A comment on a cell is never absorbed
  into that collapse — it's real, specific content and always stays its
  own item, even inside an otherwise-collapsed row. A single flagged cell
  (the genuine exception case, e.g. the PO-testing fixture) is
  untouched — collapsing only triggers on a genuinely whole-row pattern.
  On the real file this dropped 57 fragments to 3, and the CY evidence
  pool for the whole control (4 support files) to 16 items / ~10K
  characters total.
- **Native (text) PDFs:** `pdfplumber` or `PyMuPDF` for text + table
  extraction, keeping page number and bounding box per chunk so citations are
  precise and the review UI can render the exact region.

**OCR (built — `agent/extraction/ocr.py`):** image-only pages used to stay
permanently unreadable — they became `image_ocr` placeholders with
`extracted_text=None`, and a real run concluded on an approval email alone
because the invoice and payment PDFs were scans. Now each such page is
transcribed by **Claude vision**, using the client the run already has.
Vision rather than Tesseract on purpose: the user runs this on Windows,
where Tesseract means a separate binary install and PATH setup, and the
documents in this bucket (scanned invoices, E1 screenshots, remittances)
are exactly the layout-heavy material Tesseract handles worst.

This is the only place extraction talks to a model, and it stays honest
about it: OCR'd text is prefixed `[OCR transcription of a scanned page]`
and carries `extraction_confidence = 0.8`, not the 1.0 of a native read,
so a reviewer can tell which citations rest on a transcription. The prompt
forbids inferring or completing values and requires `[illegible]` for
anything unreadable — on audit evidence an invented figure is worse than a
missing one. Cost is bounded: one call per image-only page, once per run
(not per turn), ~1-2K tokens each; pages that already have text are never
re-read. Every failure path (render error, API error, empty response)
leaves the original placeholder intact, so a failed OCR degrades to the
old behavior rather than breaking a run. Toggleable in the app.
- **Scanned PDFs / embedded screenshots / emails saved as PDF:** OCR
  (Tesseract or a hosted OCR API) per page/region. Keep the OCR confidence
  score on the `EvidenceItem` — this is what lets the agent legitimately say
  "insufficient evidence, source illegible" instead of confidently
  hallucinating off garbled text. Always keep the source image alongside the
  OCR text so a human can eyeball it in review.
- **Emails:** if they arrive as `.msg`/`.eml` rather than PDF/screenshot,
  extract body + attachments as separate `EvidenceItem`s rather than one
  blob — an attachment is usually the actual evidence, the email body is
  context.

Extraction is a batch step that runs once per upload, before any Claude call.
The per-test-step agent only ever sees the `EvidenceItem` list (scoped to
that engagement) through the `search_cy_support` tool below.

## 3. Tool schemas

Four tools, bound per test-step call. Every tool call and result is appended
to the audit log by the backend automatically — not something the model has
to remember to do.

### `search_cy_support`

Keyword/BM25 (+ optional embedding) search over the `EvidenceItem`s for this
engagement. Deterministic retrieval, not model-generated — the model narrows
in on evidence, it doesn't invent what's searchable.

```
input:  { query: str, evidence_types?: [str], top_k?: int = 5 }
output: { results: [ { evidence_id, source_file, location, snippet, score } ] }
```

### `check_sample_coverage`

Pure math, run server-side — never let the model eyeball "20 of 25 samples
tested." The model submits which evidence it believes maps to which sample
items; the tool checks that against the deterministically parsed sample
population (from the CY population listing or PY sample list) and returns
arithmetic, not judgment.

```
input:  { required_sample_ids: [str], found_evidence_ids: [str] }
output: { total_required: int, total_found: int, missing: [str],
          coverage_pct: float, complete: bool }
        | { error: "no_sample_list_found" }
```

The error case matters as much as the happy path — it's the deterministic
trigger that should push the model toward `request_additional_support`
rather than guessing.

**Built (server-enforced, not just prompted):** on a sampled test step, a
`satisfied` submit_conclusion is refused unless a successful
`check_sample_coverage` call earlier in that conversation showed
`complete: true` — same mechanism as the fabrication guard. The rejection
message names the missing samples and points to the valid outs
(cover the gaps, or downgrade the conclusion / request additional support).
`not_satisfied` and `insufficient_evidence` are never gated on coverage —
an exception found in one sample doesn't require the rest to be tested
before flagging it, and blocking the "I don't have enough" close on a
coverage check would defeat its purpose.

### `flag_exception`

Structured side-channel for exceptions, independent of the prose narrative,
so exceptions are queryable/reportable on their own.

```
input:  { test_step_id: str, description: str, evidence_ids: [str],
          severity: "low" | "medium" | "high" }
output: { exception_id: str, recorded: true }
```

### `request_additional_support`

The "insufficient evidence" escape valve as a first-class action, not a
buried text field — this is what keeps it as cheap an output as "pass."

```
input:  { test_step_id: str, description: str,
          reason: "missing" | "illegible" | "ambiguous" | "insufficient_sample" }
output: { request_id: str, recorded: true }
```

### `submit_conclusion` (the forced final step)

The model doesn't just stop talking — its final turn must be a tool call
into `submit_conclusion` with the schema in section 4. The backend validates
it server-side (pydantic or equivalent) before accepting it as a draft.
Critically: **every `evidence_id` in `evidence_citations` must have actually
appeared in a `search_cy_support` result earlier in that same conversation.**
If it references an id that was never returned, reject and force a retry.
This is the concrete mechanism for the risk you flagged — the agent citing
evidence it never received — turned into a hard, mechanical check rather
than a hope.

## 4. Conclusion output schema

This is what both the review UI and the audit log are built on.

```json
{
  "test_step_id": "TS-4.2",
  "control_objective_ref": "CO-4",
  "conclusion": "satisfied",            // "satisfied" | "not_satisfied" | "insufficient_evidence"
  "narrative": "The test step is satisfied: the recalculated interest accrual in ev_012 (Recon.xlsx!C4:C18) matches the GL balance in ev_015 (GL_Export.pdf p.2), and reperformance of the reconciliation ties out with no variance.",
  "evidence_citations": [
    {
      "evidence_id": "ev_012",
      "source_file": "Recon.xlsx",
      "location": "Sheet1!C4:C18",
      "quote_or_summary": "Accrual recalculation, ending balance $482,110",
      "relevance": "Recalculated figure used to test the accrual"
    }
  ],
  "procedures_performed": ["recalculation", "reperformance", "inspection"],
  "ipe_completeness_accuracy_status": "validated",   // "validated" | "not_validated" | "not_applicable"
  "ipe_completeness_accuracy_evidence": ["ev_009"],
  "exceptions": [],
  "additional_support_requests": [],
  "confidence": "high",                 // "high" | "medium" | "low"
  "confidence_rationale": "Native-text PDF and structured Excel source, no OCR involved, full sample coverage.",
  "sample_coverage": { "total_required": 25, "total_found": 25, "missing": [], "coverage_pct": 100.0, "complete": true },
  "model_metadata": { "model": "claude-sonnet-5", "prompt_version": "v3", "timestamp": "...", "tool_call_count": 6 }
}
```

Rules baked into validation, not just prompt instructions:

- `evidence_citations` must be non-empty **unless** `conclusion ==
  "insufficient_evidence"` — you cannot claim "satisfied" or "not_satisfied"
  with zero citations.
- Every cited `evidence_id` must trace back to a `search_cy_support` result
  in-transcript (see above).
- `conclusion == "insufficient_evidence"` is validated as equally
  well-formed as any other value — no extra hoops, no "are you sure" — so
  there's no structural bias toward false completion.
- `exceptions` / `additional_support_requests` hold ids from the
  corresponding tool calls, not free text restating them, so the UI and any
  downstream report can query them directly.
- `ipe_completeness_accuracy_status` is a required tri-state, not a boolean.
  Most evidence here comes out of E1, BVE1, or OneStream as a report (aging
  reports, PO receipt matching, GL exports) — whether that report's
  completeness/accuracy was itself validated is one of the most common gaps
  in SOX testing. A first live run surfaced exactly why a boolean isn't
  enough: the model needed to say "this evidence relies on a system report,
  and I could **not** obtain completeness/accuracy support for it" — a plain
  `relies_on_system_generated_report: bool` forces
  `ipe_completeness_accuracy_evidence` to be non-empty whenever reliance is
  true, so the model got pushed into citing the report itself just to
  satisfy the schema, then had to add a narrative disclaimer explaining that
  citation wasn't real validation. `not_validated` says the gap directly, as
  a first-class, equally-valid answer — same principle as
  `insufficient_evidence` not being a second-class conclusion. Only
  `"validated"` requires `ipe_completeness_accuracy_evidence` to be
  non-empty (citing e.g. a record-count tie-out, a total agreeing to a
  source system); `"not_validated"` and `"not_applicable"` both require it
  to be empty — the field means "the evidence that proves validation
  happened," never "the report being relied on."

A later **cross-check pass** (after all test steps have a draft conclusion)
reads the full set of `submit_conclusion` outputs for the engagement and
checks for inconsistency across steps (e.g. same evidence cited as both
supporting and contradicting different steps, or contradictory confidence
levels on steps that should correlate) — this is a good candidate for the
Opus call, see section 6.

## 5. Review UI (sketch)

Per test step: narrative up top, then expandable citation cards — each
showing the extracted snippet/table plus a thumbnail/link to the actual
source region (`preview_ref`), procedure tags as pills, exceptions in a red
banner, `insufficient_evidence` in an amber banner. Approve / Edit / Reject.

**Built (partial):** `cy_testing_app.py` is a local Streamlit form covering
intake + run, not yet review — fill in control/test-step details, upload
the PY testing file and CY support files, edit the sample list as an
in-browser table (no more hand-building a control.json or an Excel sample
list by hand, though you still can upload one), click run, see each
conclusion as it lands. What it does NOT have yet: citation-card source
previews, Approve/Edit/Reject, or the immutable-draft-vs-human-edit
separation from section 4 above — right now results are display-only. It's
a UI shell over `agent.run_control.iter_control_results`, the same tested
orchestration the CLI uses, not a second implementation.

The one integrity rule worth enforcing in the UI layer, not just the prompt:
a citation card that can't resolve its `evidence_id` to a real
`EvidenceItem` should render as a hard error state, not silently disappear —
that's your last line of defense against fabricated evidence slipping past
the schema validator.

Edits are stored as a separate record from the agent's original draft — the
agent's `submit_conclusion` output is immutable in the audit log, the
human's final workpaper text is a distinct, linked record. That's what makes
"what did the agent actually say vs. what did the human change" auditable
after the fact.

## 6. Model choice (Sonnet vs. Opus, via Microsoft Foundry)

Claude is GA on Microsoft Foundry with prompt caching supported — but only
through the native `/anthropic/v1/messages` surface, not the
OpenAI-compatibility layer, so the backend should call that surface
directly to keep caching.

Recommendation:

- **Sonnet for the per-test-step tool loop.** This is your highest call
  volume (test steps × engagements, each with several tool round-trips), and
  Sonnet is more than capable of "does this evidence satisfy this test step"
  reasoning when the tools do the deterministic heavy lifting (sample math,
  retrieval). Structure the system prompt so the invariant parts — tool
  definitions, the "insufficient evidence is a first-class answer" framing,
  the PY-as-precedent-not-proof instruction — sit in a cacheable prefix,
  since they repeat identically across every test step in an engagement.
- **Opus for the cross-check pass.** Low call volume (one per engagement,
  not one per test step), higher-value reasoning (spotting inconsistency
  across many already-drafted conclusions), and a good place to also
  re-run any step that came back `insufficient_evidence` or with a flagged
  exception as a "second opinion" before it goes to the human reviewer.

## 7. Per-test-step prompt template (outline)

- **System:** role framing (audit testing agent, drafts only, human
  approves), the non-negotiables (cite only tool-returned evidence, never
  invent it; `insufficient_evidence` is exactly as valid an output as
  `satisfied`), and the PY-is-precedent-not-proof instruction — PY shows
  format, approach, and what "satisfied" looked like structurally, but PY
  having passed says nothing about whether CY's evidence supports the same
  conclusion this year.
- **Turn content:** control objective text, test step text, PY conclusion +
  relevant PY support excerpts (already extracted, same `EvidenceItem`
  shape), then an instruction to use `search_cy_support` before drawing any
  conclusion — no concluding from the prompt content alone.
- **Forced close:** the turn only ends via `submit_conclusion`; no free-text
  final answer accepted.

## 8. Cost circuit breaker & failure-preserving audit trail

**Built:** `agent/loop.py` (`IncompleteRunError`, `MAX_TOTAL_TOKENS`,
`run_test_step`'s `max_total_tokens`/`on_turn` params), `agent/run_control.py`
(`iter_control_results`/`run_control`'s `max_total_tokens`/`on_turn` params,
`IncompleteRunError` handling), `cy_testing_app.py` (sidebar spending cap,
live per-turn progress, partial-audit-log display on failure).

A real run hit `MAX_TOOL_ITERATIONS` (15) without ever calling
`submit_conclusion` — 34 API requests, ~13.7M tokens, roughly $65 — and came
back as a bare `RuntimeError` string with the entire audit log discarded.
Two separate problems, both now fixed:

1. **No spend cap other than iteration count.** 15 iterations is not a cost
   bound — each iteration can itself be an arbitrarily large request (a
   badly-chunked extraction, an uncapped PY excerpt). `run_test_step` now
   also tracks cumulative token usage (`response.usage`, every field summed:
   input, output, cache write, cache read) turn over turn and aborts via
   `IncompleteRunError` the moment it crosses `max_total_tokens` (default
   `MAX_TOTAL_TOKENS = 300_000`), independent of and generally well before
   the iteration cap. The one exception: a turn that actually reaches
   `submit_conclusion` is always allowed to return normally even if its own
   usage pushes the running total past the cap — that call already happened
   and was already paid for, and discarding a successful conclusion at the
   last second would be strictly worse. The cap is user-adjustable in the
   Streamlit sidebar ("Spending cap"), since a legitimately complex step and
   a runaway one aren't distinguishable in advance.
2. **Total loss of partial work on failure.** Every tool call before an
   abort was real, billed API usage that had already found *something* —
   discarding it on failure wasted both the money and the information.
   `IncompleteRunError` (subclass of `RuntimeError`, so old bare `except
   Exception` callers still catch it) carries `audit_log` (every tool call
   made before the abort), `reason` (`"token_budget_exceeded"`,
   `"max_iterations"`, or `"model_refusal"`), `tokens_used`, and `turns_used`.
   `iter_control_results` special-cases it and includes all four on the
   yielded error dict instead of just `{"error": str}`; the CLI
   (`run_control.main`) writes them to the step's output JSON, and the
   Streamlit app renders them in an expander ("Tool calls made before
   failure") under the error banner so a failed step still shows what was
   searched and what was found.

**Cost-weighted budget (built, corrected after a real run):** the cap
counts fresh-input-token *equivalents*, not raw usage-field sums: cache
writes ~1.25x, cache reads ~0.1x, output ~5x — the standard Anthropic
price ratios. The first version summed every field raw, which counted
cache reads at 10x their real price; a real, legitimately-progressing run
(sensible targeted searches, small evidence pool) got aborted at "300K
tokens" that were mostly cheap cache re-reads — the counter said budget
blown while the actual bill was a small fraction of 300K fresh tokens.
This is also why the app's number reads lower than Foundry's raw token
dashboard.

**Turn efficiency (built, same real run):** two changes that attack why a
model burns turns rather than how many it gets. (1) `search_cy_support`
results now return items nearly whole (2,000-char cap with an explicit
truncation notice) instead of 280-char snippets — the audit log showed
the model re-querying the SAME approval email with different field names
just to see more of it through the keyhole. (2) The first user turn now
includes a **CY evidence inventory**: the complete id/location/preview
map of every extracted evidence item (capped at 40), with an explicit
instruction that anything not listed was not provided and belongs in
`request_additional_support`, not in more searches — the same run burned
turns fishing for a delegation-of-authority matrix that was never
uploaded. The inventory is a map, not retrieved evidence: the fabrication
guard still requires an item to have been returned by search before it
can be cited. Measured on the real control's files: first-turn context ≈
4K tokens all-in.

Also built: a one-time **wrap-up warning** injected into the conversation a
turn before either cliff (the last allowed iteration, or 80% of the token
budget) telling the model to stop investigating and close via
`submit_conclusion` — `insufficient_evidence` + `additional_support_requests`
explicitly named as the valid escape. This converts the worst case from
"abort with only a partial audit log" into "a real, reviewable conclusion,"
with the hard aborts above still capping spend if the model ignores it.

Still a gap, not yet built: real-time *dollar* cost (token counts are a
proxy — cache reads are billed far cheaper than fresh input tokens, so
`tokens_used` overstates actual spend on a well-cached run) and a pre-flight
cost estimate before a run starts.

## 9. Workpaper file output

**Built:** `agent/workpaper.py`, wired into both the Streamlit app (download
button that survives script reruns via `st.session_state`) and the CLI
(written next to the per-step JSONs).

One generated workpaper file per control, in the **same file type as the
uploaded PY workpaper** (PY was a PDF → PDF via reportlab; PY was Excel →
.xlsx via openpyxl). It's a clean generated document with standard workpaper
sections — control header, then per test step: conclusion + confidence,
documentation narrative, procedures performed, evidence citation table,
sample coverage, IPE status, exceptions, additional support requested, and
prepared-by metadata (model, prompt version, timestamp, tool calls) — NOT a
cell-level edit of the PY file. That was an explicit user choice:
predictable layout over a fragile in-place edit of an arbitrary workbook.
Failed/incomplete steps appear in the workpaper marked INCOMPLETE with the
abort reason and the searches attempted, so the file never looks finished
while a step is missing.

Zero LLM calls: the file is rendered entirely from the structured
ConclusionOutput the run already produced, so generation is deterministic
and free. Every sheet/page carries a DRAFT banner ("AI-prepared, pending
human reviewer approval") — same non-negotiable as the system prompt.

**Sheet organization (built, after reviewing a real generated
workpaper):** the first real output buried its own conclusions — two
full-page exhibit renders sat inline in the test-step sheet, pushing
sample coverage, IPE status, exceptions, and open support requests to
around row 140, so a reviewer had to scroll past the pictures to reach
the answers. Now: exhibits move to a dedicated `<step> - Exhibits` sheet
(the step sheet cross-references it by name), and the step sheet is
ordered **answers first** — CONCLUSION (verdict, confidence + rationale,
sample coverage, IPE status), EXCEPTIONS, ADDITIONAL SUPPORT REQUESTED —
then the supporting detail: DOCUMENTATION, PROCEDURES PERFORMED, IPE C&A
EVIDENCE, EVIDENCE CITED, PREPARED BY. Sections carry banded headers, the
verdict is color-coded (green/red/amber), and empty EXCEPTIONS /
ADDITIONAL SUPPORT sections print "None noted." rather than rendering
nothing — in a workpaper, "no exceptions" and "we didn't look" must not be
indistinguishable. The Summary sheet gained IPE status, exception count,
and open-request count columns so a multi-step control's state is legible
without opening every sheet.

**Evidence exhibits with tickmarks (built):** mirroring how a human
workpaper points at evidence, each citation gets a red tickmark letter
(A, B, C…) in the citation table, and the cited source is embedded as an
exhibit: the cited PDF page is rendered to an image with red boxes drawn
where the quoted text was located (tight boxes via pdfplumber text search
against the citation's quote; falls back to the extracted region's bbox; a
scanned/screenshot page with no text layer becomes a full-page exhibit
labeled with its letter — which is exactly the E1-screenshot case in real
support files). Excel-sourced citations get a text excerpt of the cited
range instead of an image. The evidence_id → source-region mapping comes
from re-running the deterministic extraction over the same support files
(same files, same order ⇒ same ids), so this too costs zero LLM calls.
Every exhibit is best-effort: a page that can't render degrades to the
text-only citation row, never a failed workpaper build.
