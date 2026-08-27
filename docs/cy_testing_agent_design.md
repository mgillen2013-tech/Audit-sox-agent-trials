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
  population_description: str       # e.g. "All POs > $5,000 issued Oct 2025-Sep 2026"
  population_size: int | null
  sample_size: int
  selection_method: "random" | "haphazard" | "judgmental" | "all_items"
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
  to remove. This path takes any spreadsheet — sample_id is auto-detected
  from a likely id-ish column (falling back to row position),
  `identifying_details` and `key_fields` are built from every column
  present, and population description / selection method / population size
  are entered on the form once per test step instead of expected to live in
  the file. One file = one test step's samples here, matching how a real
  export naturally exists.

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
- **Native (text) PDFs:** `pdfplumber` or `PyMuPDF` for text + table
  extraction, keeping page number and bounding box per chunk so citations are
  precise and the review UI can render the exact region.
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
