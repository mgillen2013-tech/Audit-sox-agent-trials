# CY Testing Agent — Design Notes

Draft design for the per-test-step CY (current year) testing agent: upload PY
(prior year) testing + CY support → deterministic extraction → per-test-step
Claude call using PY as precedent → tool-based reasoning → structured draft
conclusion → human review → workpaper.

This doc focuses on the two pieces called out as needing the most work
(tool schemas, conclusion output schema), and records the surrounding
decisions so they don't have to be re-derived later.

## 1. Document extraction strategy (PDFs + Excel)

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

## 2. Tool schemas

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
into `submit_conclusion` with the schema in section 3. The backend validates
it server-side (pydantic or equivalent) before accepting it as a draft.
Critically: **every `evidence_id` in `evidence_citations` must have actually
appeared in a `search_cy_support` result earlier in that same conversation.**
If it references an id that was never returned, reject and force a retry.
This is the concrete mechanism for the risk you flagged — the agent citing
evidence it never received — turned into a hard, mechanical check rather
than a hope.

## 3. Conclusion output schema

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

A later **cross-check pass** (after all test steps have a draft conclusion)
reads the full set of `submit_conclusion` outputs for the engagement and
checks for inconsistency across steps (e.g. same evidence cited as both
supporting and contradicting different steps, or contradictory confidence
levels on steps that should correlate) — this is a good candidate for the
Opus call, see section 5.

## 4. Review UI (sketch)

Per test step: narrative up top, then expandable citation cards — each
showing the extracted snippet/table plus a thumbnail/link to the actual
source region (`preview_ref`), procedure tags as pills, exceptions in a red
banner, `insufficient_evidence` in an amber banner. Approve / Edit / Reject.

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

## 5. Model choice (Sonnet vs. Opus, via Microsoft Foundry)

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

## 6. Per-test-step prompt template (outline)

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
