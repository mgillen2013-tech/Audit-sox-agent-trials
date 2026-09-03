"""A second model pass that checks the narrative boxes against the support.

The testing agent asserts things; nothing currently checks that what it
asserted is what the evidence says. The schema-level guards catch
fabricated evidence ids and contradictory roll-ups, but they are
mechanical -- they cannot tell that "approved 02/09/2026 by Greg
Miraglia" was written beside a document that actually shows a different
approver, because both are well-formed strings.

This is the other half of the experiment in agent/narrative_styles.py: a
reviewer model reads each legend line beside the extracted text of the
evidence it cites, and says whether the line is supported. It is OFF by
default and costs one extra API call per sampled item, so it is a
deliberate choice, not a default tax.

WHY A SEPARATE PASS AND NOT MORE INSTRUCTIONS
---------------------------------------------
The testing agent has already committed to its answer by the time the
narrative exists. Asking it to check its own work in the same breath
mostly produces agreement -- it is arguing for a conclusion it just
reached, with the same reading of the same documents. A separate call
starts from the artifact rather than the intent: here is a sentence, here
is the text of the document it points at, is the sentence true of the
document. That question has an answer that does not depend on how the
first pass reasoned.

WHAT IT DELIBERATELY DOES NOT DO
--------------------------------
It does not rewrite anything, and its verdict does not gate the
workpaper. A reviewer model that silently edited an auditor's conclusion
would be worse than none: the file would carry assertions nobody chose.
It reports, and the report goes on the Summary tab next to the row it
concerns.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any

# Bounded hard. This runs once per sampled item and the evidence text can
# be long; without a cap a scanned bundle would make the review cost more
# than the testing did.
_MAX_EVIDENCE_CHARS = 6_000
_MAX_OUTPUT_TOKENS = 1_000

_SYSTEM = """\
You are reviewing a completed audit workpaper, not producing one.

You will be given the tickmark legend written on one sampled item's tab,
and the extracted text of the evidence documents those tickmarks point
at. For each lettered line, decide whether the evidence text actually
supports what the line claims.

Rules:
- Judge ONLY against the evidence text provided. If a line's value does
  not appear in it, that is "unsupported" -- not "supported" on the
  assumption it is elsewhere.
- A value that appears in a different FORMAT is still supported: dates
  and amounts are rendered differently by different systems ("02/09/2026"
  and "February 9, 2026" are the same date).
- OCR text is imperfect. A near-match with an obvious character misread is
  supported; a different number is not.
- You are not re-performing the test. Do not judge whether the control is
  effective, only whether each line is faithful to the documents.

Reply with JSON only, no prose around it:
{"lines": [{"letter": "A", "verdict": "supported"|"unsupported"|"unclear",
            "note": "<= 15 words, only if not supported"}]}
"""


@dataclass
class LineReview:
    letter: str
    verdict: str  # supported | unsupported | unclear
    note: str = ""


@dataclass
class SampleReview:
    sample_id: str
    lines: list[LineReview]
    error: str = ""

    @property
    def summary(self) -> str:
        """One cell's worth: what a reviewer needs to see on the Summary."""
        if self.error:
            return f"Not reviewed ({self.error})"
        if not self.lines:
            return "Nothing to review"
        bad = [line for line in self.lines if line.verdict != "supported"]
        if not bad:
            return f"All {len(self.lines)} tickmark(s) supported by the cited evidence"
        return "; ".join(
            f"{line.letter} {line.verdict}" + (f" - {line.note}" if line.note else "")
            for line in bad
        )

    @property
    def all_supported(self) -> bool:
        return bool(self.lines) and not self.error and all(
            line.verdict == "supported" for line in self.lines
        )


def _evidence_text(evidence_items: list[Any], wanted_ids: set[str]) -> str:
    parts: list[str] = []
    used = 0
    for item in evidence_items:
        if item.evidence_id not in wanted_ids:
            continue
        body = item.extracted_text or ""
        if not body.strip():
            continue
        chunk = f"[{item.source_file} {item.location}]\n{body.strip()}"
        if used + len(chunk) > _MAX_EVIDENCE_CHARS:
            chunk = chunk[: max(0, _MAX_EVIDENCE_CHARS - used)]
        parts.append(chunk)
        used += len(chunk)
        if used >= _MAX_EVIDENCE_CHARS:
            break
    return "\n\n".join(parts)


def review_sample(
    client: Any,
    model: str,
    *,
    sample_id: str,
    legend: list[tuple[str, str, str]],
    evidence_items: list[Any],
    evidence_ids: set[str],
) -> SampleReview:
    """Check one sampled item's legend against its cited evidence.

    Never raises. A review that fails is reported as "not reviewed" with
    the reason -- the workpaper it was checking is already built and valid,
    and losing it because the optional second opinion failed would be
    absurd.
    """
    if not legend:
        return SampleReview(sample_id, [], error="no tickmarks on this item")

    text = _evidence_text(evidence_items, evidence_ids)
    if not text.strip():
        return SampleReview(
            sample_id, [], error="none of the cited evidence had readable text"
        )

    lines = "\n".join(f"{letter} - {attribute}: {value}" for letter, attribute, value in legend)
    user = f"TICKMARK LEGEND FOR SELECTION {sample_id}:\n{lines}\n\nEVIDENCE TEXT:\n{text}"

    kwargs = {
        "model": model,
        "max_tokens": _MAX_OUTPUT_TOKENS,
        "system": _SYSTEM,
        "messages": [{"role": "user", "content": user}],
    }
    try:
        response = (
            client.messages.create(**kwargs)
            if hasattr(client, "messages")
            else client.create_message(**kwargs)
        )
        raw = "".join(
            b.text for b in response.content if getattr(b, "type", None) == "text"
        ).strip()
        # The model is asked for bare JSON but may fence it; take the
        # outermost object rather than failing on a stray backtick.
        start, end = raw.find("{"), raw.rfind("}")
        payload = json.loads(raw[start : end + 1]) if start != -1 and end != -1 else {}
        reviewed = [
            LineReview(
                letter=str(entry.get("letter", "?")),
                verdict=str(entry.get("verdict", "unclear")).lower(),
                note=str(entry.get("note", ""))[:120],
            )
            for entry in payload.get("lines", [])
        ]
        return SampleReview(sample_id, reviewed)
    except Exception as exc:  # noqa: BLE001 -- an optional check must never break the build
        return SampleReview(sample_id, [], error=str(exc)[:120])
