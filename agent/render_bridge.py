"""Translate what the agent concluded into what the OOXML builder draws.

This is the seam between the two halves of this repo. The agent
(agent/loop.py) is all judgment and no layout: it decides which evidence
supports which attribute of which sampled item, and says so as a
ConclusionOutput. The builder (agent/ooxml/) is all layout and no
judgment: give it a SampleItem and it reproduces the prior-year workpaper
with native Excel callouts. Neither knows the other exists, and this
module is the only thing that does.

The one genuinely hard translation is the ANCHOR TEXT. The agent reports
what it observed in prose -- "Approved 10/1/2025 by D. Milosevic; paid
11/10/2025, $30,000.00" -- and the builder needs a single whitespace-free
token it can find in the OCR output of a page. Picking that token badly is
not a cosmetic failure: it puts the red box around the wrong number.
_anchor_candidates() is that choice, and it deliberately reuses the rules
agent/wordboxes.py already learned the hard way on real invoices:

  - Amounts are taken WHOLE. "$30,000.00" anchors as "30,000.00", never as
    the fragments a naive split produces -- "000" is a substring of every
    round amount on a page and boxed the wrong line item on a real run.
  - Bare years are dropped. "2025" appears in every date on the document.
  - Anchors below _MIN_ANCHOR_LEN characters are dropped as
    non-identifying.

Unlike wordboxes.py, anchors here keep their punctuation. That module
normalises ("30,000.00" -> "3000000") because it matches against
normalised OCR text; the OOXML builder does a raw substring match against
the OCR token, so the anchor has to look like what is printed on the page.
"""

from __future__ import annotations

import re
from pathlib import Path

from agent.ooxml.models import (
    EvidenceImage,
    EvidenceTieOut,
    NarrativeParagraph,
    NarrativeRun,
    PopulationRow,
    RawDataRow,
    SampleItem as OoxmlSampleItem,
    WorkpaperRequest,
)
from agent.schemas import ConclusionOutput, SamplePopulationManifest

# Same patterns as agent/wordboxes.py, and for the same reasons -- see the
# module docstring. Kept as separate constants rather than imported so the
# two can diverge if the OOXML matcher's needs ever differ from the
# normalised matcher's, without one silently changing the other.
_AMOUNT_RE = re.compile(r"\d[\d,]*\.\d{2}")
_LONG_NUM_RE = re.compile(r"\b\d{4,}\b")
_ALNUM_ID_RE = re.compile(r"\b(?=[A-Za-z]*\d)[A-Za-z0-9\-]{5,}\b")
_DATE_RE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")
_YEAR_RE = re.compile(r"^(19|20)\d{2}$")

_MIN_ANCHOR_LEN = 4

# A..Z. A sampled item needing more than 26 tickmarks is not a workpaper
# problem, it is a scoping problem, so the extras are dropped rather than
# rolling over to AA -- a two-letter tickmark reads as a typo in a legend.
_LETTERS = [chr(ord("A") + i) for i in range(26)]

# "p.3" inside a location string like "invoice.pdf p.3" or
# "invoice.pdf p.3 (bbox 12,34,56,78)".
_PAGE_RE = re.compile(r"\bp\.(\d+)\b")


def anchor_candidates(text: str) -> list[str]:
    """Distinctive, whitespace-free tokens from `text`, most specific first.

    Returns every candidate rather than one pick, because the caller can
    only find out which of them is actually locatable by OCRing the page --
    and the builder reports an unfindable anchor instead of failing, so
    offering the best candidate first and letting it degrade is better than
    guessing once.
    """
    found: list[str] = []

    # Order is by RELIABILITY ACROSS DOCUMENTS, not by length. A real run
    # made that distinction concrete: an attribute observed as "Approved Inv
    # #35713082 on 02/09/2026" ranked the date first because it is two
    # characters longer -- but the approval email writes that date as
    # "February 9, 2026", so the anchor was unfindable and the tickmark was
    # dropped, while the invoice number sitting beside it was right there on
    # the page and would have matched.
    #
    # Amounts and identifiers are printed the same way wherever they appear;
    # DATE FORMATS are re-rendered by every system that touches them
    # (02/09/2026, February 9 2026, 9-Feb-26). So dates go last -- still
    # offered, because on a screenshot of a payment screen the numeric form
    # is often exactly right, but never preferred over a stable identifier.
    amounts: list[str] = []
    identifiers: list[str] = []
    dates: list[str] = []

    def add(raw: str, bucket: list[str]) -> None:
        token = raw.strip().strip(".,;:")
        if (
            len(token) >= _MIN_ANCHOR_LEN
            and " " not in token
            and token not in found
            and not _YEAR_RE.match(token)
        ):
            found.append(token)
            bucket.append(token)

    for m in _AMOUNT_RE.findall(text):
        add(m, amounts)
    # Blank the amounts before scanning for plain numbers so their digit
    # groups cannot come back as separate, far less specific anchors.
    remainder = _AMOUNT_RE.sub(" ", text)
    for m in _DATE_RE.findall(remainder):
        add(m, dates)
    remainder = _DATE_RE.sub(" ", remainder)
    for pattern in (_LONG_NUM_RE, _ALNUM_ID_RE):
        for m in pattern.findall(remainder):
            add(m, identifiers)

    # Within a bucket, longest first -- a longer token identifies one place
    # on the page more often than a short one.
    by_length = lambda xs: sorted(xs, key=len, reverse=True)  # noqa: E731
    return by_length(amounts) + by_length(identifiers) + by_length(dates)


def _page_of(location: str) -> int:
    """0-indexed page from a citation's location string.

    The agent's locations are 1-indexed ("invoice.pdf p.1"); EvidenceImage
    is 0-indexed, matching pdf2image's original convention. Getting this
    off by one silently boxes a value on the wrong page, so it is converted
    in exactly one place.
    """
    m = _PAGE_RE.search(location or "")
    return max(0, int(m.group(1)) - 1) if m else 0


def _legend_line(letter: str, label: str) -> NarrativeParagraph:
    """One legend row: a red letter, then the attribute it marks.

    The " - " separator is not cosmetic. OoxmlSampleItem's own validator
    rejects a workpaper whose callout letters are not defined in the
    narrative, and it looks for exactly this shape.
    """
    return NarrativeParagraph(
        runs=[
            NarrativeRun(text=f"{letter} - ", bold=True, color="RED"),
            NarrativeRun(text=label, color="BLACK"),
        ]
    )


_ISO_DATE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$")
_US_DATE_RE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")
# A number as it comes out of a spreadsheet read: optional sign, optional
# thousands separators, optional decimals. Deliberately NOT a general numeric
# match -- an invoice number ("35713082") is a numeric STRING and must stay
# one, or it renders right-aligned and comma-grouped where the prior-year
# workpaper shows it as plain text.
_MONEY_RE = re.compile(r"^-?\d{1,3}(,\d{3})*\.\d{1,2}$|^-?\d+\.\d{1,2}$")


def _coerce(key: str, value: str) -> object:
    """Restore the type a spreadsheet value lost on the way in.

    agent.schemas.SampleItem.key_fields is dict[str, str] -- intake reads a
    sample tab into text, which is right for searching and matching. The
    OOXML builder does the opposite: it picks the cell's Excel type from the
    PYTHON type it is handed, so a float writes a right-aligned number and a
    date writes a real date serial, while a str writes a shared-string text
    cell.

    Passing "32677.00" straight through therefore produces a left-aligned
    text cell that will not sum and does not match the prior-year workpaper
    beside it. That is invisible in any unit test built from typed
    fixtures and showed up the moment real intake data was used.

    Conservative on purpose: only what is unambiguously a money amount or a
    date converts. Bare digit strings stay strings, because in this extract
    they are identifiers -- invoice number, payee number, document number,
    cost centre -- and formatting an invoice number as "35,713,082" would be
    a visible defect in an audit deliverable.
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None

    m = _ISO_DATE_RE.match(text)
    if m:
        from datetime import date

        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = _US_DATE_RE.match(text)
    if m:
        from datetime import date

        return date(int(m.group(3)), int(m.group(1)), int(m.group(2)))

    if _MONEY_RE.match(text):
        return float(text.replace(",", ""))

    # "0" appears in Discount Available/Taken, which ARE numeric columns in
    # the template. Bare integers convert only when the column name says
    # money -- otherwise an identifier would be caught here.
    if text.lstrip("-").isdigit() and any(
        word in key.lower() for word in ("amount", "discount", "total")
    ):
        return int(text)

    return text


def _raw_data_row(manifest: SamplePopulationManifest, sample_id: str) -> RawDataRow:
    """The sampled item's own fields, keyed by template column letter.

    Column letters come from the ORDER of the sample tab's columns, which
    is the same order the population extract has -- the template's header
    row and the extract are the same query. If a control's template ever
    reorders its columns relative to the extract, this is the assumption
    that breaks, and it breaks visibly (values under the wrong headings)
    rather than silently.
    """
    item = next((s for s in manifest.samples if s.sample_id == sample_id), None)
    if item is None or not item.key_fields:
        return RawDataRow(values={})
    return RawDataRow(
        values={
            _LETTERS[i]: _coerce(k, v)
            for i, (k, v) in enumerate(item.key_fields.items())
            if i < 26
        }
    )


def sample_items_from_conclusion(
    conclusion: ConclusionOutput,
    manifest: SamplePopulationManifest | None,
    control_id: str,
    support_dir: Path,
) -> tuple[list[OoxmlSampleItem], list[str]]:
    """One OOXML SampleItem per sampled item, plus any warnings.

    Tickmark letters are assigned per ATTRIBUTE, not per citation: an
    attribute is the unit a reviewer signs off ("approval precedes
    payment"), and it is what the agent already ties to specific evidence
    via AttributeResult.evidence_ids. Letters restart at A for each sample,
    matching how the prior-year workpapers read.

    Warnings are returned rather than raised. A sample that cannot be fully
    rendered should still produce a workpaper with everything else in it --
    same rule the rest of this project follows -- but it must say what it
    could not do.
    """
    warnings: list[str] = []
    citations_by_id = {c.evidence_id: c for c in conclusion.evidence_citations}

    # Which samples exist, in the sample listing's own order.
    sample_ids: list[str] = []
    if manifest:
        sample_ids = [s.sample_id for s in manifest.samples]
    else:
        for c in conclusion.evidence_citations:
            if c.sample_id and c.sample_id not in sample_ids:
                sample_ids.append(c.sample_id)

    verdicts = {r.sample_id: r.conclusion for r in conclusion.sample_results}

    items: list[OoxmlSampleItem] = []
    for sample_id in sample_ids:
        attributes = [a for a in conclusion.attribute_results if a.sample_id == sample_id]
        if not attributes:
            warnings.append(
                f"sample {sample_id!r}: no attribute results, so it gets no tickmarks. "
                f"The evidence is still attached; only the callouts are missing."
            )

        narrative: list[NarrativeParagraph] = [
            NarrativeParagraph(
                runs=[NarrativeRun(text=conclusion.narrative, color="BLACK")]
            ),
            NarrativeParagraph(),
        ]

        # letter -> (attribute, anchor candidates, evidence ids)
        by_image: dict[tuple[str, int], list[EvidenceTieOut]] = {}
        for i, attr in enumerate(attributes):
            if i >= len(_LETTERS):
                warnings.append(
                    f"sample {sample_id!r}: more than 26 attributes; "
                    f"{len(attributes) - 26} were left untickmarked."
                )
                break
            letter = _LETTERS[i]
            narrative.append(_legend_line(letter, f"{attr.attribute}: {attr.value_observed}"))

            anchors = anchor_candidates(attr.value_observed)
            if not anchors:
                warnings.append(
                    f"sample {sample_id!r} tickmark {letter} ({attr.attribute!r}): no "
                    f"distinctive value to anchor on in {attr.value_observed!r} -- "
                    f"the legend entry is written but no box is drawn."
                )
                continue

            for evidence_id in attr.evidence_ids:
                citation = citations_by_id.get(evidence_id)
                if citation is None:
                    warnings.append(
                        f"sample {sample_id!r} tickmark {letter}: cites {evidence_id!r}, "
                        f"which is not in evidence_citations -- no box drawn."
                    )
                    continue
                key = (citation.source_file, _page_of(citation.location))
                # Best candidate only. The builder reports an anchor it
                # cannot find, so a second box "just in case" would just be
                # a second box in the wrong place when the first was right.
                by_image.setdefault(key, []).append(
                    EvidenceTieOut(
                        letter=letter,
                        anchor_text=anchors[0],
                        fallback_anchors=anchors[1:],
                    )
                )

        evidence_images: list[EvidenceImage] = []
        for (source_file, page), tie_outs in by_image.items():
            path = support_dir / source_file
            if not path.exists():
                warnings.append(
                    f"sample {sample_id!r}: {source_file!r} is cited but not present in "
                    f"{support_dir} -- that exhibit and its "
                    f"{len(tie_outs)} tickmark(s) are omitted."
                )
                continue
            evidence_images.append(
                EvidenceImage(
                    name=Path(source_file).stem,
                    source_path=str(path),
                    pdf_page=page,
                    # ~6.5in wide: a full page of a letter-size document at a
                    # size a reviewer can actually read on screen.
                    display_width_emu=6_000_000,
                    tie_outs=tie_outs,
                )
            )

        items.append(
            OoxmlSampleItem(
                control_id=control_id,
                raw_data=_raw_data_row(manifest, sample_id) if manifest else RawDataRow(values={}),
                narrative=narrative,
                raw_data_tie_outs=[],
                evidence_images=evidence_images,
                # Absent a per-sample verdict, fall back to the step's
                # roll-up. Never default to True: a workpaper that silently
                # reads "satisfied" is the one failure mode worth being
                # paranoid about.
                conclusion_satisfied=(
                    verdicts.get(sample_id, conclusion.conclusion) == "satisfied"
                ),
            )
        )

    return items, warnings


def build_workpaper_request(
    conclusion: ConclusionOutput,
    manifest: SamplePopulationManifest | None,
    *,
    control_id: str,
    template_path: str | Path,
    output_path: str | Path,
    support_dir: str | Path,
    population_rows: list[dict] | None = None,
) -> tuple[WorkpaperRequest, list[str]]:
    """The whole translation: one test step's conclusion -> a build request."""
    items, warnings = sample_items_from_conclusion(
        conclusion, manifest, control_id, Path(support_dir)
    )
    return (
        WorkpaperRequest(
            template_path=str(template_path),
            output_path=str(output_path),
            sample_items=items,
            population_rows=[
                PopulationRow(values={_LETTERS[i]: v for i, (_k, v) in enumerate(row.items()) if i < 26})
                for row in (population_rows or [])
            ],
        ),
        warnings,
    )
