"""
Structured objects the "critical thinking" agent must produce.

Nothing in this file touches XML, EMU, OCR, or zip files. An agent (Opus/Sonnet)
should be prompted to fill these in from (a) the raw sample extract row and
(b) the evidence PDFs/images, and the result should be validated against these
models before being handed to build_workpaper.build_workpaper().

Design note: RunOption for narrative color intentionally only allows the two
colors actually used in every workpaper observed (control-red for
labels/conclusions, black for descriptive body text). Extend the Literal if a
new template introduces a third color.
"""
from __future__ import annotations
from typing import Literal, Optional
from pydantic import BaseModel, Field, model_validator


# ---------------------------------------------------------------------------
# Raw data (the JDE / Data Access Studio extract row for the sample)
# ---------------------------------------------------------------------------

class RawDataRow(BaseModel):
    """One row of the F0413/F0414/F0411-style extract. Keys are the column
    letters used in the template (A..T) -> value. Populate every column the
    template header row defines, even if blank (use None)."""
    values: dict[str, object] = Field(
        description="Column letter ('A'..'T') -> cell value. Dates as "
                    "datetime.date/datetime, numbers as int/float, text as str, "
                    "blank as None."
    )


class PopulationRow(BaseModel):
    """One row of the Population tab. Same column-letter convention as
    RawDataRow but without any tie-out/callout metadata — Population is never
    annotated with shapes."""
    values: dict[str, object]


# ---------------------------------------------------------------------------
# Tie-out / callout metadata
# ---------------------------------------------------------------------------

class RawDataTieOut(BaseModel):
    """A callout box drawn over one or more cells in the sample's raw-data row."""
    letter: str = Field(description="Single capital letter, e.g. 'A'")
    columns: list[str] = Field(
        description="Column letters this callout spans, e.g. ['C','D'] to box "
                    "a payee-number column next to a payee-name column."
    )


class EvidenceTieOut(BaseModel):
    """A callout box drawn over a value found on an evidence image via OCR."""
    letter: str = Field(description="Single capital letter, matches the legend "
                                     "used in the narrative textbox, e.g. 'A'")
    anchor_text: str = Field(
        description="Exact text to search for in the OCR word boxes on this "
                    "image, e.g. '21022452' or '30,000.00'. Must be copyable "
                    "verbatim from the source document — do not paraphrase."
    )
    fallback_anchors: list[str] = Field(
        default_factory=list,
        description="Other verbatim tokens identifying the SAME value, tried "
                    "in order when anchor_text is not found. The same fact is "
                    "written differently on different documents: an approval "
                    "email says 'February 9, 2026' where the payment screen "
                    "says '02/09/2026', so a callout anchored on the numeric "
                    "date is unfindable on the email even though the invoice "
                    "number beside it would have matched. Same rules as "
                    "anchor_text: verbatim, no spaces."
    )
    occurrence: int = Field(
        default=0,
        description="If anchor_text appears more than once on the image, which "
                    "match to use (0-indexed, in reading order top-to-bottom)."
    )
    extra_words: int = Field(
        default=0,
        description="How many additional OCR word-boxes immediately following "
                    "the anchor to include in the same box (e.g. anchor_text="
                    "'Diane' extra_words=3 to box 'Diane Milosevic' across two "
                    "lines/runs). Usually 0."
    )


class EvidenceImage(BaseModel):
    """One piece of evidence (a rendered PDF page or an already-cropped image)
    to embed on the Sample tab, with its callouts."""
    name: str = Field(description="Human label, e.g. 'Invoice', 'Approval Email', "
                                   "'Payment Screenshot'")
    source_path: str = Field(description="Path to the source file: a .pdf "
                                          "(will be rendered) or an image file.")
    pdf_page: int = Field(default=0, description="If source is a PDF, which "
                                                   "page (0-indexed) to render.")
    display_width_emu: int = Field(
        description="Target on-sheet width in EMU (914400 EMU = 1 inch). "
                    "Height is derived preserving aspect ratio. Pick something "
                    "in the 6,000,000-9,000,000 range (~6.5-9.8in) for a "
                    "full-page document; smaller for a cropped screenshot."
    )
    tie_outs: list[EvidenceTieOut] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# Narrative
# ---------------------------------------------------------------------------

class NarrativeRun(BaseModel):
    text: str
    bold: bool = False
    color: Literal["RED", "BLACK"] = "BLACK"


class NarrativeParagraph(BaseModel):
    """A single paragraph in the test-step narrative textbox. Empty runs list
    = a blank line (paragraph spacer)."""
    runs: list[NarrativeRun] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# IA Calculation (optional supporting tie-out sum, plain cells not a shape)
# ---------------------------------------------------------------------------

class IACalculationLine(BaseModel):
    label: str
    amount: float


class IACalculation(BaseModel):
    lines: list[IACalculationLine]
    total_label: str = "Total"

    @property
    def total(self) -> float:
        return round(sum(l.amount for l in self.lines), 2)


# ---------------------------------------------------------------------------
# Top-level: one tested sample
# ---------------------------------------------------------------------------

class SampleItem(BaseModel):
    control_id: str = Field(description="e.g. 'SS.PTP.AP.156'")
    sample_id: str = Field(
        default="",
        description="The selection's id from the sample listing. This names "
                    "the item's own tab, and the Summary tab links to it by "
                    "the same name -- derive both from this one field or the "
                    "links silently point at tabs that do not exist.",
    )
    raw_data: RawDataRow
    narrative: list[NarrativeParagraph]
    raw_data_tie_outs: list[RawDataTieOut] = Field(default_factory=list)
    evidence_images: list[EvidenceImage] = Field(default_factory=list)
    ia_calculation: Optional[IACalculation] = None
    conclusion_satisfied: bool = Field(
        description="True = 'Test step satisfied. No exceptions noted.' style "
                    "conclusion. False = an exception was found (the narrative "
                    "must describe it; do not silently mark satisfied)."
    )

    @model_validator(mode="after")
    def _letters_are_consistent(self):
        """Guardrail: every letter used in a callout must also appear in the
        narrative legend, and vice versa, or the workpaper is misleading."""
        narrative_text = " ".join(
            r.text for p in self.narrative for r in p.runs
        )
        used_letters = {t.letter for t in self.raw_data_tie_outs}
        for img in self.evidence_images:
            used_letters |= {t.letter for t in img.tie_outs}
        for letter in used_letters:
            if f"{letter} - " not in narrative_text and f"{letter} –" not in narrative_text:
                raise ValueError(
                    f"Callout letter '{letter}' is used but not defined in the "
                    f"narrative legend (expected a line like '{letter} - <label>')."
                )
        return self

    @model_validator(mode="after")
    def _ia_calc_ties_to_raw_data(self):
        """Guardrail: if an IA Calculation is present, its total should equal
        the payment/invoice amount actually recorded in the raw data row, or
        this workpaper is internally inconsistent."""
        if self.ia_calculation is None:
            return self
        # Try common amount columns; skip check if none present rather than
        # guessing which column is "the" amount for an unfamiliar template.
        for col in ("G", "K"):
            v = self.raw_data.values.get(col)
            if isinstance(v, (int, float)):
                if round(float(v), 2) != self.ia_calculation.total:
                    raise ValueError(
                        f"IA Calculation total ({self.ia_calculation.total}) does "
                        f"not tie to raw data column {col} ({v})."
                    )
                return self
        return self


# ---------------------------------------------------------------------------
# Parameters tab
# ---------------------------------------------------------------------------

class ParametersCallout(BaseModel):
    anchor_text: str
    occurrence: int = 0
    extra_words: int = 0


class ParametersTab(BaseModel):
    screenshot_path: str
    screenshot_dpi: float = Field(
        description="DPI the screenshot was captured/rendered at. Check the "
                    "image's own metadata (PIL Image.info['dpi']) rather than "
                    "assuming 96 - paste tools vary."
    )
    display_width_emu: Optional[int] = Field(
        default=None,
        description="If None, use the image's native size at its stated DPI "
                    "(i.e. paste at 100%, matching audit convention)."
    )
    narrative: list[NarrativeParagraph]
    callouts: list[ParametersCallout] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# Whole-workpaper request
# ---------------------------------------------------------------------------

class WorkpaperRequest(BaseModel):
    template_path: str
    output_path: str
    sample_items: list[SampleItem]
    population_rows: list[PopulationRow]
    parameters: Optional[ParametersTab] = None
