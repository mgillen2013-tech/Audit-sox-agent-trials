"""Vision-OCR tests against a fake client -- no network, no API key.

The real behavior being locked in: image-only pages get transcribed and
marked as OCR-derived (not passed off as natively-extracted text), and
every failure path degrades to the original placeholder rather than
breaking a run.
"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from agent.extraction import extract, ocr_image_items
from agent.extraction.ocr import OCR_CONFIDENCE
from agent.schemas import EvidenceItem


class _OCRClient:
    """Returns canned transcriptions; records what it was sent."""

    def __init__(self, texts: list[str] | None = None, raises: bool = False):
        self._texts = texts if texts is not None else ["INVOICE 2859\nPremiere Onboard LLC | $30,000.00"]
        self._i = 0
        self.raises = raises
        self.calls: list[dict] = []

    def create_message(self, **kwargs):
        self.calls.append(kwargs)
        if self.raises:
            raise RuntimeError("vision endpoint unavailable")
        text = self._texts[min(self._i, len(self._texts) - 1)]
        self._i += 1
        return SimpleNamespace(content=[SimpleNamespace(type="text", text=text)])


@pytest.fixture
def scanned_pdf(tmp_path: Path) -> Path:
    """A page with only a rasterized image -- no text layer, so extraction
    produces an image_ocr placeholder.
    """
    img = tmp_path / "scan.png"
    Image.new("RGB", (600, 400), color="white").save(img)
    path = tmp_path / "invoice_2859.pdf"
    c = canvas.Canvas(str(path), pagesize=letter)
    c.drawImage(str(img), 50, 400, width=500, height=300)
    c.showPage()
    c.save()
    return path


def test_scanned_page_gets_transcribed_and_marked(scanned_pdf: Path, tmp_path: Path):
    items = extract(scanned_pdf)
    assert [i.source_type for i in items] == ["image_ocr"]
    assert items[0].extracted_text is None

    client = _OCRClient()
    out, count = ocr_image_items(items, tmp_path, client, "claude-opus-5")

    assert count == 1
    assert "INVOICE 2859" in out[0].extracted_text
    # Marked as a transcription, not passed off as a native read.
    assert out[0].extracted_text.startswith("[OCR transcription")
    assert out[0].extraction_confidence == OCR_CONFIDENCE
    # Identity is preserved -- workpaper.py re-extracts and maps by these.
    assert out[0].evidence_id == items[0].evidence_id
    assert out[0].location == items[0].location


def test_image_is_actually_sent_to_the_model(scanned_pdf: Path, tmp_path: Path):
    client = _OCRClient()
    ocr_image_items(extract(scanned_pdf), tmp_path, client, "claude-opus-5")

    content = client.calls[0]["messages"][0]["content"]
    image_blocks = [b for b in content if b["type"] == "image"]
    assert len(image_blocks) == 1
    assert image_blocks[0]["source"]["media_type"] == "image/png"
    assert image_blocks[0]["source"]["data"]  # non-empty base64

    prompt = " ".join(b["text"] for b in content if b["type"] == "text")
    assert "verbatim" in prompt
    # The anti-fabrication instruction is the whole point on audit evidence.
    assert "Never infer" in prompt
    assert "[illegible]" in prompt


def test_ocr_failure_degrades_to_placeholder(scanned_pdf: Path, tmp_path: Path):
    items = extract(scanned_pdf)
    out, count = ocr_image_items(items, tmp_path, _OCRClient(raises=True), "claude-opus-5")

    assert count == 0
    assert out[0].extracted_text is None
    assert out[0].extraction_confidence == 0.0


def test_empty_transcription_is_not_recorded_as_evidence(scanned_pdf: Path, tmp_path: Path):
    out, count = ocr_image_items(extract(scanned_pdf), tmp_path, _OCRClient(texts=[""]), "claude-opus-5")
    assert count == 0
    assert out[0].extracted_text is None


def test_non_image_items_pass_through_untouched(tmp_path: Path):
    # Native-text pages must never be re-read by a model: they're already
    # exact, and a transcription would be strictly worse evidence.
    native = EvidenceItem(
        evidence_id="ev_0001",
        source_file="recon.xlsx",
        source_type="excel_table",
        location="Sheet1!A1:B2",
        extracted_text="Accrual 482110",
        extraction_confidence=1.0,
        preview_ref="recon.xlsx!Sheet1!A1:B2",
    )
    client = _OCRClient()
    out, count = ocr_image_items([native], tmp_path, client, "claude-opus-5")

    assert count == 0
    assert client.calls == []
    assert out[0] == native


def test_progress_callback_reports_each_page(scanned_pdf: Path, tmp_path: Path):
    seen: list[tuple] = []
    ocr_image_items(
        extract(scanned_pdf),
        tmp_path,
        _OCRClient(),
        "claude-opus-5",
        on_page=lambda f, p, ok: seen.append((f, p, ok)),
    )
    assert seen == [("invoice_2859.pdf", 1, True)]
