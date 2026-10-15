"""Emails as support: each .msg/.eml becomes a PDF plus its readable
attachments. Every email here is built in the test -- none is copied in.
"""

from __future__ import annotations

from email.message import EmailMessage
from io import BytesIO
from pathlib import Path

from PIL import Image
from reportlab.pdfgen import canvas

from agent.extraction import extract_many
from agent.extraction.email_convert import expand_emails, parse_eml, parse_msg
from agent.tests.msg_fixture import message_tree, write_msg


def _pdf_bytes(text: str) -> bytes:
    buf = BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(72, 720, text)
    c.save()
    return buf.getvalue()


def _png_bytes(size: int) -> bytes:
    buf = BytesIO()
    Image.effect_noise((size, size), 64).convert("RGB").save(buf, "PNG")
    return buf.getvalue()


def _eml(tmp: Path) -> Path:
    inner = EmailMessage()
    inner["From"] = "Vendor AR <ar@vendor.example>"
    inner["Subject"] = "Invoice 9850995"
    inner.set_content("Please find the invoice attached.")

    msg = EmailMessage()
    msg["From"] = "Morgan Chen <mchen@example.com>"
    msg["To"] = "Alex Rivera <arivera@example.com>"
    msg["Cc"] = "AP Team <ap@example.com>"
    msg["Date"] = "Wed, 11 Mar 2026 07:55:00 -0500"
    msg["Subject"] = "RE: JE-2026-0388 for review"
    msg.set_content("Looks good, go ahead.")
    msg.add_alternative("<html><body><p>Looks good, <b>go ahead</b>.</p></body></html>", subtype="html")
    msg.add_attachment(_pdf_bytes("Accrual schedule total 18,250.00"), maintype="application",
                       subtype="pdf", filename="schedule.pdf")
    msg.add_attachment(_png_bytes(8), maintype="image", subtype="png", filename="logo.png")
    msg.add_attachment(_png_bytes(400), maintype="image", subtype="png", filename="screenshot.png")
    msg.add_attachment(b"PK\x03\x04", maintype="application", subtype="octet-stream", filename="memo.docx")
    msg.add_attachment(inner)
    path = tmp / "approval.eml"
    path.write_bytes(bytes(msg))
    return path


def test_eml_headers_body_and_attachments_are_read(tmp_path: Path):
    rec = parse_eml(_eml(tmp_path).read_bytes())
    assert rec.sender == "Morgan Chen <mchen@example.com>"
    assert rec.cc == "AP Team <ap@example.com>"
    assert "07:55:00" in rec.sent
    assert rec.body == "Looks good, go ahead."
    assert [n for n, _ in rec.attachments][:4] == ["schedule.pdf", "logo.png", "screenshot.png", "memo.docx"]


def test_an_email_becomes_a_pdf_and_its_readable_attachments_become_support(tmp_path: Path):
    files, notes = expand_emails([_eml(tmp_path)])
    names = [f.name for f in files]
    assert names[0] == "approval (email).pdf"
    assert "approval (email) - schedule.pdf" in names
    assert "approval (email) - screenshot.pdf" in names  # images become one-page PDFs
    assert not any("logo" in n for n in names)            # too small to be evidence
    assert any(n.startswith("approval (email) - Invoice 9850995") for n in names)  # attached email
    assert any("memo.docx" in n for n in notes)            # unreadable, and said so

    text = " ".join(i.extracted_text or "" for i in extract_many(files))
    for expected in ("From: Morgan Chen", "Sent: Wed, 11 Mar 2026 07:55:00 -0500", "Cc: AP Team",
                     "Subject: RE: JE-2026-0388 for review", "Looks good, go ahead.",
                     "Accrual schedule total 18,250.00", "Please find the invoice attached."):
        assert expected in text, expected


def test_an_html_only_email_is_read_as_text(tmp_path: Path):
    msg = EmailMessage()
    msg["From"] = "a@example.com"
    msg["Subject"] = "Approval"
    msg.set_content("<html><head><style>p{color:red}</style></head><body><p>Approved.</p>"
                    "<table><tr><td>Amount</td><td>$1,200.00</td></tr></table></body></html>", subtype="html")
    rec = parse_eml(bytes(msg))
    assert "Approved." in rec.body and "Amount" in rec.body and "$1,200.00" in rec.body
    assert "color" not in rec.body


def test_outlook_msg_is_read_with_its_attachments_and_attached_emails(tmp_path: Path):
    inner = message_tree(subject="Invoice 9850995", sender="Vendor AR", body="Invoice attached.",
                         embedded=True, attachments=[("invoice.pdf", _pdf_bytes("Invoice total 14,945.60"))])
    tree = message_tree(
        subject="RE: Approval needed", sender="Morgan Chen", sender_email="mchen@example.com",
        to="Alex Rivera", cc="AP Team", body="Approved.\n\nMorgan",
        sent_filetime=134_000_000_000_000_000,  # a fixed instant in 2025
        attachments=[("check_request.pdf", _pdf_bytes("Check request 14,945.60")), ("Fwd invoice.msg", inner)],
    )
    path = tmp_path / "approval.msg"
    path.write_bytes(write_msg(tree))

    rec = parse_msg(path.read_bytes())
    assert (rec.subject, rec.sender, rec.to, rec.cc) == (
        "RE: Approval needed", "Morgan Chen <mchen@example.com>", "Alex Rivera", "AP Team")
    assert rec.sent.endswith("2025 14:13 UTC")
    assert rec.body == "Approved.\n\nMorgan"

    files, notes = expand_emails([path])
    assert notes == []
    text = " ".join(i.extracted_text or "" for i in extract_many(files))
    for expected in ("Subject: RE: Approval needed", "Approved.", "Check request 14,945.60",
                     "Subject: Invoice 9850995", "Invoice total 14,945.60"):
        assert expected in text, expected


def test_outlook_sent_time_prefers_the_messages_own_date_header(tmp_path: Path):
    tree = message_tree(subject="s", sender="x", body="b", sent_filetime=134_000_000_000_000_000)
    tree["__substg1.0_007D001F"] = "Received: by x\r\nDate: Tue, 3 Feb 2026 09:15:00 -0600\r\n".encode("utf-16-le")
    assert parse_msg(write_msg(tree)).sent == "Tue, 3 Feb 2026 09:15:00 -0600"


def test_non_email_files_pass_through_and_a_broken_email_is_reported(tmp_path: Path):
    pdf = tmp_path / "invoice.pdf"
    pdf.write_bytes(_pdf_bytes("x"))
    bad = tmp_path / "broken.msg"
    bad.write_bytes(b"not an outlook file")
    files, notes = expand_emails([pdf, bad])
    assert files == [pdf]
    assert any("broken.msg could not be read" in n for n in notes)
