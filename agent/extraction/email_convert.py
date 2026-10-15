"""Emails (.msg from Outlook, .eml) -> PDFs the rest of the pipeline reads.

An approval is very often an email, and its evidence is often an
attachment to it. Rather than teach every later step (extraction, OCR,
exhibits, tickmarks, the mirror engine) a new file type, each email is
turned into a PDF once, at upload: From / To / Cc / Sent / Subject and the
body, exactly as a person would print it. Its readable attachments become
support files of their own, named after the email they came with, and an
email attached to an email is converted the same way.

.eml is read with the standard library. .msg is read with olefile -- an
Outlook file is an OLE container of named property streams -- because the
usual .msg package needs a compiler to install on some machines.
"""

from __future__ import annotations

import email
import email.policy
import re
import struct
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from html.parser import HTMLParser
from io import BytesIO
from pathlib import Path

EMAIL_SUFFIXES = (".msg", ".eml")
_READABLE = (".pdf", ".xlsx", ".xlsm", ".xls")
_IMAGES = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff")
# Smaller than this an image cannot carry readable evidence -- it is a
# signature logo or a spacer, and would only clutter the support.
_MIN_IMAGE_BYTES = 5_000


@dataclass
class EmailRecord:
    subject: str = ""
    sender: str = ""
    to: str = ""
    cc: str = ""
    sent: str = ""
    body: str = ""
    # (file name, bytes) or (file name, EmailRecord) for an attached email
    attachments: list[tuple[str, "bytes | EmailRecord"]] = field(default_factory=list)


# ── .eml ─────────────────────────────────────────────────────────────────


class _HtmlText(HTMLParser):
    _BREAKS = {"p", "br", "div", "tr", "li", "h1", "h2", "h3", "h4", "h5", "h6", "table", "blockquote"}

    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):  # noqa: ANN001
        if tag in ("style", "script", "head"):
            self._skip += 1
        elif tag in self._BREAKS:
            self.parts.append("\n")
        elif tag == "td":
            self.parts.append("  ")

    def handle_endtag(self, tag):  # noqa: ANN001
        if tag in ("style", "script", "head"):
            self._skip = max(0, self._skip - 1)
        elif tag in self._BREAKS:
            self.parts.append("\n")

    def handle_data(self, data):  # noqa: ANN001
        if not self._skip:
            self.parts.append(data)


def html_to_text(html: str) -> str:
    parser = _HtmlText()
    parser.feed(html)
    text = "".join(parser.parts).replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n\s*\n\s*\n+", "\n\n", text).strip()


def parse_eml(data: bytes) -> EmailRecord:
    msg = email.message_from_bytes(data, policy=email.policy.default)
    return _from_message(msg)


def _from_message(msg) -> EmailRecord:  # noqa: ANN001
    rec = EmailRecord(
        subject=str(msg.get("Subject", "") or ""),
        sender=str(msg.get("From", "") or ""),
        to=str(msg.get("To", "") or ""),
        cc=str(msg.get("Cc", "") or ""),
        sent=str(msg.get("Date", "") or ""),
    )
    body = msg.get_body(preferencelist=("plain", "html"))
    if body is not None:
        content = body.get_content()
        rec.body = html_to_text(content) if body.get_content_type() == "text/html" else content.strip()
    for part in msg.iter_attachments():
        name = part.get_filename() or ""
        if part.get_content_type() == "message/rfc822":
            inner = part.get_content()
            rec.attachments.append((name or f"{inner.get('Subject', 'attached email')}.eml", _from_message(inner)))
        else:
            payload = part.get_payload(decode=True)
            if payload:
                rec.attachments.append((name or "attachment", payload))
    return rec


# ── .msg (Outlook) ───────────────────────────────────────────────────────


def _filetime(value: int) -> str:
    if not value:
        return ""
    when = datetime(1601, 1, 1, tzinfo=timezone.utc) + timedelta(microseconds=value / 10)
    return when.strftime("%a, %d %b %Y %H:%M UTC")


def parse_msg(data: bytes) -> EmailRecord:
    import olefile

    ole = olefile.OleFileIO(BytesIO(data))
    try:
        return _msg_record(ole, "", 32)
    finally:
        ole.close()


def _msg_record(ole, prefix: str, header_len: int) -> EmailRecord:  # noqa: ANN001
    def prop(pid: str) -> str | bytes | None:
        for kind in ("001F", "001E", "0102"):
            path = f"{prefix}__substg1.0_{pid}{kind}"
            if ole.exists(path):
                raw = ole.openstream(path).read()
                if kind == "001F":
                    return raw.decode("utf-16-le", errors="replace").rstrip("\x00")
                if kind == "001E":
                    return raw.decode("cp1252", errors="replace").rstrip("\x00")
                return raw
        return None

    def text(pid: str) -> str:
        value = prop(pid)
        return value.strip() if isinstance(value, str) else ""

    sender = text("0C1A")
    address = text("0C1F") or text("5D01")
    if address and "@" in address and address not in sender:
        sender = f"{sender} <{address}>" if sender else address

    headers = text("007D")
    sent = ""
    m = re.search(r"^Date:\s*(.+)$", headers, re.M)
    if m:
        sent = m.group(1).strip()
    else:
        sent = _msg_time(ole, prefix, header_len)

    body = text("1000")
    if not body:
        html = prop("1013")
        if isinstance(html, bytes):
            html = html.decode("utf-8", errors="replace")
        if html:
            body = html_to_text(html)

    rec = EmailRecord(subject=text("0037"), sender=sender, to=text("0E04"), cc=text("0E03"),
                      sent=sent, body=body)

    seen = set()
    for entry in ole.listdir(streams=True, storages=True):
        path = "/".join(entry)
        if not path.startswith(prefix):
            continue
        rest = path[len(prefix):].split("/")
        if not rest[0].startswith("__attach_version1.0_#") or rest[0] in seen:
            continue
        seen.add(rest[0])
        att = f"{prefix}{rest[0]}/"

        def att_text(pid: str, _att=att) -> str:
            for kind in ("001F", "001E"):
                p = f"{_att}__substg1.0_{pid}{kind}"
                if ole.exists(p):
                    raw = ole.openstream(p).read()
                    return raw.decode("utf-16-le" if kind == "001F" else "cp1252", errors="replace").rstrip("\x00").strip()
            return ""

        name = att_text("3707") or att_text("3704") or att_text("3001") or "attachment"
        if ole.exists(f"{att}__substg1.0_3701000D"):
            inner = _msg_record(ole, f"{att}__substg1.0_3701000D/", 24)
            rec.attachments.append((name if name != "attachment" else f"{inner.subject or 'attached email'}.msg", inner))
        elif ole.exists(f"{att}__substg1.0_37010102"):
            rec.attachments.append((name, ole.openstream(f"{att}__substg1.0_37010102").read()))
    return rec


def _msg_time(ole, prefix: str, header_len: int) -> str:  # noqa: ANN001
    """Sent (or delivered) time from the message's fixed-size properties."""
    path = f"{prefix}__properties_version1.0"
    if not ole.exists(path):
        return ""
    raw = ole.openstream(path).read()
    found: dict[int, int] = {}
    for off in range(header_len, len(raw) - 15, 16):
        ptype, pid = struct.unpack_from("<HH", raw, off)
        if ptype == 0x0040 and pid in (0x0039, 0x0E06):
            found[pid] = struct.unpack_from("<Q", raw, off + 8)[0]
    return _filetime(found.get(0x0039) or found.get(0x0E06) or 0)


# ── to PDF ───────────────────────────────────────────────────────────────


def _safe(text: str) -> str:
    """Escaped for reportlab, limited to what its standard font can draw."""
    text = text.encode("cp1252", errors="replace").decode("cp1252")
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def email_to_pdf(rec: EmailRecord, out_path: Path, notes: list[str] | None = None) -> None:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    styles = getSampleStyleSheet()
    body_style = styles["BodyText"]
    rows = [("From", rec.sender), ("Sent", rec.sent), ("To", rec.to), ("Cc", rec.cc), ("Subject", rec.subject)]
    if rec.attachments:
        rows.append(("Attachments", "; ".join(name for name, _ in rec.attachments)))
    table = Table(
        [[Paragraph(f"<b>{k}:</b>", body_style), Paragraph(_safe(v), body_style)] for k, v in rows if v],
        colWidths=[1.1 * inch, 6.0 * inch],
    )
    table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                               ("LINEBELOW", (0, -1), (-1, -1), 0.75, colors.grey)]))
    story = [table, Spacer(1, 0.2 * inch)]
    for para in re.split(r"\n\s*\n", rec.body or "(no message body)"):
        story.append(Paragraph(_safe(para.strip()).replace("\n", "<br/>"), body_style))
    for note in notes or []:
        story.append(Paragraph(f"<i>{_safe(note)}</i>", body_style))
    SimpleDocTemplate(str(out_path), pagesize=letter, leftMargin=0.7 * inch, rightMargin=0.7 * inch,
                      topMargin=0.7 * inch, bottomMargin=0.7 * inch).build(story)


def _unique(path: Path) -> Path:
    n = 2
    candidate = path
    while candidate.exists():
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        n += 1
    return candidate


def _clean(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|\r\n]+', "_", name).strip() or "attachment"


def _write_email(rec: EmailRecord, stem: str, folder: Path, out: list[Path], notes: list[str]) -> None:
    skipped: list[str] = []
    pending: list[tuple[str, "bytes | EmailRecord"]] = []
    for name, content in rec.attachments:
        suffix = Path(name).suffix.lower()
        if isinstance(content, EmailRecord) or suffix in _READABLE + EMAIL_SUFFIXES:
            pending.append((name, content))
        elif suffix in _IMAGES:
            if len(content) >= _MIN_IMAGE_BYTES:
                pending.append((name, content))
        else:
            skipped.append(name)
    own_notes = [f"Attachment not readable by the agent (save it as PDF and upload it): {n}" for n in skipped]
    pdf = _unique(folder / f"{stem}.pdf")
    email_to_pdf(rec, pdf, own_notes)
    out.append(pdf)
    notes.extend(f"{stem}: {n}" for n in own_notes)

    for name, content in pending:
        base = f"{stem} - {_clean(Path(name).stem)}"
        suffix = Path(name).suffix.lower()
        if isinstance(content, EmailRecord):
            _write_email(content, base, folder, out, notes)
        elif suffix in EMAIL_SUFFIXES:
            try:
                inner = parse_msg(content) if suffix == ".msg" else parse_eml(content)
            except Exception as exc:  # noqa: BLE001
                notes.append(f"{stem}: attached email {name} could not be read ({exc})")
                continue
            _write_email(inner, base, folder, out, notes)
        elif suffix in _IMAGES:
            from PIL import Image

            target = _unique(folder / f"{base}.pdf")
            try:
                Image.open(BytesIO(content)).convert("RGB").save(target, "PDF", resolution=150)
                out.append(target)
            except Exception as exc:  # noqa: BLE001
                notes.append(f"{stem}: image {name} could not be read ({exc})")
        else:
            target = _unique(folder / f"{base}{suffix}")
            target.write_bytes(content)
            out.append(target)


def expand_emails(paths: list[Path]) -> tuple[list[Path], list[str]]:
    """Support files with every email replaced by its PDF and its readable
    attachments, in upload order. Non-email files pass through unchanged.

    Returns (files, notes): notes say which attachments could not be read,
    so a reviewer knows evidence was attached that the agent did not see.
    """
    out: list[Path] = []
    notes: list[str] = []
    for path in paths:
        path = Path(path)
        suffix = path.suffix.lower()
        if suffix not in EMAIL_SUFFIXES:
            out.append(path)
            continue
        try:
            rec = parse_msg(path.read_bytes()) if suffix == ".msg" else parse_eml(path.read_bytes())
        except Exception as exc:  # noqa: BLE001
            notes.append(f"{path.name} could not be read as an email ({exc})")
            continue
        _write_email(rec, f"{path.stem} (email)", path.parent, out, notes)
    return out, notes
