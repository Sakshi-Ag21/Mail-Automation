from __future__ import annotations

import mimetypes
import smtplib
from dataclasses import dataclass
from email.message import EmailMessage

from .config import SmtpConfig


@dataclass(frozen=True)
class Attachment:
    filename: str
    content_type: str | None
    data: bytes


def render_template(text: str, *, name: str, email: str) -> str:
    """
    Basic personalization using Python format placeholders.

    Supported placeholders:
    - {name}
    - {email}
    """
    safe = {
        "name": name,
        "email": email,
    }
    try:
        return text.format(**safe)
    except Exception:
        # If the user types unmatched braces or an invalid placeholder,
        # fall back to raw text rather than failing the whole send.
        return text


def build_message(
    *,
    smtp: SmtpConfig,
    to_email: str,
    subject: str,
    body: str,
    attachments: list[Attachment] | None = None,
    from_email: str | None = None,
    from_name: str | None = None,
) -> EmailMessage:
    msg = EmailMessage()
    effective_email = (from_email or "").strip() or smtp.email
    effective_name = (from_name or smtp.sender_name or "").strip()
    from_display = effective_email if not effective_name else f"{effective_name} <{effective_email}>"
    msg["From"] = from_display
    msg["To"] = to_email
    msg["Subject"] = subject

    # Plain text for maximum compatibility (Gmail SMTP).
    msg.set_content(body)

    for att in attachments or []:
        ctype = att.content_type
        if not ctype:
            ctype, _ = mimetypes.guess_type(att.filename)
        if not ctype:
            ctype = "application/octet-stream"

        maintype, subtype = ctype.split("/", 1)
        msg.add_attachment(
            att.data,
            maintype=maintype,
            subtype=subtype,
            filename=att.filename,
        )

    return msg


def send_via_gmail_smtp(*, smtp: SmtpConfig, message: EmailMessage) -> None:
    with smtplib.SMTP(smtp.host, smtp.port, timeout=30) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(smtp.email, smtp.app_password)
        server.send_message(message)

