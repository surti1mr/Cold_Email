"""Send mail via Office 365 SMTP — browser click sends, no Azure app registration."""

from __future__ import annotations

import os
import smtplib
import ssl
import time
from email.message import EmailMessage
from pathlib import Path
from typing import Callable

import mimetypes

from outlook_sender import (
    SendResult,
    parse_recipients,
    render_template,
)


def _smtp_config() -> dict[str, str]:
    host = (os.getenv("SMTP_HOST") or "smtp.office365.com").strip()
    port = (os.getenv("SMTP_PORT") or "587").strip()
    email = (os.getenv("SMTP_EMAIL") or os.getenv("DEFAULT_FROM") or "").strip()
    password = (os.getenv("SMTP_PASSWORD") or "").strip()
    return {"host": host, "port": port, "email": email, "password": password}


def is_configured() -> bool:
    cfg = _smtp_config()
    return bool(cfg["email"] and cfg["password"])


def test_connection() -> str:
    cfg = _smtp_config()
    if not cfg["email"] or not cfg["password"]:
        raise RuntimeError(
            "SMTP not configured. Add SMTP_EMAIL and SMTP_PASSWORD to your .env file. "
            "See SETUP_SMTP.md."
        )
    with _connect(cfg) as server:
        server.noop()
    return f"Connected as {cfg['email']}"


def _connect(cfg: dict[str, str]):
    port = int(cfg["port"])
    context = ssl.create_default_context()
    server = smtplib.SMTP(cfg["host"], port, timeout=60)
    server.ehlo()
    server.starttls(context=context)
    server.ehlo()
    server.login(cfg["email"], cfg["password"])
    return server


def _build_message(
    from_email: str,
    to: str,
    subject: str,
    body: str,
    attachment_path: str | None,
    attachment_filename: str | None,
) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = from_email
    msg["To"] = to
    msg["Subject"] = subject
    msg.set_content(body)
    if attachment_path:
        path = Path(attachment_path)
        display = attachment_filename or path.name
        mime_type, _ = mimetypes.guess_type(display)
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"
        msg.add_attachment(
            path.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            filename=display,
        )
    return msg


def send_batch(
    *,
    from_email: str,
    subject: str,
    body_template: str,
    recipients_text: str,
    attachment_path: str | None,
    attachment_filename: str | None = None,
    as_draft: bool = False,
    delay_seconds: float = 0,
    on_progress: Callable[[int, int, SendResult], None] | None = None,
) -> list[SendResult]:
    if as_draft:
        raise RuntimeError("SMTP sends immediately; drafts are not supported. Uncheck save as drafts.")

    cfg = _smtp_config()
    if not is_configured():
        raise RuntimeError(
            "SMTP not configured. Add SMTP_PASSWORD to .env — see SETUP_SMTP.md."
        )
    if cfg["email"].lower() != from_email.strip().lower():
        raise RuntimeError(
            f"SMTP_EMAIL in .env is {cfg['email']}, but app sends from {from_email}. "
            "They must match."
        )

    recipients = parse_recipients(recipients_text)
    results: list[SendResult] = []
    total = len(recipients)

    with _connect(cfg) as server:
        for index, recipient in enumerate(recipients, start=1):
            body = render_template(body_template, recipient.name)
            try:
                msg = _build_message(
                    from_email,
                    recipient.email,
                    subject,
                    body,
                    attachment_path,
                    attachment_filename,
                )
                server.send_message(msg)
                results.append(
                    SendResult(
                        email=recipient.email,
                        name=recipient.name,
                        success=True,
                        message="Sent",
                    )
                )
            except Exception as exc:
                results.append(
                    SendResult(
                        email=recipient.email,
                        name=recipient.name,
                        success=False,
                        message=str(exc),
                    )
                )

            if on_progress:
                on_progress(index, total, results[-1])

            if delay_seconds > 0 and index < total:
                time.sleep(delay_seconds)

    return results
