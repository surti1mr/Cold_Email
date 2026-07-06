"""Create .eml files and open them in Outlook — no Azure, works with New Outlook."""

from __future__ import annotations

import os
import re
import subprocess
import time
from datetime import datetime
from email.message import EmailMessage
from email.policy import SMTP
from pathlib import Path
from typing import Callable

import mimetypes

from outlook_sender import (
    SendResult,
    parse_recipients,
    render_template,
)
from runtime import data_dir

OUTPUT_ROOT = data_dir() / "generated_emails"

CLASSIC_OUTLOOK_PATHS = (
    Path(r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"),
    Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"),
)


def _safe_filename(name: str, email: str) -> str:
    slug = re.sub(r"[^\w\-]+", "_", name.strip())[:40].strip("_") or "recipient"
    domain = email.split("@")[-1].split(".")[0][:20]
    return f"{slug}_{domain}.eml"


def _build_message(
    from_email: str,
    to: str,
    subject: str,
    body: str,
    attachment_path: str | None,
    attachment_filename: str | None = None,
) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = from_email
    msg["To"] = to
    msg["Subject"] = subject
    msg.set_content(body)
    if attachment_path:
        path = Path(attachment_path)
        display_name = attachment_filename or path.name
        mime_type, _ = mimetypes.guess_type(display_name)
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"
        msg.add_attachment(
            path.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            filename=display_name,
        )
    return msg


def open_eml_in_outlook(eml_path: Path) -> str:
    """
    Try to open a compose window for this .eml file.
    Returns how it was opened: classic, shell, or failed.
    """
    eml_path = eml_path.resolve()
    if not eml_path.is_file():
        raise FileNotFoundError(f"Email file not found: {eml_path}")

    for exe in CLASSIC_OUTLOOK_PATHS:
        if exe.is_file():
            subprocess.Popen(
                [str(exe), "/eml", str(eml_path)],
                close_fds=True,
            )
            return "classic"

    # New Outlook / default mail app — may not show a compose window
    os.startfile(str(eml_path))
    return "shell"


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
    open_in_outlook: bool = True,
    on_progress: Callable[[int, int, SendResult], None] | None = None,
) -> tuple[list[SendResult], str]:
    """
    Create one .eml per recipient. Optionally open each in Outlook (you click Send).
    Returns (results, output_folder_path).
    """
    recipients = parse_recipients(recipients_text)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = OUTPUT_ROOT / stamp
    out_dir.mkdir(parents=True, exist_ok=True)

    results: list[SendResult] = []
    total = len(recipients)
    opened_via = "shell"

    for index, recipient in enumerate(recipients, start=1):
        body = render_template(body_template, recipient.name)
        filename = _safe_filename(recipient.name, recipient.email)
        eml_path = out_dir / filename
        try:
            msg = _build_message(
                from_email,
                recipient.email,
                subject,
                body,
                attachment_path,
                attachment_filename,
            )
            eml_path.write_bytes(msg.as_bytes(policy=SMTP))
            if open_in_outlook:
                opened_via = open_eml_in_outlook(eml_path)
                if opened_via == "classic":
                    detail = (
                        f"Compose window opened — click Send in Outlook ({filename})"
                    )
                else:
                    detail = (
                        f"File created — double-click {filename} in the folder, "
                        f"then click Send in Outlook (New Outlook may not auto-open)"
                    )
            else:
                detail = f"Saved {filename} — double-click to open in Outlook"
            results.append(
                SendResult(
                    email=recipient.email,
                    name=recipient.name,
                    success=True,
                    message=detail,
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

        if delay_seconds > 0 and index < total and open_in_outlook:
            time.sleep(delay_seconds)

    # Always open the folder so user can double-click .eml if windows did not appear
    try:
        os.startfile(str(out_dir.resolve()))
    except OSError:
        pass

    return results, str(out_dir)
