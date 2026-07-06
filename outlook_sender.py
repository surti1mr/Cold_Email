"""Send personalized emails through Outlook desktop (Windows + pywin32)."""

from __future__ import annotations

import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

OL_MAIL_ITEM = 0
OL_FOLDER_DRAFTS = 16


@dataclass
class Recipient:
    email: str
    name: str


@dataclass
class SendResult:
    email: str
    name: str
    success: bool
    message: str


def parse_recipients(text: str) -> list[Recipient]:
    """Parse lines: email,name or email<TAB>name. Skips blanks and # comments."""
    rows: list[Recipient] = []
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "," in line:
            parts = [p.strip() for p in line.split(",", 1)]
        elif "\t" in line:
            parts = [p.strip() for p in line.split("\t", 1)]
        else:
            raise ValueError(
                f'Invalid line (need email,name): "{line}"'
            )
        if len(parts) != 2 or not parts[0] or not parts[1]:
            raise ValueError(f'Invalid line: "{line}"')
        rows.append(Recipient(email=parts[0], name=parts[1]))
    if not rows:
        raise ValueError("Add at least one recipient (email,name per line).")
    return rows


def render_template(template: str, name: str) -> str:
    """Replace {{name}} and {name} placeholders."""
    return (
        template.replace("{{name}}", name)
        .replace("{{NAME}}", name)
        .replace("{name}", name)
    )


def _get_outlook():
    import os
    import win32com.client
    from pywintypes import com_error

    errors: list[str] = []
    for factory in (
        lambda: win32com.client.GetActiveObject("Outlook.Application"),
        lambda: win32com.client.Dispatch("Outlook.Application"),
        lambda: win32com.client.gencache.EnsureDispatch("Outlook.Application"),
    ):
        try:
            return factory()
        except com_error as exc:
            errors.append(str(exc))

    new_outlook = os.path.expandvars(
        r"%LOCALAPPDATA%\Microsoft\WindowsApps\olk.exe"
    )
    hint = "Cannot connect to classic Outlook."
    if os.path.isfile(new_outlook):
        hint = (
            "New Outlook does not support automation — the app will create .eml files "
            "and open them in Outlook instead (no Azure)."
        )
    raise RuntimeError(hint) from None


def list_accounts() -> list[dict[str, str]]:
    outlook = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    accounts = []
    for i in range(1, namespace.Accounts.Count + 1):
        acc = namespace.Accounts.Item(i)
        smtp = getattr(acc, "SmtpAddress", "") or ""
        display = getattr(acc, "DisplayName", "") or smtp
        accounts.append({"smtp": smtp, "display": display})
    return accounts


def _find_account(namespace, from_email: str):
    target = from_email.strip().lower()
    for i in range(1, namespace.Accounts.Count + 1):
        acc = namespace.Accounts.Item(i)
        smtp = (getattr(acc, "SmtpAddress", "") or "").lower()
        if smtp == target:
            return acc
    available = [
        getattr(namespace.Accounts.Item(i), "SmtpAddress", "")
        for i in range(1, namespace.Accounts.Count + 1)
    ]
    raise ValueError(
        f'Account "{from_email}" not found in Outlook. Available: {", ".join(available)}'
    )


def _create_mail(
    outlook,
    account,
    to: str,
    subject: str,
    body: str,
    attachment_path: str | None,
    attachment_filename: str | None = None,
):
    mail = outlook.CreateItem(OL_MAIL_ITEM)
    mail.SendUsingAccount = account
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    if attachment_path:
        display = attachment_filename or Path(attachment_path).name
        # olByValue=1 — attach file; last arg sets the name shown in the email
        att = mail.Attachments.Add(attachment_path, 1, 1, display)
        att.DisplayName = display
    return mail


def preview_first(
    subject: str,
    body_template: str,
    recipients_text: str,
    from_email: str,
) -> dict:
    recipients = parse_recipients(recipients_text)
    r = recipients[0]
    return {
        "from_email": from_email,
        "to": r.email,
        "name": r.name,
        "subject": subject,
        "body": render_template(body_template, r.name),
        "total": len(recipients),
    }


def send_batch(
    *,
    from_email: str,
    subject: str,
    body_template: str,
    recipients_text: str,
    attachment_path: str | None,
    attachment_filename: str | None = None,
    as_draft: bool,
    delay_seconds: float,
    on_progress: Callable[[int, int, SendResult], None] | None = None,
) -> list[SendResult]:
    recipients = parse_recipients(recipients_text)
    outlook = _get_outlook()
    namespace = outlook.GetNamespace("MAPI")
    account = _find_account(namespace, from_email)

    results: list[SendResult] = []
    total = len(recipients)

    for index, recipient in enumerate(recipients, start=1):
        body = render_template(body_template, recipient.name)
        try:
            mail = _create_mail(
                outlook,
                account,
                recipient.email,
                subject,
                body,
                attachment_path,
                attachment_filename,
            )
            if as_draft:
                drafts = namespace.GetDefaultFolder(OL_FOLDER_DRAFTS)
                mail.Save()
                mail.Move(drafts)
                msg = "Saved to Drafts"
            else:
                mail.Send()
                msg = "Sent"
            results.append(
                SendResult(
                    email=recipient.email,
                    name=recipient.name,
                    success=True,
                    message=msg,
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
