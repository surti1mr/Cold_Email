"""Send mail via Microsoft Graph (works with New Outlook / M365). Free — no paid API."""

from __future__ import annotations

import base64
import time
from pathlib import Path
from typing import Callable

import msal
import requests

from outlook_sender import (
    Recipient,
    SendResult,
    parse_recipients,
    render_template,
)

BASE_DIR = Path(__file__).resolve().parent
TOKEN_CACHE = BASE_DIR / "token_cache.json"
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Mail.Send", "User.Read"]
REDIRECT_URI = "http://127.0.0.1:5000/auth/callback"


def _client_id() -> str:
    import os

    cid = (os.getenv("AZURE_CLIENT_ID") or "").strip()
    if not cid:
        raise RuntimeError(
            "Microsoft Graph is not configured. Copy config.example.env to .env "
            "and add your free Azure Application (client) ID. See SETUP_GRAPH.md."
        )
    return cid


def _msal_app(cache: msal.SerializableTokenCache | None = None):
    cache = cache or _load_cache()
    return msal.PublicClientApplication(
        _client_id(),
        authority=AUTHORITY,
        token_cache=cache,
    )


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE.exists():
        cache.deserialize(TOKEN_CACHE.read_text(encoding="utf-8"))
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        TOKEN_CACHE.write_text(cache.serialize(), encoding="utf-8")


def is_configured() -> bool:
    import os

    return bool((os.getenv("AZURE_CLIENT_ID") or "").strip())


def is_authenticated() -> bool:
    if not is_configured():
        return False
    return get_access_token() is not None


def get_signed_in_email() -> str | None:
    token = get_access_token()
    if not token:
        return None
    r = requests.get(
        f"{GRAPH_ROOT}/me",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    if r.ok:
        return r.json().get("mail") or r.json().get("userPrincipalName")
    return None


def get_access_token() -> str | None:
    if not is_configured():
        return None
    cache = _load_cache()
    app = _msal_app(cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]
    return None


def get_auth_url() -> str:
    app = _msal_app()
    return app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )


def complete_auth(code: str) -> str:
    cache = _load_cache()
    app = _msal_app(cache)
    result = app.acquire_token_by_authorization_code(
        code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    if "error" in result:
        raise RuntimeError(result.get("error_description") or result["error"])
    _save_cache(cache)
    email = get_signed_in_email()
    return email or "Signed in"


def logout() -> None:
    if TOKEN_CACHE.exists():
        TOKEN_CACHE.unlink()


def _attachment_payload(path: str) -> dict:
    data = Path(path).read_bytes()
    name = Path(path).name
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": name,
        "contentBytes": base64.b64encode(data).decode("ascii"),
    }


def _create_draft(token: str, to: str, subject: str, body: str, attachment_path: str | None) -> None:
    message: dict = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": [{"emailAddress": {"address": to}}],
    }
    if attachment_path:
        message["attachments"] = [_attachment_payload(attachment_path)]

    r = requests.post(
        f"{GRAPH_ROOT}/me/mailFolders/drafts/messages",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=message,
        timeout=60,
    )
    if not r.ok:
        err = r.json() if r.headers.get("content-type", "").startswith("application/json") else {}
        raise RuntimeError(err.get("error", {}).get("message", r.text))


def _send_now(token: str, to: str, subject: str, body: str, attachment_path: str | None) -> None:
    message: dict = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": [{"emailAddress": {"address": to}}],
    }
    if attachment_path:
        message["attachments"] = [_attachment_payload(attachment_path)]

    r = requests.post(
        f"{GRAPH_ROOT}/me/sendMail",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={"message": message, "saveToSentItems": True},
        timeout=60,
    )
    if not r.ok:
        err = r.json() if r.headers.get("content-type", "").startswith("application/json") else {}
        raise RuntimeError(err.get("error", {}).get("message", r.text))


def send_batch(
    *,
    from_email: str,
    subject: str,
    body_template: str,
    recipients_text: str,
    attachment_path: str | None,
    as_draft: bool,
    delay_seconds: float,
    on_progress: Callable[[int, int, SendResult], None] | None = None,
) -> list[SendResult]:
    token = get_access_token()
    if not token:
        raise RuntimeError(
            "Not signed in to Microsoft. Click “Sign in with Microsoft” on the page first."
        )

    signed_in = (get_signed_in_email() or "").lower()
    if signed_in and signed_in != from_email.strip().lower():
        raise RuntimeError(
            f"Signed in as {signed_in}, but this app sends from {from_email}. "
            "Sign out and sign in with your CMU account."
        )

    recipients = parse_recipients(recipients_text)
    results: list[SendResult] = []
    total = len(recipients)

    for index, recipient in enumerate(recipients, start=1):
        body = render_template(body_template, recipient.name)
        try:
            if as_draft:
                _create_draft(token, recipient.email, subject, body, attachment_path)
                msg = "Saved to Outlook Drafts (via Microsoft 365)"
            else:
                _send_now(token, recipient.email, subject, body, attachment_path)
                msg = "Sent via Microsoft 365"
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
