"""Send: classic Outlook COM → Playwright (OWA) → SMTP → .eml fallback."""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

import eml_sender
import outlook_sender
import smtp_sender
from runtime import is_serverless


def detect_outlook_kind() -> str:
    classic = Path(r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE")
    classic_x86 = Path(
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
    )
    new_stub = Path(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\WindowsApps\olk.exe"))
    if classic.exists() or classic_x86.exists():
        return "classic"
    if new_stub.exists():
        return "new"
    return "unknown"


def com_available() -> bool:
    try:
        outlook_sender._get_outlook()
        return True
    except Exception:
        return False


def get_status() -> dict:
    smtp_ok = smtp_sender.is_configured()
    if is_serverless():
        if smtp_ok:
            return {
                "outlook_kind": "cloud",
                "com_available": False,
                "smtp_configured": True,
                "owa_logged_in": False,
                "backend": "smtp",
                "can_auto_send": True,
                "can_send": True,
                "message": (
                    "Hosted on Vercel with SMTP configured — sending is enabled. "
                    "Outlook Web sign-in only works on your local machine."
                ),
            }
        return {
            "outlook_kind": "cloud",
            "com_available": False,
            "smtp_configured": False,
            "owa_logged_in": False,
            "backend": "none",
            "can_auto_send": False,
            "can_send": True,
            "message": (
                "Hosted on Vercel — generate and preview emails here. "
                "To send, run the app locally or add SMTP env vars in Vercel."
            ),
        }

    import playwright_sender

    kind = detect_outlook_kind()
    com_ok = com_available()
    owa_ok = playwright_sender.is_logged_in()

    if com_ok:
        backend = "com"
        can_auto_send = True
        message = "Classic Outlook ready — browser Send works."
    elif owa_ok:
        backend = "owa"
        can_auto_send = True
        message = "Signed in to Outlook Web — browser Send works."
    elif smtp_ok:
        backend = "smtp"
        can_auto_send = True
        message = (
            "SMTP ready — emails send automatically from "
            f"{smtp_sender._smtp_config()['email']}."
        )
    else:
        backend = "none"
        can_auto_send = False
        message = (
            "Sign in to Outlook Web first (button below) to enable one-click Send."
        )

    return {
        "outlook_kind": kind,
        "com_available": com_ok,
        "smtp_configured": smtp_ok,
        "owa_logged_in": owa_ok,
        "backend": backend,
        "can_auto_send": can_auto_send,
        "can_send": True,
        "message": message,
    }


def send_batch(*, mode: str = "auto", open_in_outlook: bool = True, scheduled_at: str | None = None, **kwargs):
    """mode: auto | owa | smtp | com | eml.
    scheduled_at: ISO datetime string ('YYYY-MM-DDTHH:MM') — only supported by owa backend.
    """
    if mode == "eml":
        return eml_sender.send_batch(open_in_outlook=open_in_outlook, **kwargs), None

    if mode == "com" or (mode == "auto" and com_available()):
        results = outlook_sender.send_batch(**kwargs)
        return results, None

    if not is_serverless():
        import playwright_sender

        if mode == "owa" or (mode == "auto" and playwright_sender.is_logged_in()):
            results = playwright_sender.send_batch(scheduled_at=scheduled_at, **kwargs)
            return results, None

    if mode == "smtp" or (mode == "auto" and smtp_sender.is_configured()):
        results = smtp_sender.send_batch(**kwargs)
        return results, None

    raise RuntimeError(get_status()["message"])
