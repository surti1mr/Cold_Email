"""Send emails via Outlook web (outlook.office.com) using Playwright browser automation.
No Azure, no SMTP password, works with New Outlook and MFA.
"""

from __future__ import annotations

import json
import re
import shutil
import tempfile
import time
from pathlib import Path
from typing import Callable

from outlook_sender import SendResult, parse_recipients, render_template
from runtime import data_dir, is_serverless

SESSION_FILE = data_dir() / "owa_session.json"


def _require_local_browser(action: str) -> None:
    if is_serverless():
        raise RuntimeError(
            f"{action} is not available on Vercel. Run the app locally to sign in and send email."
        )


def _sync_playwright():
    from playwright.sync_api import sync_playwright

    return sync_playwright
OWA_URL = "https://outlook.cloud.microsoft/mail/"
COMPOSE_TIMEOUT = 30_000
SEND_TIMEOUT = 15_000


# ---------------------------------------------------------------------------
# Session helpers
# ---------------------------------------------------------------------------

def _save_session(context) -> None:
    cookies = context.cookies()
    storage = context.storage_state()
    SESSION_FILE.write_text(json.dumps(storage), encoding="utf-8")


def _session_exists() -> bool:
    return SESSION_FILE.exists() and SESSION_FILE.stat().st_size > 100


def is_logged_in() -> bool:
    """Quick non-interactive check — try to load OWA without a login page."""
    if is_serverless() or not _session_exists():
        return False
    try:
        sync_playwright = _sync_playwright()
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            ctx = browser.new_context(storage_state=json.loads(SESSION_FILE.read_text()))
            page = ctx.new_page()
            page.goto(OWA_URL, wait_until="domcontentloaded", timeout=20_000)
            time.sleep(2)
            url = page.url.lower()
            browser.close()
            return "login" not in url and "microsoftonline" not in url
    except Exception:
        return False


def do_browser_login() -> None:
    """Open a visible browser, let the user sign in with MFA, then save the session."""
    _require_local_browser("Outlook Web sign-in")
    sync_playwright = _sync_playwright()
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, slow_mo=50)
        ctx = browser.new_context()
        page = ctx.new_page()
        page.goto(OWA_URL)
        print("Waiting for you to sign in to Outlook Web …")
        # Wait until the inbox/mail page loads (URL no longer contains login domains)
        page.wait_for_url(
            lambda url: (
                "outlook.office.com" in url
                and "login" not in url
                and "microsoftonline" not in url
            ),
            timeout=300_000,  # 5 minutes to log in
        )
        # Extra wait for page to settle
        time.sleep(3)
        _save_session(ctx)
        browser.close()
        print("Session saved.")


def logout() -> None:
    if SESSION_FILE.exists():
        SESSION_FILE.unlink()


def _fill_to_field(page, to: str) -> None:
    """Fill the To field and auto-confirm OWA's 'Use this address' prompt."""
    to_field = page.locator("div[aria-label='To'][contenteditable='true']").first
    to_field.wait_for(state="visible", timeout=COMPOSE_TIMEOUT)
    to_field.click()
    time.sleep(0.2)
    page.keyboard.press("Control+a")
    page.keyboard.press("Delete")
    time.sleep(0.2)

    # Paste is more reliable than character-by-character typing
    page.evaluate("(text) => navigator.clipboard.writeText(text)", to)
    page.keyboard.press("Control+v")
    time.sleep(1.0)

    # OWA pauses until you pick a suggestion or confirm an unknown address
    resolved = False
    for attempt in [
        lambda: page.get_by_text("Use this address", exact=False).first.click(timeout=2_000),
        lambda: page.get_by_text("Use this email address", exact=False).first.click(timeout=2_000),
        lambda: page.locator("[aria-label*='Use this address']").first.click(timeout=2_000),
        lambda: page.get_by_role("option", name=re.compile(re.escape(to), re.I)).first.click(timeout=2_000),
        lambda: page.get_by_role("listbox").get_by_text(to, exact=False).first.click(timeout=2_000),
    ]:
        try:
            attempt()
            resolved = True
            break
        except Exception:
            continue

    if not resolved:
        page.keyboard.press("Enter")
        time.sleep(0.4)

    page.keyboard.press("Tab")
    time.sleep(0.5)


def _wait_for_attachment(page, filename: str) -> None:
    """Wait until OWA shows the attachment chip (not 'not found')."""
    stem = Path(filename).stem
    for _ in range(30):  # up to ~15s
        if page.get_by_text("couldn't be found", exact=False).count() > 0:
            raise RuntimeError(f"Attachment failed: {filename} not found by Outlook")
        if page.get_by_text("not found", exact=False).count() > 0:
            # Only treat as error if near attachment area
            err = page.locator("[class*='attachment'], [data-automationid*='attachment']").filter(
                has_text=re.compile("not found", re.I)
            )
            if err.count() > 0:
                raise RuntimeError(f"Attachment failed: {filename} not found by Outlook")
        # Attachment chip visible — name or partial name
        if page.get_by_text(filename, exact=False).count() > 0:
            return
        if stem and page.get_by_text(stem, exact=False).count() > 0:
            return
        time.sleep(0.5)
    time.sleep(1.0)  # extra buffer even if chip text not matched


def _wait_for_compose_closed(page) -> None:
    """Wait until the current compose window is gone before starting the next email."""
    time.sleep(1.5)
    try:
        page.locator("div[aria-label='Message body'][contenteditable='true']").wait_for(
            state="hidden", timeout=12_000
        )
    except Exception:
        pass
    time.sleep(0.5)


# ---------------------------------------------------------------------------
# Single email
# ---------------------------------------------------------------------------

def _send_one(
    page,
    to: str,
    subject: str,
    body: str,
    attachment_path: str | None,
    attachment_filename: str | None,
    scheduled_at: str | None = None,
) -> None:
    """Compose and send (or schedule) one email in an already-open OWA page.

    scheduled_at: ISO datetime string like '2026-05-28T09:30' from datetime-local input.
                  If None, email is sent immediately.
    """

    # 1. Click the "New" button (exact label in current Outlook Cloud UI)
    page.get_by_role("button", name="New", exact=True).wait_for(
        state="visible", timeout=COMPOSE_TIMEOUT
    )
    page.get_by_role("button", name="New", exact=True).click()
    time.sleep(1.5)

    # 2. To field — paste email and auto-confirm "Use this address" if shown
    _fill_to_field(page, to)

    # 3. Subject — input[aria-label='Subject']
    subject_field = page.locator("input[aria-label='Subject']").first
    subject_field.wait_for(state="visible", timeout=10_000)
    subject_field.click()
    subject_field.fill(subject)
    time.sleep(0.3)

    # 4. Body — paste via clipboard to preserve spacing exactly
    body_area = page.locator(
        "div[aria-label='Message body'][contenteditable='true']"
    ).first
    body_area.wait_for(state="visible", timeout=10_000)
    body_area.click()
    page.keyboard.press("Control+a")
    page.keyboard.press("Delete")
    time.sleep(0.2)
    # Write body to clipboard and paste — preserves newlines without extra spacing
    page.evaluate("(text) => navigator.clipboard.writeText(text)", body)
    page.keyboard.press("Control+v")
    time.sleep(0.5)

    # 5. Attach file — attachment_path must stay on disk until upload completes
    if attachment_path:
        attach_file = Path(attachment_path).resolve()
        if not attach_file.is_file():
            raise RuntimeError(f"Attachment missing on disk: {attach_file}")
        display_name = attachment_filename or attach_file.name

        attach_btn = page.locator("button[aria-label='Attach file']").first
        attach_btn.wait_for(state="visible", timeout=8_000)
        attach_btn.click()
        time.sleep(1)
        try:
            browse_btn = page.get_by_role("menuitem", name="Browse this computer")
            browse_btn.wait_for(state="visible", timeout=5_000)
            with page.expect_file_chooser(timeout=8_000) as fc_info:
                browse_btn.click()
            fc_info.value.set_files(str(attach_file))
        except Exception:
            file_inputs = page.locator("input[type='file']")
            if file_inputs.count() > 0:
                file_inputs.first.set_input_files(str(attach_file))
            else:
                raise
        _wait_for_attachment(page, display_name)

    # 6. Send now OR schedule
    if scheduled_at:
        _schedule_send(page, scheduled_at)
    else:
        send_btn = page.locator("button[aria-label='Send']").first
        send_btn.wait_for(state="visible", timeout=SEND_TIMEOUT)
        send_btn.click()

    _wait_for_compose_closed(page)


def _schedule_send(page, scheduled_at: str) -> None:
    """Use OWA's Schedule Send feature to queue the email for a specific time.

    scheduled_at format: 'YYYY-MM-DDTHH:MM' (from HTML datetime-local input).
    """
    import re as _re
    from datetime import datetime

    dt = datetime.fromisoformat(scheduled_at)
    date_str = dt.strftime("%m/%d/%Y")
    time_str = dt.strftime("%I:%M %p").lstrip("0")  # e.g. "9:30 AM"

    # ── Step 1: click the dropdown arrow next to the Send button ──────────────
    dropdown_selectors = [
        "button[aria-label='Send options']",
        "button[aria-label*='More send options']",
        "button[aria-label*='Send options']",
        "[data-automationid='splitButtonSecondary']",
    ]
    clicked = False
    for sel in dropdown_selectors:
        try:
            btn = page.locator(sel).first
            btn.wait_for(state="visible", timeout=3_000)
            btn.click()
            clicked = True
            break
        except Exception:
            continue

    if not clicked:
        raise RuntimeError(
            "Schedule Send not available: could not find the 'More send options' "
            "dropdown next to the Send button in OWA."
        )
    time.sleep(1)

    # ── Step 2: click "Schedule send" in the dropdown menu ────────────────────
    scheduled_clicked = False
    for attempt in [
        lambda: page.get_by_role("menuitem", name="Schedule send").click(),
        lambda: page.get_by_role("menuitem", name=_re.compile("schedule", _re.I)).first.click(),
        lambda: page.get_by_text("Schedule send", exact=False).first.click(),
    ]:
        try:
            attempt()
            scheduled_clicked = True
            break
        except Exception:
            continue

    if not scheduled_clicked:
        raise RuntimeError("Could not find 'Schedule send' option in the send menu.")
    time.sleep(1.5)

    # ── Step 3 & 4: fill date and time inside the dialog ──────────────────────
    # OWA's schedule dialog shows preset options first ("Tonight", "Tomorrow"…).
    # We must click "Custom time" to reveal the date/time inputs.
    dialog = page.locator("[role='dialog']").first
    try:
        dialog.wait_for(state="visible", timeout=5_000)
    except Exception:
        dialog = page  # fallback: search whole page

    # Click "Custom time" to reveal the date/time picker.
    # Search the full page (not just dialog) for reliability.
    custom_time_clicked = False
    for attempt in [
        lambda: page.get_by_role("button",   name="Custom time").click(),
        lambda: page.get_by_role("menuitem", name="Custom time").click(),
        lambda: page.get_by_role("option",   name="Custom time").click(),
        lambda: page.get_by_text("Custom time", exact=True).click(),
        lambda: page.get_by_text("Custom time", exact=False).first.click(),
    ]:
        try:
            attempt()
            custom_time_clicked = True
            break
        except Exception:
            continue
    time.sleep(1.2)  # wait for date/time inputs to animate in

    def _fill_field(locator, value: str) -> bool:
        """Click, select-all, type. Returns True on success."""
        try:
            locator.wait_for(state="visible", timeout=3_000)
            locator.click()
            locator.press("Control+a")
            locator.press("Delete")
            locator.type(value, delay=50)
            return True
        except Exception:
            return False

    # Date — try aria-label first, then fall back to first input in dialog
    date_filled = False
    for sel in ["input[aria-label='Date']", "input[aria-label*='date' i]",
                "input[placeholder*='date' i]", "input[type='date']"]:
        if _fill_field(dialog.locator(sel).first, date_str):
            date_filled = True
            break
    if not date_filled:
        # Last resort: first input inside dialog
        _fill_field(dialog.locator("input").nth(0), date_str)

    page.keyboard.press("Tab")
    time.sleep(0.3)

    # Time — try aria-label first, then fall back to second input in dialog
    time_filled = False
    for sel in ["input[aria-label='Time']", "input[aria-label*='time' i]",
                "input[placeholder*='time' i]", "input[type='time']"]:
        if _fill_field(dialog.locator(sel).first, time_str):
            time_filled = True
            break
    if not time_filled:
        _fill_field(dialog.locator("input").nth(1), time_str)

    page.keyboard.press("Tab")
    time.sleep(0.5)

    # ── Step 5: confirm / schedule ─────────────────────────────────────────────
    for attempt in [
        lambda: dialog.get_by_role("button", name="Send").click(),
        lambda: dialog.get_by_role("button", name="Schedule").click(),
        lambda: dialog.get_by_role("button", name=_re.compile(r"send|schedule", _re.I)).last.click(),
        lambda: page.keyboard.press("Enter"),
    ]:
        try:
            attempt()
            break
        except Exception:
            continue

    time.sleep(2)


# ---------------------------------------------------------------------------
# Batch send
# ---------------------------------------------------------------------------

def send_batch(
    *,
    from_email: str,
    subject: str,
    body_template: str,
    recipients_text: str,
    attachment_path: str | None,
    attachment_filename: str | None = None,
    as_draft: bool = False,
    delay_seconds: float = 5,
    scheduled_at: str | None = None,
    on_progress: Callable[[int, int, SendResult], None] | None = None,
) -> list[SendResult]:
    _require_local_browser("Outlook Web sending")
    if not _session_exists():
        raise RuntimeError(
            "Not signed in to Outlook Web. Click 'Sign in to Outlook' on the page first."
        )

    recipients = parse_recipients(recipients_text)
    results: list[SendResult] = []
    total = len(recipients)

    # One temp copy for the whole batch — OWA reads the file async; deleting
    # after the first email caused "attachment not found" on email #2+.
    attach_tmp_dir: Path | None = None
    batch_attachment_path: str | None = None
    if attachment_path:
        src = Path(attachment_path).resolve()
        if not src.is_file():
            raise RuntimeError(f"Attachment not found: {src}")
        clean_name = attachment_filename or src.name
        attach_tmp_dir = Path(tempfile.mkdtemp(prefix="cold_email_attach_"))
        batch_file = attach_tmp_dir / clean_name
        shutil.copy2(src, batch_file)
        batch_attachment_path = str(batch_file)

    sync_playwright = _sync_playwright()
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, slow_mo=30)
        ctx = browser.new_context(
            storage_state=json.loads(SESSION_FILE.read_text(encoding="utf-8")),
            permissions=["clipboard-read", "clipboard-write"],
        )
        page = ctx.new_page()

        try:
            page.goto(OWA_URL, wait_until="domcontentloaded", timeout=30_000)
            time.sleep(4)

            # Check we landed on mail, not login
            if "login" in page.url.lower() or ("microsoftonline" in page.url.lower() and "oauth2" in page.url.lower()):
                browser.close()
                SESSION_FILE.unlink(missing_ok=True)
                raise RuntimeError(
                    "Outlook session expired. Click 'Sign in to Outlook' on the page to sign in again."
                )

            for index, recipient in enumerate(recipients, start=1):
                body = render_template(body_template, recipient.name)
                try:
                    _send_one(
                        page,
                        to=recipient.email,
                        subject=subject,
                        body=body,
                        attachment_path=batch_attachment_path,
                        attachment_filename=attachment_filename,
                        scheduled_at=scheduled_at,
                    )
                    msg = (
                        f"Scheduled for {scheduled_at} via Outlook Web"
                        if scheduled_at
                        else "Sent via Outlook Web"
                    )
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

            _save_session(ctx)
        finally:
            browser.close()
            if attach_tmp_dir:
                shutil.rmtree(attach_tmp_dir, ignore_errors=True)

    return results
