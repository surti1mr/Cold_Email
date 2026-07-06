# Cold Email Sender (browser Send — no Azure)

Personalized job emails from **surti1mr@cmich.edu**. Click **Send to all recipients** in the browser and mail goes out automatically.

## Setup (one time)

New Outlook cannot be automated. Use **Office 365 SMTP** with a Microsoft **app password** (not Azure):

See **`SETUP_SMTP.md`** — add to `.env`:

```env
SMTP_EMAIL=surti1mr@cmich.edu
SMTP_PASSWORD=your-16-char-app-password
```

Restart `run.bat`, click **Test SMTP**, then send a test to yourself.

## Quick start

1. `run.bat` → http://127.0.0.1:5000
2. Fill form, attach resume, paste `email,name` recipients
3. **Send to all recipients**

## Fallback

**Prepare .eml files only** — if SMTP is blocked by CMU, creates files you open manually in Outlook.

## Classic Outlook

If you install classic Office Outlook (New Outlook off), the app can send via Outlook without SMTP password.
