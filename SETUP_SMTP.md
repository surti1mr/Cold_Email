# Send from the browser (no Azure) — Office 365 SMTP

This lets you click **Send to all recipients** in the app and emails go out automatically from `surti1mr@cmich.edu`.

You need a **Microsoft app password** (not Azure). CMU may or may not allow this on student accounts.

## 1. Create an app password

1. Sign in at https://mysignins.microsoft.com/security-info with **surti1mr@cmich.edu**
2. If you see **App passwords** → create one named `Cold Email` → copy the 16-character password
3. If you do **not** see App passwords:
   - CMU may have disabled them. Try https://account.microsoft.com/security (personal) only if you use a personal SMTP mailbox instead
   - Or install **classic Office Outlook** and turn off New Outlook (then the app uses Outlook automation, no password file)

## 2. Configure `.env`

In the `Cold_Email` folder, create or edit `.env`:

```env
SMTP_HOST=smtp.office365.com
SMTP_PORT=587
SMTP_EMAIL=surti1mr@cmich.edu
SMTP_PASSWORD=paste-the-16-char-app-password-here
```

**No spaces** in the app password. Do not commit `.env` to git.

## 3. Restart the app

1. Stop the terminal (Ctrl+C)
2. Run `run.bat`
3. The page should say **SMTP configured — Send sends mail automatically**
4. Click **Test SMTP** on the page (optional)
5. Send a test to your own Gmail first

## Troubleshooting

| Error | Fix |
|-------|-----|
| `535 Authentication unsuccessful` | Wrong password, or CMU blocked SMTP — use classic Outlook instead |
| `SmtpClientAuthentication is disabled` | CMU disabled SMTP for students — use classic Outlook |
| `SMTP_EMAIL must match` | Set `SMTP_EMAIL=surti1mr@cmich.edu` in `.env` |
