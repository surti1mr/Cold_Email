# Free Microsoft Graph setup (for New Outlook)

New Outlook (`olk.exe`) cannot be automated from Python. This **free** one-time setup uses Microsoft’s official mail API instead — still sends from `surti1mr@cmich.edu`.

## 1. Register an app (free)

1. Open https://portal.azure.com and sign in with your **CMU** account.
2. Search **App registrations** → **New registration**.
3. Name: `Cold Email Local` (any name).
4. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**.
5. Redirect URI: **Web** → `http://127.0.0.1:5000/auth/callback`
6. Click **Register**.
7. Copy the **Application (client) ID** (looks like `a1b2c3d4-...`).

## 2. Enable public client + permissions

1. In your app → **Authentication** → **Advanced settings** → enable **Allow public client flows** → **Yes** → Save.
2. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**:
   - `Mail.Send`
   - `User.Read`
3. **Grant admin consent** — click if available. If CMU blocks it, try **Consent on behalf of your organization** or sign in anyway; you may get a personal consent prompt when signing in.

## 3. Configure this project

```powershell
cd Cold_Email
copy config.example.env .env
```

Edit `.env` and set:

```env
AZURE_CLIENT_ID=paste-your-client-id-here
```

## 4. Run and sign in

1. `run.bat` or `python app.py`
2. Open http://127.0.0.1:5000
3. Click **Sign in with Microsoft** → use **surti1mr@cmich.edu**
4. Send emails as usual (drafts still work — they appear in Outlook Drafts online)

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Azure won’t let students register apps | Ask CMU IT, or use a **personal** Microsoft account to register the app (you can still sign in with CMU when sending). |
| Admin consent required | IT must approve, or register app under a personal Azure account. |
| Wrong account signed in | Click **Sign out**, sign in again with CMU. |
| Still want COM automation | Install **classic** Microsoft 365 Office Outlook, turn **off** “New Outlook” toggle, use `olk.exe` only after classic is installed. |
