# 📧 Outlook Bulk Mailer

A modern Windows desktop application for sending **personalised bulk emails** via Microsoft Outlook using the **Microsoft Graph API** and secure OAuth2 sign-in.

Built with Python · PySide6 · MSAL · Microsoft Graph API

---

## Features

| Feature | Description |
|---------|-------------|
| **Excel Import** | Import `.xlsx` files with `Email` and `Name` columns; duplicate/format validation |
| **Email Composer** | Rich subject + HTML body; personalise with `{{Name}}`; add file attachments |
| **Secure Login** | Microsoft OAuth Device Code flow — password is **never** stored or requested |
| **Bulk Send** | Separate personalised email per recipient; saved to Sent Items |
| **Progress View** | Real-time progress bar, per-email status, pause / stop / resume |
| **Connection Test** | Verify auth → Graph API → Outlook pipeline before sending |
| **History** | SQLite-backed campaign log; export to CSV or Excel |
| **Modern UI** | Dark / light theme, sidebar nav, stat cards |

---

## Prerequisites

- **Python 3.11+** (Windows recommended; works on Linux/macOS for development)
- **Microsoft 365 or personal Microsoft account** (e.g. `@outlook.com`, `@hotmail.com`, work/school)
- **A free Azure App Registration** (one-time, ~5 minutes) — instructions below

---

## Step 1 — Register Your App in Microsoft Entra (Azure AD)

> **Why?** Microsoft Graph API requires an "app identity" so it knows which application is requesting access. This is a free, one-time setup. No credit card needed.

### 1.1 — Go to the Entra Admin Centre

1. Open your browser and go to **[https://entra.microsoft.com](https://entra.microsoft.com)**
2. Sign in with your **Microsoft account** (the same one you'll use to send emails).
3. In the left sidebar expand **Applications** → click **App registrations**.

### 1.2 — Create a New Registration

1. Click **+ New registration** (top-left of the page).
2. Fill in the form:

   | Field | Value |
   |-------|-------|
   | **Name** | `OutlookBulkMailer` (or any name you prefer) |
   | **Supported account types** | **"Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"** |
   | **Redirect URI** | Leave blank for now |

3. Click **Register**.

> **Personal accounts only?** If you only use `@outlook.com` / `@hotmail.com`, the "Multitenant + personal" option above works perfectly. If you have a work/school account only, you can choose "Accounts in this organizational directory only" instead.

### 1.3 — Copy Your Application (Client) ID

After registering you land on the **Overview** page.

1. Locate the **Application (client) ID** field — it looks like:
   ```
   12345678-abcd-1234-efgh-123456789abc
   ```
2. Click the copy icon next to it.
3. **Keep this value** — you'll paste it into the app in Step 3.

### 1.4 — Configure the Authentication Platform

1. In the left panel click **Authentication**.
2. Click **+ Add a platform**.
3. Select **Mobile and desktop applications**.
4. Under **Custom redirect URIs** tick the checkbox for:
   ```
   https://login.microsoftonline.com/common/oauth2/nativeclient
   ```
5. Scroll down and set **"Allow public client flows"** to **Yes**.
6. Click **Save**.

> **Why "public client"?** The app uses the **Device Code Flow** (RFC 8628). This is a public-client flow — it never uses a client secret, which means your credentials are always secure.

### 1.5 — Add Microsoft Graph Permissions

1. In the left panel click **API permissions**.
2. Click **+ Add a permission**.
3. Select **Microsoft Graph**.
4. Select **Delegated permissions**.
5. In the search box type `Mail.Send` → tick **Mail.Send**.
6. Clear the box, type `User.Read` → tick **User.Read**.
7. Click **Add permissions**.

Your permissions panel should now show:

| Permission | Type | Status |
|-----------|------|--------|
| Mail.Send | Delegated | ✅ |
| User.Read | Delegated | ✅ |

> **Grant admin consent**: If you see a yellow warning about admin consent and your account is a personal Microsoft account, you can ignore it — personal accounts grant consent interactively during sign-in. Work/school accounts managed by an IT department may require an admin to click "Grant admin consent for [tenant]".

---

## Step 2 — Install and Run the App

```bash
# 1. Clone or download the project
cd OutlookBulkMailer

# 2. Create a virtual environment
python -m venv venv

# 3. Activate it
venv\Scripts\activate          # Windows
# source venv/bin/activate     # macOS / Linux

# 4. Install dependencies
pip install -r requirements.txt

# 5. Launch
python app.py
```

---

## Step 3 — Configure the App

1. In the sidebar click **⚙ Settings**.
2. In the **Azure App Registration** card paste your **Application (client) ID** from Step 1.3.
3. Click **Save** — the field turns green with ✅ if the UUID format is valid.

---

## Step 4 — Sign In with Microsoft

1. Still in Settings, click **🔐 Sign in with Microsoft**.
2. A browser tab opens automatically to `https://microsoft.com/devicelogin`.
3. Enter the **one-time code** shown in the app (click 📋 Copy Code if needed).
4. Sign in with your Microsoft account.
5. When prompted, click **Accept** to grant `Mail.Send` and `User.Read`.
6. The app updates automatically — you'll see ✅ **Signed in as [Your Name]**.

---

## Step 5 — Verify the Connection (Recommended)

1. In Settings → **Connection Status** card click **🔄 Verify Now**.
2. The indicator turns 🟢 green showing your display name and email.
3. Optionally click **📧 Send Test Email** to send a test message to yourself.

---

## Step 6 — Send Your Campaign

1. **Compose** tab → enter subject, body (use `{{Name}}` for personalisation), add attachments.
2. **Import Excel** → select a `.xlsx` file with `Email` and `Name` columns.
3. Click **Preview & Send** → review the recipient list and campaign summary.
4. Click **Start Sending** → watch the real-time progress.
5. **History** tab → view results; export to CSV or Excel.

---

## Excel File Format

| Email | Name |
|-------|------|
| alice@example.com | Alice |
| bob@example.com | Bob |

- Column headers must be exactly `Email` and `Name` (case-insensitive).
- Duplicate email addresses are automatically detected and flagged.

---

## In-App Setup Guide

You can also view the complete step-by-step Microsoft Entra setup guide directly inside the application:
1. Open the app and click **⚙ Settings** in the sidebar.
2. In the **Microsoft Entra App Registration Guide** card, click **▶ Show Guide**.
3. Follow the 6 easy steps with direct links to the Entra portal and Microsoft Graph documentation.

---

## 🎨 Application Icon & Customization

The app includes a modern, high-resolution custom icon. The icon files are located in the `assets/` folder:

```
OutlookBulkMailer/
└── assets/
    ├── icon.png      # 1024x1024 high-res PNG (window title bar, taskbar, About screen)
    └── icon.ico      # Multi-resolution Windows ICO: 16, 24, 32, 48, 64, 128, 256 px
```

### How to Use Your Own Custom Icon:

1. **Replace the PNG**:
   Place your logo or image at `assets/icon.png` (a square PNG of at least 256x256 or 1024x1024 is recommended).

2. **Generate the Windows `.ico`**:
   Run this one-line command to generate the multi-resolution `.ico` with Pillow:
   ```bash
   python -c "from PIL import Image; Image.open('assets/icon.png').save('assets/icon.ico', format='ICO', sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])"
   ```

3. **How It Works in Code**:
   - **Window & Taskbar (`app.py`)**:
     ```python
     # Sets Windows AppUserModelID so taskbar displays the custom icon:
     ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("vikrantyadav.outlookbulkmailer.app.1.0")

     # Loads icon with sys._MEIPASS support for both source and PyInstaller:
     _base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
     app.setWindowIcon(QIcon(os.path.join(_base, "assets", "icon.png")))
     ```
   - **Standalone Executable (`build.bat`)**:
     PyInstaller packages the `.ico` directly into the `.exe` binary:
     ```bat
     pyinstaller --onefile --windowed --icon "assets\icon.ico" --add-data "assets;assets" app.py
     ```

---

## Building a Windows `.exe`

To build a standalone Windows executable (`OutlookBulkMailer.exe`):

```bat
REM Run from project root on Windows:
build.bat
```

The resulting `OutlookBulkMailer.exe` is placed in `dist/`. It contains the embedded icon, all modules, and runs without requiring Python to be installed on the target machine.


---

## Troubleshooting

### "Azure App Client ID is not configured"
→ Go to **Settings**, paste your Client ID, click **Save**.

### "AADSTS700016: Application not found"
→ Your Client ID is wrong or was copied incompletely. It must be a UUID like `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`.

### "AADSTS50194: Application is not configured as a multi-tenant application"
→ In the Entra portal go to your app → **Authentication** → change **Supported account types** to "Multitenant + personal Microsoft accounts".

### "Permission denied (HTTP 403)"
→ The `Mail.Send` permission may not be added. Re-check Step 1.5. If you're on a work account, ask your IT admin to grant consent.

### "Access token expired (HTTP 401)"
→ Go to **Settings** and sign in again. Tokens last ~1 hour; the app refreshes them automatically when possible.

### "Sending throttled (HTTP 429)"
→ Microsoft limits sending rate (~30/min for most accounts). Increase the inter-email delay in the Send page and try again.

### Sign-in code expired
→ The device code is valid for ~15 minutes. Click **Sign in with Microsoft** to get a fresh code.

### Consent was declined
→ Sign in again and click **Accept** when Microsoft asks for permissions.

---

## Security

- ✅ OAuth2 Device Code Flow — no password is ever entered in or stored by this app.
- ✅ Access tokens are only used in `Authorization` headers — never logged or written to disk.
- ✅ Client secrets are not used (public-client application).
- ✅ Token cache is stored locally in `data/token_cache.bin` (MSAL encrypted format).
- ✅ Client ID is a public identifier (not a secret) — safe to store in `data/settings.json`.

---

## References

- [App registration (Microsoft Entra / Azure AD)](https://learn.microsoft.com/en-us/entra/identity-platform/scenario-desktop-app-registration)
- [Microsoft Graph sendMail API](https://learn.microsoft.com/en-us/graph/api/user-sendmail)
- [MSAL Device Code Flow (Python)](https://learn.microsoft.com/en-us/azure/active-directory/develop/scenario-desktop-acquire-token-device-code-flow)
- [Graph API error codes](https://learn.microsoft.com/en-us/graph/errors)

---

## Developer

**Vikrant Yadav** · [@engvikrant21](https://github.com/engvikrant21)

---

## License

MIT License — © 2026 Vikrant Yadav
