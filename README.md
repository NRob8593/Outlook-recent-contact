# Recent contact — Outlook add-in

Checks the "To" field while you're composing/replying and tells you whether
you've already emailed that person 1:1 in the last 7 days, by searching
your Sent Items via Microsoft Graph.

## What's in this folder

- `manifest.xml` — the add-in manifest (classic XML format, works with your
  Win32 Outlook + Exchange).
- `taskpane.html` / `taskpane.css` / `taskpane.js` — the panel itself.
- `commands.html` — lets the pane auto-open on a new compose window.
- `config.js` — placeholders for your Azure AD app details.
- `assets/` — placeholder icons (swap for your own branding if you like).

## Setup (three things need doing before this works)

### 1. Register an app in Azure AD
This add-in needs its own app registration to call Microsoft Graph on your
behalf.
1. Go to [entra.microsoft.com](https://entra.microsoft.com) → **App registrations** → **New registration**.
2. Name it (e.g. "Circuit32 Recent Contact"), leave account type as
   single-tenant, add a **Web** redirect URI once you know your hosting URL
   (step 2) — e.g. `https://your-domain/taskpane.html`.
3. Under **API permissions**, add Microsoft Graph → Delegated →
   `Mail.Read`, then grant admin consent if your tenant requires it.
4. Copy the **Application (client) ID** and **Directory (tenant) ID**.
5. Paste them into `config.js`, along with your redirect URI.

### 2. Host the files somewhere with HTTPS
Outlook add-ins must be served over HTTPS — it won't load from your local
disk. Easiest options:
- **Azure Static Web Apps** (free tier, straightforward for a single user)
- **GitHub Pages** (free, if you don't mind the code being public — or use
  a private repo with Pages via GitHub Pro)
- Any existing web host you already have

Once hosted, replace every `YOUR-HOSTING-DOMAIN` placeholder in
`manifest.xml` and `config.js` with your real domain.

### 3. Sideload the manifest into Outlook
In Outlook (classic desktop):
**File → Manage Add-ins** (or **Home → Get Add-ins → My add-ins → Add a
custom add-in → Add from file**) → select your `manifest.xml`.

If "Get Add-ins" isn't visible, your admin may need to enable custom add-in
sideloading for your account.

## Using it
Open a new email or reply, add someone in "To" — the pane checks Sent Items
automatically. First use will prompt a Microsoft sign-in popup (this is
separate from your Outlook login; it's what lets the add-in call Graph).

## Notes / things you might want to change
- **7-day window** is hardcoded in `taskpane.js` (`findRecentSend`) —
  change the `setDate(since.getDate() - 7)` line to adjust.
- **"1:1 only"** is enforced by requiring the past sent message had exactly
  one "To" recipient and no CC — adjust the `isSoleRecipient` check in
  `taskpane.js` if you want to loosen that (e.g. allow CC but not BCC).
- Currently checks **every** address in "To" if you add more than one —
  each gets its own banner.
