# Thusa

A chat-to-calendar meeting assistant, built as an installable PWA for Android (Chrome) and iPhone/iPad (Safari).

Type a meeting request in plain English → **Google Gemini** (`generateContent` + function calling)
parses it into a structured event → the app creates it in your **Microsoft 365** calendar via
Microsoft Graph and sends the invites. Timezone: Africa/Gaborone.

**→ See [SETUP-GUIDE.md](SETUP-GUIDE.md) for the one-time setup (button by button).**

## How it works

- **Vanilla JS PWA** — no build step. `index.html` + `app.js` + `styles.css` + `sw.js` + manifest.
- **AI:** Google Gemini `generativelanguage.googleapis.com` `:generateContent`, called directly
  from the browser (API key as query param — the endpoint allows CORS). Model: `gemini-2.5-flash`.
  A single `create_meeting` function declaration does the structured extraction.
- **Calendar:** Microsoft Graph `POST /me/events` via [MSAL.js v3](https://github.com/AzureAD/microsoft-authentication-library-for-js)
  (SPA auth-code + PKCE, token cache in `localStorage`). Sign in once; silent renewal after.
- **Storage:** everything (Gemini key, Azure client/tenant IDs, team directory) lives only in
  the browser's `localStorage`. **No secrets in this repo**, no backend, no server.

## Privacy

Chat text → Google Gemini. Meeting details → Microsoft. Nothing else leaves the device.

## Local dev

```sh
python -m http.server 4173
# open http://localhost:4173/
```

Deployed via GitHub Pages. The Azure app registration's redirect URI must match the Pages URL
exactly (trailing slash included) — see the setup guide.
