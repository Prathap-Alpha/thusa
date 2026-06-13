# Thusa — Setup Guide

Thusa is your meeting assistant. You type ("Set up a meeting with Kago at 5pm today about
3D accounts") and it creates the calendar invite and sends it from **your own** Outlook
calendar. The clever part is Claude (the same AI as Javis); the calendar part is Microsoft 365.

You set this up **once**. After that, you open it like any app and just chat.

There are three things to give it: your **team list**, an **AI key**, and a **Microsoft
connection**. Sections 1–3 below. The only fiddly bit is the Microsoft connection (Section 3),
and that is a one-time job.

---

## What you'll end up with

A little orange "T" icon on your S25 Ultra home screen. Tap it, type, done.

---

## Step 0 — Open it on your phone

1. On your Samsung, open **Chrome**.
2. Go to: **https://prathap-alpha.github.io/thusa/**
3. Chrome will show a banner **"Add Thusa to Home screen"** (or tap the **⋮** menu top-right →
   **Add to Home screen** → **Install**).
4. Now you have the **T** icon. Open it from there from now on — it runs full-screen like a real app.

The first time it opens it will land on the **Setup** screen because nothing is filled in yet.

---

## Step 1 — Your AI key (Anthropic Claude)

This is what makes Thusa understand plain English.

1. On a computer, go to **https://console.anthropic.com** and sign in with the Alpha Direct
   Anthropic account.
2. Left menu → **API Keys** → **Create Key**. Name it `Thusa`. Copy the key (starts with `sk-ant-`).
3. In Thusa, open the **⚙️ Setup** tab.
4. Paste the key into **API key**.
5. Leave **Model** as `claude-haiku-4-5` (fast and cheap — our approved vendor). Don't change it
   unless you have a reason.

> The key lives only on your phone. It is never sent anywhere except Anthropic.

---

## Step 2 — Your team list

Thusa can only invite people it knows by first name.

1. Open the **👥 Team** tab.
2. Type a **First name** and their **email**, tap **Add**. Repeat for each person.
3. Faster way: tap **"Paste a whole list at once"**, paste lines like:
   ```
   Kago, kago@alphadirect.co.bw
   Arun, arun@alphadirect.co.bw
   Unami, unami@alphadirect.co.bw
   ```
   then tap **Import list**.

> The list is stored **only on your phone**. You can edit it any time.

---

## Step 3 — Connect Microsoft 365 (the one-time fiddly bit)

Thusa never sees your password. Instead it uses Microsoft's official sign-in. But first,
Microsoft needs to be told that "Thusa is allowed to ask." That's a one-time registration in
Azure. **If you'd rather, forward this section to IT / TheRiskCo and ask them to send you back
the two IDs in Step 3.2** — then you only do Step 3.3.

### 3.1 — Register Thusa in Azure (once)

1. On a computer, go to **https://entra.microsoft.com** (or portal.azure.com → "Microsoft Entra ID").
   Sign in as **pganesharajah@alphadirect.co.bw**.
2. Left menu → **App registrations** → **+ New registration**.
3. **Name:** `Thusa`
4. **Supported account types:** choose **"Accounts in this organizational directory only
   (Alpha Direct only — Single tenant)"**.
5. **Redirect URI:** in the dropdown pick **Single-page application (SPA)** and paste:
   ```
   https://prathap-alpha.github.io/thusa/
   ```
   (include the trailing slash, exactly).
6. Click **Register**.

### 3.2 — Copy the two IDs

On the app's **Overview** page you'll now see:
- **Application (client) ID** — a long code like `1a2b3c4d-...`
- **Directory (tenant) ID** — another long code.

Copy both.

### 3.3 — Give it calendar permission

1. Still in the Thusa app registration, left menu → **API permissions**.
2. **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**.
3. Search and tick **`Calendars.ReadWrite`** and **`User.Read`** → **Add permissions**.
4. (If a **"Grant admin consent for Alpha Direct"** button is shown, click it. If it's greyed
   out, ask IT to click it — it just pre-approves so you aren't nagged.)

### 3.4 — Put the IDs into Thusa and sign in

1. Back in the Thusa app on your phone, **⚙️ Setup** tab, section 2.
2. Paste the **Application (client) ID** and the **Directory (tenant) ID**.
3. Tap **Save settings**.
4. Tap **Sign in with Microsoft**. Microsoft's own page opens.
5. Sign in as **pganesharajah@alphadirect.co.bw**, and when it asks **"Stay signed in?"** tap
   **Yes**. ← this is what stops it asking again.
6. It bounces back to Thusa and the badge at the top turns green: **"connected"**.

That's it. It will not ask for your password again.

---

## Using it day to day

Open the **T** app, **💬 Chat** tab, and type naturally:

- *"Set up a meeting with Kago at 5pm today regarding 3D accounts."*
- *"Meeting with Arun tomorrow morning about performance appraisal."*
- *"30 min with Unami at 8am Monday — staff dismissals. Boardroom, no Teams link."*
- *"Meeting with Kago and Arun Thursday 2pm about the Q3 numbers."* (several people at once)
- You can even line up a few in one message — it makes one invite per meeting.

Thusa shows you a **card** for each meeting (subject, time, who, where). Tap **Send invite** to
fire it, or **Cancel**. Invites go out by email from your calendar with a Teams link by default.

> Want it to skip the card and just send immediately? **Setup → section 3 → tick "Send invites
> immediately."** I'd leave it off at first so you can eyeball each one.

---

## Notes & troubleshooting

- **Times** are Botswana time (UTC+2). "Tomorrow morning" with no time becomes 08:00 and it tells
  you it assumed that.
- **"Not signed in" after a while:** open Setup, tap **Sign in with Microsoft** once more. With
  "Stay signed in" ticked this is rare.
- **"API key was rejected":** re-check the key in Setup (Step 1).
- **A name isn't recognised:** add the person in the **Team** tab, or type their full email in
  the chat.
- **Everything** (key, IDs, team) is stored **only on this phone**. Reinstalling clears it — just
  redo Steps 1–3 (Azure registration in 3.1 stays, so you only re-enter the IDs and sign in).
- **Privacy:** your chat text goes to Anthropic (to understand it); meeting details go to Microsoft
  (to create the invite). Nothing goes anywhere else, and the app code holds no secrets.
