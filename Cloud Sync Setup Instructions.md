# Cloud Sync Setup Instructions (Google Sheets)

This is a one-time, free setup that lets your shift register data show up
in a Google Sheet you own — viewable from any phone, tablet or computer,
not just the one device the app runs on. You do this once; after that, the
app quietly mirrors every saved shift to your Sheet in the background.

**The app works perfectly fine without ever doing this.** Cloud Sync is
optional. Skip this entirely if you only ever use the app on one device.

## Step 1 — Create a Google Sheet

1. Go to **sheets.google.com** (free with any Google/Gmail account).
2. Click **Blank** to create a new spreadsheet.
3. Give it a name at the top, e.g. "Lalitha Naturals — Shift Register Sync".

## Step 2 — Open the Apps Script editor

1. In the menu bar, click **Extensions > Apps Script**.
2. A new tab opens with a code editor and a file called `Code.gs`, containing
   some placeholder text like `function myFunction() {}`.
3. Click inside that editor, select everything (Ctrl+A or Cmd+A), and delete it.
4. Open the file named **`Code.gs`** that came with this app (in the same
   folder as the register HTML file), select all of its text, and paste it
   into the now-empty Apps Script editor.
5. Save it (the disk icon, or Ctrl/Cmd+S). You can rename the project at the
   top (e.g. "Shift Register Sync") if you like — this is optional.

## Step 3 — Deploy it as a Web App

1. Click the blue **Deploy** button (top right) > **New deployment**.
2. Click the gear/settings icon next to "Select type" and choose **Web app**.
3. Fill in the form exactly like this:
   - **Description:** Shift Register Sync (or anything you want)
   - **Execute as:** **Me** (your own Google account)
   - **Who has access:** **Anyone with the link**
4. Click **Deploy**.
5. The first time, Google will interrupt with an authorization screen because
   this is a script you just pasted in:
   - Click **Review permissions**
   - Choose your Google account
   - You'll see a warning screen saying "Google hasn't verified this app" —
     this is expected for your own personal script. Click **Advanced**, then
     click **Go to (your project name) (unsafe)**.
   - Click **Allow**.
6. You'll now see a **Web app URL** ending in `/exec`. Click the copy icon
   next to it, or select and copy it manually. This is your Cloud Sync URL.

## Step 4 — Paste the URL into the app

1. Open the Shift Register app.
2. Click **⋯ More > Settings** (or wherever the Settings button is in the
   header).
3. Go to the **☁️ Cloud Sync** tab in the left-hand list.
4. Paste the URL you copied into the **Google Apps Script Web App URL** field.
5. Tick **Enable Cloud Sync**.
6. Click **🔌 Test Connection** to confirm the app can reach it.
7. Click **Save Settings**.

That's it. From now on, every time a shift is saved with **Save & Commit
Entry**, it's also sent to your Google Sheet in the background — the local
save always happens first and is never slowed down or blocked by this.

## If something doesn't sync

- If you're offline when a shift is saved, or the URL is temporarily
  unreachable, the app keeps that record queued locally (nothing is lost)
  and shows how many records are waiting in the Cloud Sync settings panel.
- Click **☁️ Sync Now** any time (e.g. once you're back online) to retry
  every queued record.
- Open your Google Sheet directly and look for a tab named
  **"ShiftRegister"** — that's where the rows land, one per shift.

## If you ever need to change the script

If you paste in an updated version of `Code.gs` later, you must go back to
**Deploy > Manage deployments**, click the pencil/edit icon on the existing
deployment, and choose **New version** before it takes effect — editing the
file alone does not update the live URL.
