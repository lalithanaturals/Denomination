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

## Item Catalog sync (new)

Cloud Sync now also mirrors your **Item Catalog** (the barcode → item
lookup used in Cash Bill) across devices, using the same Web App URL —
there is nothing extra to set up. Items you add or edit on one device are
pushed to a new sheet tab called **"ItemCatalog"**, and each device pulls
the latest catalog the first time you open Cash Bill in a session. If you
want to force an immediate sync (e.g. right after adding several new items
on another device), go to **Settings > ☁️ Cloud Sync** and click
**☁️ Sync Catalog Now**.

## If you ever need to change the script

If you (or a future update) paste in an updated version of `Code.gs`, you
must tell Google to actually use the new code — simply pasting and saving
is not enough. In plain terms:

1. Go back to the Apps Script editor tab (Extensions > Apps Script) and
   paste in the new `Code.gs` contents, then save (the disk icon or
   Ctrl/Cmd+S).
2. Click **Deploy** (top-right) > **Manage deployments**.
3. Click the pencil/edit icon next to the existing deployment.
4. Where it says "Version", change the dropdown to **New version**.
5. Click **Deploy**.

You do **not** need to create a brand-new deployment, and you do **not**
need to copy a new URL or update it in the app's Settings — the existing
`/exec` URL keeps working exactly as before, it just now runs the updated
code behind it. Skipping steps 2–5 (i.e. only pasting and saving the code)
means the live URL keeps running the OLD code, so this step is easy to
forget but important.
