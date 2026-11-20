# Sheet API — add foods to the `dict` tab remotely

A small Google Apps Script web app, attached to the macros Google Sheet, that lets
Claude (or anything holding the token) list and add rows to the `dict` tab.

## One-time setup (about 5 minutes)

0. **Convert to a native Google Sheet (if needed).** If the title bar shows an `.XLSX`
   badge, the file is an uploaded Excel file and Apps Script can't attach to it.
   Use **File → Save as Google Sheets**. This creates a new file with a new ID. From then on,
   use the new file and update `link.txt` to
   `https://docs.google.com/spreadsheets/d/<NEW_ID>/export?format=xlsx`
   (the sheet must be shared "Anyone with the link can view" for `export_snapshots.py`).
1. **Open the Google Sheet** → **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` with [`Code.gs`](Code.gs) from this folder. Save.
3. **Set the token:** ⚙ Project Settings → *Script Properties* → Add property
   `TOKEN` = a long random string, e.g. generated with
   `python -c "import secrets; print(secrets.token_urlsafe(32))"`.
4. **Deploy:** Deploy → New deployment → type *Web app*
   - Execute as: **Me**
   - Who has access: **Anyone** (the token is what protects it)
   - Authorize when prompted (Google shows an "unverified app" warning for your own
     script — Advanced → Go to project).
   - Copy the **Web app URL** (ends in `/exec`).
5. **Give the URL + token to the client.** For Remote Control (session on your own PC), put
   both `KEY=value` lines in `sheet_api/.env` (gitignored). For cloud sessions, instead add them as
   environment variables in the cloud environment settings:
   ```
   MACROS_SHEET_URL=https://script.google.com/macros/s/XXXX/exec
   MACROS_SHEET_TOKEN=<the token>
   ```
   and make sure the environment's network access allows `script.google.com` and
   `script.googleusercontent.com` (use *Full* network access if custom domains
   aren't available).
   **Do not commit the URL or token to the repo.**

To test locally (PowerShell):
```powershell
$env:MACROS_SHEET_URL="https://script.google.com/macros/s/XXXX/exec"
$env:MACROS_SHEET_TOKEN="..."
python sheet_api/dict_api.py list
```

## Updating the script later

After editing `Code.gs`, use **Deploy → Manage deployments → ✎ Edit → Version: New
version → Deploy**. This keeps the same URL. ("New deployment" would create a new URL.)

## What it does

- `list` — returns all `dict` rows (name, unit, carbs, protein, fat, calories).
- `upsert` — appends a new row (columns A–F) and writes the validation formula
  `=4*C{r} + 4*D{r} + 9*E{r}` in column G. If the name already exists, it refuses
  with `error: exists` unless `overwrite: true`, in which case it updates B–F in place.
  Names must be unique and exact, because the `food` tab looks them up with `VLOOKUP`.

## Rotating the token

Change the `TOKEN` script property (no redeploy needed) and update the env var.
