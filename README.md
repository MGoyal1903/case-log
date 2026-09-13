# Case Log

A personal case tracker: dashboard, searchable case log, per-case document uploads, CSV exports, and a PIN lock. Your data lives in a Google Sheet you control.

## Files

- `index.html` — page structure
- `style.css` — all styling
- `script.js` — app logic
- `seed-data.js` — your original 50 cases (only used the very first time the app runs, before you've added anything of your own)

## Hosting it on GitHub Pages (free, gives you a permanent URL)

1. Create a new repository on GitHub (Settings icon → or the "+" in the top right → New repository). Any name, e.g. `case-log`. Keep it **Public** (GitHub Pages on a free account needs the repo to be public) or use a **Private** repo if you have GitHub Pro/Team/Enterprise, which supports private Pages.
2. Upload all four files in this folder (`index.html`, `style.css`, `script.js`, `seed-data.js`) — drag and drop them into the repo via "Add file → Upload files" on GitHub's website, or push them with git if you're comfortable with that.
3. Go to the repo's **Settings → Pages**.
4. Under "Build and deployment", set **Source** to `Deploy from a branch`, choose the `main` branch and `/ (root)` folder, then **Save**.
5. GitHub will give you a URL shortly after, usually `https://<your-username>.github.io/<repo-name>/`. That's your permanent link — bookmark it on your phone, laptop, anywhere.

## First time you open it

- You'll be asked to set a PIN (stored in that browser only — you'll set one per device).
- Go to the **Database** tab and connect your Google Sheet (see the in-app instructions and the `Copy script` button for the Apps Script code). This is what makes your data sync across every device that opens the URL.

## Notes

- Because this is now served over `https://`, browser storage (your PIN, your Sheets URL) is far more reliable here than it was opening the file directly from disk.
- If you ever update these files, just re-upload the changed one(s) to the same GitHub repo — the live URL updates automatically within a minute or so.
