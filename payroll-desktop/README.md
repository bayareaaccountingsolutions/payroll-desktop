# Payroll System — Desktop App (Windows + Mac)

A desktop app for your **Attendance, Payroll & Salary Slip** Google Sheet (script v2.3).
All data stays in the **one shared Google Sheet**, so several staff can work at the same
time from Windows or Mac. The app updates itself when you publish a new version.

| Screen | What it does |
|---|---|
| Dashboard | Today's attendance, payroll of any month, Jan–Dec charts and overview |
| Attendance | Month grid — click a cell or press 1–5 / P W H L A, then **Save**. Payroll recalculates instantly |
| Payroll Register | Every employee's earnings, deductions, advance recovery and net pay; PDF and Excel (CSV) |
| Salary Slip | Preview, then the exact sheet slip as PDF; email it to the employee (admin) |
| Employees | Employee Master with salary heads, status and exits |
| Google Sheet | Opens the live sheet inside the app (all original menus and buttons work) |
| Admin Tools | Recalculate, sync, mark offs, self-test, finalize / unlock month, update sheet script |
| Settings | Connection, updates, light / dark theme |

The payroll maths is **not duplicated** — the app calls the same engine in `Code.gs`, so
figures are always identical to the sheet. Every change made from the app is recorded in
the sheet's hidden **App Log** tab with the user's name.

---

## Part A — Connect the Google Sheet (admin, once, ~5 min)

1. Open your payroll Google Sheet ▸ **Extensions ▸ Apps Script**.
2. Your existing code stays in `Code.gs` — do not change it.
3. Click **＋ ▸ Script**, name it `DesktopApi`, paste the contents of `apps-script/DesktopApi.gs`, click **Save**.
4. In the function list choose **setupDesktopApp** ▸ **Run** ▸ allow the permissions.
5. A pop-up shows two keys — **ADMIN key** (owner / HR head) and **STAFF key** (attendance & slips). Keep them safe.
6. **Deploy ▸ New deployment** ▸ gear icon ▸ **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone** (the keys protect the data — without a key nothing is returned)
   - **Deploy** ▸ copy the **Web app URL** (ends in `/exec`).
7. Lost a key or a staff member left? Run **resetDesktopKeys** — old keys stop working at once.

## Part B — Publish the app (admin, once, ~15 min)

Auto-update works through **GitHub Releases** (free).

1. Create a free account at github.com and a **new private or public repository** named `payroll-desktop`.
   - Auto-update downloads must be reachable by the app: use a **public** repo, or a private repo with a read token (ask your IT person).
2. In `package.json`, replace `YOUR-GITHUB-USERNAME` with your GitHub user / organisation name.
3. Upload this whole folder to the repository (GitHub Desktop or `git push`). Do **not** upload `node_modules` or `dist`.
4. Create the first release: in the repo ▸ **Releases ▸ Draft a new release ▸ Choose a tag ▸ type `v1.0.0` ▸ Create tag ▸ Publish**.
   (Or with git: `git tag v1.0.0 && git push origin v1.0.0`.)
5. The **Actions** tab builds the installers (about 10 minutes):
   - Windows: `Payroll-System-Setup-1.0.0.exe`
   - Mac: `.dmg` for Apple-silicon (`arm64`) and Intel (`x64`)
   They appear under **Releases**.

## Part C — Install on each computer (staff)

1. Download the installer from the release page and run it.
   - **Windows**: if SmartScreen appears ▸ **More info ▸ Run anyway** (shown only because the app is not code-signed).
   - **Mac**: open the `.dmg`, drag *Payroll System* to Applications. First launch: **right-click ▸ Open ▸ Open**.
2. Open the app ▸ **Settings**: enter your name, the **Web App URL** and your **key** ▸ **Test & Save**.
3. Done — the Dashboard loads from the shared sheet.

## Part D — Updating and adding features

### D1. Releasing a new app version (new screens, fixes, features)
1. Change the code in `src/` (or ask Claude to add the feature).
2. Raise `"version"` in `package.json` — e.g. `1.0.0` → `1.1.0`.
3. Commit, then push a matching tag: `git tag v1.1.0 && git push origin v1.1.0`.
4. GitHub builds and publishes it. Every installed app checks at start-up and every 4 hours:
   - **Windows**: downloads silently, then shows **Restart & update**.
   - **Mac**: shows **New version available ▸ Download** (silent install on Mac needs an Apple Developer ID, $99/year — add the secrets listed in `.github/workflows/release.yml` and set `"macSigned": true` in `package.json` to enable it).
5. Payroll data is never touched by an app update — it lives in the Google Sheet.

### D2. Updating the sheet script (new payroll rules / tabs)
1. Put the new `Code.gs` and/or `DesktopApi.gs` in `apps-script/`, and raise `SHEET_SCRIPT_VERSION` (and `DESKTOP_API_VERSION` if the API changed) in `DesktopApi.gs`.
2. Release the app (D1). Admins then see "sheet script update available".
3. **Admin Tools ▸ Sheet script ▸ Copy** each file ▸ paste into Apps Script ▸ Save.
4. **Deploy ▸ Manage deployments ▸ ✏ Edit ▸ Version: New version ▸ Deploy** — the URL stays the same.

### D3. Adding a feature that needs new sheet data
Add an entry to `API_ACTIONS_` in `DesktopApi.gs`, call it from the app with `call('yourAction', {...})`,
then release as in D1 + D2. Old app versions ignore new actions safely; a new app talking to an old
sheet shows "Unknown action — please update the sheet script".

## Part E — Try it without Google (demo mode)
```
npm install
PAYROLL_DEMO=1 npm start          # Mac / Linux
set PAYROLL_DEMO=1 && npm start   # Windows (cmd)
```

## Troubleshooting
| Message | Fix |
|---|---|
| "Google returned a sign-in page" | Deployment ▸ Who has access must be **Anyone**; use the `/exec` URL, not `/dev`. |
| "Access key is not valid" | Key was reset — get the new key from the admin. |
| "Unknown action" | The sheet script is older than the app — Part D2. |
| Google sheet window says "browser not secure" | Use **Google Sheet ▸ Open in my browser**. All other screens are unaffected. |
| Figures look stale | ↻ Refresh, or Admin Tools ▸ Recalculate all months. |
| "… is finalized (locked)" | Admin Tools ▸ Unlock the month, make the change, Finalize again. |

## Limits to know
- Each request runs on Google's servers; a save takes about 2–6 seconds (the same recalculation the sheet does).
- PDF slips are created one at a time (Google locks the Salary Slip tab briefly so two users never mix up slips).

## Project layout
```
apps-script/Code.gs        your original system (unchanged)
apps-script/DesktopApi.gs  the secure API add-on
src/main.js                window, Google connection, auto-update
src/preload.js             safe bridge between window and app
src/renderer/              screens (index.html, app.js, styles.css)
src/mock.js                demo data (PAYROLL_DEMO=1 only)
scripts/check.js           pre-release checks (run by the build)
.github/workflows/         builds Windows + Mac installers on each tag
```
