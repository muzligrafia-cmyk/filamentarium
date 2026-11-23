# Filamentarium

A 3D printing manager that works as a front end for a Google Spreadsheet. Your data stays in the sheet: print log, filament stock, project orders and calibration notes. Filamentarium is a single web page for reading the sheet and adding to it quickly from your phone or PC.

The interface is in Hungarian.

## What it does

- **Napló (print log):** log a print in a few taps: client, project, model, pieces, grams, waste, time, filament, printer, settings, support and status. You can show it as cards or as a list with resizable columns, and colour-code it by project. Running prints (`Fut` / `Felvéve`) appear at the top with one-tap **Kész** / **Rossz** buttons.
- **Slicer file import:** drop a sliced file onto the page, or use **Fájl betöltése** in *Új nyomtatás*. It fills in the print time, grams, filament, printer and settings for you.
  - `.gcode` files from Creality Print, OrcaSlicer, PrusaSlicer and Bambu Studio
  - `.gcode.3mf` files from Bambu Studio (*Export plate sliced file*). If the file has several sliced plates, or you drop several files, each one becomes its own entry, opened one after the other.
  - Only the small summary parts of the file are read. Nothing is uploaded anywhere, and even 30 MB files load instantly.
- **Készlet (stock):** remaining grams per filament, plus deliveries (*Beérkezés*) and new filament types. You can sort, hide empty spools and switch between card and list view. Low stock shows as *Fogyóban*; 0 g shows as *Elfogyott* in red.
- **Projektek (projects):** ordered vs. printed pieces per item, with progress bars. You can add projects and items (*+ Új tétel*).
- **Kalibráció (calibration):** your filament calibration library: temperatures, flow ratio and pressure advance.
- **Statisztika (statistics):** print hours, material use, waste and failure rate, a 12-month chart, and breakdowns by printer, client, filament and project.
- **Phone friendly:** large touch targets, and the Android back button steps back through screens. New prints and deliveries are kept and sent later if you're offline.

## How it fits together

```
 Filamentarium (index.html)  ──►  Apps Script web app (Code.gs)  ──►  Your Google Sheet
   browser / phone                  runs inside your sheet               the only data store
```

`Code.gs` adds a small private web address to your spreadsheet. The app talks to that address using a password you choose.

The script **never changes your sheet's layout**. It only adds rows to the ends of your existing tables, copies your own formula patterns, and edits rows you choose to edit. The sheet's own formulas still do all the calculations.

It writes formulas in your sheet's own syntax: semicolons for Hungarian and other European settings, commas for English.

### Expected sheet layout

The script works with the *3D Munkavezetés* layout:

| Tab | Used for |
|---|---|
| `3D Gyártás` | Print log, columns A–O: date, client, project, model, pcs, g, waste g, total g (formula), time, filament, colour, status, settings, support, printer |
| `Projektek` | Project groups and ordered items, columns B–H |
| `Filamentkészletek` | "Felhasznált … / Elérhető még" blocks in B–E, deliveries in J–O |
| `Kalibrációk` | Calibration notes in row 1, table from row 4 |

The tab names are set at the top of `Code.gs` (`CONFIG`) if yours differ.

## Installation

You only do this once. It takes about 5 minutes.

### 1. Add the script to your Google Sheet

1. Open your spreadsheet in Google Sheets.
2. Click **Extensions → Apps Script** (*Bővítmények → Apps Script*).
3. Delete everything in the editor and paste in the full contents of [`Code.gs`](Code.gs).
4. Near the top, replace `IRD-AT-EGY-SAJAT-JELSZORA` with a password of your own:
   ```js
   SECRET: 'your-own-password',
   ```
5. Click 💾 **Save**.
6. Choose `teszt` in the function dropdown and click **Run**. Google will ask for permission:
   - If it says *"Google hasn't verified this app"*, click **Advanced → Go to (project name) (unsafe)** and then **Allow**. This is normal for your own scripts.
   - The log at the bottom should show how many log rows, projects, filaments and calibrations it found.

### 2. Publish it as a web app

1. Click **Deploy → New deployment** (*Telepítés → Új telepítés*).
2. Click the ⚙️ next to *Select type* and choose **Web app** (*Webalkalmazás*).
3. Set **Execute as: Me** and **Who has access: Anyone**.
4. Click **Deploy** and copy the **Web app URL**. It ends in `/exec`.

"Anyone" only means the address is reachable. Without the password, every request is refused.

### 3. Open the app

The app is one file, `index.html`. Either:

- **Open it locally:** download `index.html` and open it in Chrome or Edge, **or**
- **Use GitHub Pages:** in this repository go to **Settings → Pages**, set **Source: Deploy from a branch**, **Branch: `main`**, folder **`/ (root)`**, then click **Save**. After about a minute the app is live at `https://<your-username>.github.io/filamentarium/`.

### 4. Connect it

1. In the app, open **Beállítások** (settings, the ⚙️ icon).
2. Paste the **Web app URL** from step 2 and type your **password**.
3. Click **Mentés és kapcsolódás**. Your data loads within a few seconds.

Do this once on each device or browser. The address and password are saved only in that browser, never in the HTML file.

### 5. (Optional) Put it on your phone's home screen

In Chrome on Android, open the app's address and tap **⋮ → Add to home screen**. It gets its own icon and opens full screen.

## Updating

- **New `index.html`:** replace the file, or upload the new one to this repository. Devices pick it up on the next refresh.
- **New `Code.gs`:**
  1. Open **Extensions → Apps Script** and paste in the new code, keeping your own password in `SECRET`.
  2. Click **Save**, then **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**.

  The web address stays the same, so the app needs no changes.

## Privacy and security

- Never commit a `Code.gs` that contains your real password. The copy in this repository uses the placeholder on purpose. Your real password lives only in your Apps Script project.
- The HTML contains no addresses, passwords or data. Everything personal comes from your sheet after you connect.
- Project colours are stored in the script's own storage (*Script Properties*), so all your devices share them and the sheet stays untouched.
- View settings are stored in each browser: card or list view, column widths, font sizes and theme.

## Good to know

- **Rows are never deleted** from the app. Deleting rows would shift the hand-picked row ranges in the sheet's formulas, so fix a mistake by editing it or marking it *Rossz* instead.
- **Slicer imports default to status `Fut`**, because you usually load the file when you start a print. Tap **Kész** on the card when it's done.
- If a filament from a slicer file isn't in your log yet, for example *Generic PLA Silk*, the app says so and lets you pick the spool you actually used.

## Files

| File | What it is |
|---|---|
| `index.html` | The whole app: one self-contained file with the Mulish font embedded. No internet needed apart from talking to your sheet. |
| `Code.gs` | The Apps Script back end, with a placeholder password. |
