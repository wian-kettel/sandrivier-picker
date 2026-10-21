# Sandrivier picker tracker

Bag counting per picker for the Sandrivier pear harvest (Early Light Trading, Barrydale).
Replaces the quality controller's clipboard: each picker wears an NFC band, the quality controller
taps it on the farm phone for every full bag, and the counts land in a Google Sheet with a dashboard.

**No picker data is stored in this repository.** Names, numbers and bag counts live only in the
farm's private Google Sheet. The farm PIN keeps anyone else from sending scans.

## Parts
| Part | Where | What it does |
|---|---|---|
| Scanning screen (`index.html`) | GitHub Pages, opened in Chrome on the Android phone (Blackview FORT 1) | Reads bands, one tap = one bag, blocks repeat taps, works without signal, sends scans when signal returns |
| Phone test (`test.html`) | GitHub Pages | Checks the phone can read bands in Chrome |
| Google Sheet backend (`apps-script/`) | Pasted into the Sheet's Apps Script | Stores pickers and scans, dashboard, end of day bag-vs-bin check, Excel download |

## Rules built in
- Picker numbers are the ID. Bands are linked to numbers, so a replaced band keeps the picker's history.
- Pre-loaded list: paste picker numbers and names (and Day/Contract) into the Pickers tab with Tag UID empty.
  On the phone, tapping a new band and typing the number links it. Name only needed for someone not on the list.
- Replace band: tap the new band, type the picker number, confirm "replace". The old band is kept in "Old tags".
- Same picker again inside the repeat window (Settings: `Repeat window minutes`, default 2): **not counted**, logged as DUPLICATE.
- Manual mode: **Type number** stays open, number then Count bag, for a day with broken or forgotten bands (logged MANUAL).
- Day from paper: Details page, one line per picker "number bags" (logged PAPER).
- Scans are sent in batches every 30 seconds (or at 100 waiting). Works without signal.

## Dashboard
- **Today:** export bins (estimate) against the target, total bins, export share, run rate and forecast to the end of the
  last picking session, top 10 and bottom 10 (per hour, breaks excluded), items that need attention, and Fanus's actual bins.
- **Details:** every picker, fastest/slowest pairs, 7-day top 10, bags per block and per hour, day-from-paper entry.
- Estimates use `Bags per bin` (44) and `Expected export %` until Fanus has entered actual bins; after that they learn
  from the last 7 entered days. Target: `Export bin target` (72). Breaks: `Picking sessions`.

## Load test
Run `loadTest` in the Apps Script editor: about 7,500 scans for 50 demo pickers, sent like two phones would.
Timings land on the "Load test" tab. Run `clearDemo` afterwards.

## Setting up a phone
1. Open `https://wian-kettel.github.io/sandrivier-picker/` in **Chrome** on the phone.
2. Menu > Phone setup: paste the Google web app link (ends in `/exec`), the farm PIN (Settings tab of the sheet), team and a phone name.
3. Chrome menu > Add to Home screen. Open it from the home screen from then on.
4. Press START each morning (Android asks for NFC permission the first time).

## Updating the Google Sheet side
Paste `apps-script/Code.gs` and `apps-script/Dashboard.html` into the sheet's Apps Script, run `setup`,
then Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy. The web app link stays the same.

Built by Kettel (Pty) Ltd.
