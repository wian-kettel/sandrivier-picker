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
- Unknown band: offers to register it (picker number, name, Day or Contract). Needs signal.
- Same band again inside the repeat window (Settings: `Repeat window minutes`, default 2): **not counted**, logged as DUPLICATE.
- Lost or forgotten band: **Type number** counts the bag by picker number, marked MANUAL so it shows on the dashboard.
- Each phone is set to Team 1 or Team 2; the block is chosen on the phone each morning.
- End of day: enter plastic and wooden bins per team on the dashboard. With `Kg per bag`, `Kg per plastic bin`
  and `Kg per wooden bin` set on the Settings tab, the dashboard estimates how full the bags were on average.

## Setting up a phone
1. Open `https://wian-kettel.github.io/sandrivier-picker/` in **Chrome** on the phone.
2. Menu > Phone setup: paste the Google web app link (ends in `/exec`), the farm PIN (Settings tab of the sheet), team and a phone name.
3. Chrome menu > Add to Home screen. Open it from the home screen from then on.
4. Press START each morning (Android asks for NFC permission the first time).

## Updating the Google Sheet side
Paste `apps-script/Code.gs` and `apps-script/Dashboard.html` into the sheet's Apps Script, run `setup`,
then Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy. The web app link stays the same.

Built by Kettel (Pty) Ltd.
