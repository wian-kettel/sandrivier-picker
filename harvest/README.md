# Harvest Tracker (Sandrivier)

One web app for the Sandrivier pear harvest (Early Light Trading, Barrydale):
FORT 1 band scanning, picker performance, daily bins, day-worker pay and loans,
season comparison, orchards and bloom dates.

- App: https://wian-kettel.github.io/sandrivier-picker/harvest/
- Prototype with demo data: https://wian-kettel.github.io/sandrivier-picker/harvest/prototype/
- Data: Firebase Firestore (Kettel's project), with an offline copy on every device.
- Logins: Firebase email and password. Access is set per person in Settings (admin only).

## Files
- `index.html`: the whole app (office screens and the FORT 1 scanner screen).
- `config.js`: the Firebase project settings (not secret).
- `firebase/firestore.rules`: who may read and write what. Paste into Firebase console, Firestore, Rules.
- `sw.js`: lets the app open with no signal.

## Who sees what
| Access | Who | Can |
|---|---|---|
| Admin | Wian | Everything, plus give people access |
| Office | Fanus, Adolf | Everything (pay, loans, corrections, settings) |
| FORT 1, team 1 or 2 | the team lead phones | Count bags, link bands, pick the block. No pay, loans or change log. |
| Waiting | anyone new | Nothing until the admin gives access |

## Rules worth knowing
- One tap = one bag. A second tap of the same band within the repeat window (Settings) is not counted.
- Bags per picker per day are counters, so two devices can never overwrite each other.
  Every tap is also kept in `scans` and can never be changed from the app.
- Corrections are added on top of the scanned count, with a reason, and show in the change log.
- Pay week runs Friday to Thursday. Once a week is marked as paid its loans are locked (also enforced by the rules).

## Testing locally
`?emu=1` connects to the Firebase emulators (Firestore 8080, Auth 9099). `?test=1` shows a "simulate a band tap" button on the scanner screen.
