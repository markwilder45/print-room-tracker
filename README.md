# Print Room Tracker

Floor log, volume dashboard, time study and staffing forecast for the print department
(posters, vinyl/decals, canvas, wrapping paper singles and bulk, ribbons, ScanCafe photo prints).

- **Page:** `index.html`, served by GitHub Pages.
- **Data:** a Google Sheet. `apps-script/Code.gs` runs inside the sheet and saves every entry as a row
  on the `Entries` tab and the setup (products, steps, hourly targets, people, printers) on the `Config` tab.

## How it fits together

1. People open the GitHub Pages link and enter the team passcode once per device.
2. Timers and entries are sent to the sheet's Apps Script web app, which checks the passcode and writes the row.
3. Every open copy of the page checks the sheet for changes every 20 seconds.

The page is public, but it shows nothing and saves nothing without the passcode. The passcode lives only
in the copy of `Code.gs` inside the Google Sheet. The copy in this repo says `CHANGE-ME` on purpose.

## Changing the passcode

1. Open the sheet → Extensions → Apps Script.
2. Edit the `PASSCODE` line and save.
3. Deploy → Manage deployments → edit (pencil) → Version: New version → Deploy.
   The web app URL stays the same.

## Updating the script

When a new `Code.gs` comes along, paste it over the old one in Extensions → Apps Script, put your passcode
back on the `PASSCODE` line, save, then Deploy → Manage deployments → edit (pencil) → Version: New version → Deploy.
The web app URL stays the same.

## Raw data

The `Entries` tab has one row per task: date, start, end (epoch ms plus readable times), minutes, employee,
product, step, labor or printer time, units, distributor, printer and note. `minutes` is worked time:
the span from start to end minus `pausedMs` (time the timer sat paused). `pausedAt` is filled in only while a timer is paused. Deleted entries stay as rows
with `deleted = TRUE` so every device can sync the deletion. Don't edit the `id` or `updatedAt` columns by hand.
