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

## Emailed forecast reports

The Forecast tab's **Email report** button sends the forecast shown on screen to the addresses on Setup →
Forecast email list. The sheet's script sends it with Gmail from the Google account that owns the sheet, so
replies go to that account. Recipients always come from the saved list, never from the page.

- Limits: 25 recipients per report and 20 reports a day. A personal Gmail account can send to about 100
  recipients a day through scripts in total.
- Every report sent is logged on the sheet's `Reports` tab.
- After pasting a script version that adds email, run `authorizeEmail` once from the Apps Script editor and
  approve the "send email as you" permission, then deploy a new version.

## Yesterday's results and productivity

The Forecast tab's Day view starts with Yesterday's results. The tech types each person's hours worked from the
timeclock; everything else comes from the floor log for that date.

- **Earned hours:** for each hands-on step, units completed ÷ that step's target rate in Setup. Printer steps are
  left out. Maintenance and clean-up are credited at the time logged. Steps with no target earn nothing and are listed.
- **Productivity:** earned hours ÷ hours worked, per tech and for the department. The goal (95% by default) is set
  on Setup. Targets are the standard, so only managers should change them or use "Use actual".
- **Planned vs finished:** each day report stores its planned units on the `Reports` tab (`reportDate`, `plan`),
  and the next day's results compare finished units against them. Needs script version 5.

## Orders and carry-over

- **Orders:** the finishing step (Pack & ship, or Package) asks for units and orders side by side; orders starts at 1.
  Stored in the `orders` column (script version 6).
- **Carry-over:** for each product, units started at the first step minus units shipped, carried day to day from the
  "Count carry-over from" date on Setup. The morning report lists products with open units and needs a short note for
  each before it sends. Units that aren't really pending can be cleared with a reason: "Shipped, but not logged" counts
  them as shipped; "Misprint or scrap" and "Entered by mistake" just remove them. Clears are rows on `Entries` with
  stepId `_shipped`, `_scrap` or `_error` and kind `adjust`; deleting one from the Day log undoes it.
- **Output vs target:** the work finished at target pace divided by hours worked, with units done vs expected per step.
- The forecast now uses the planned units at every step (the old learned step ratios are off).

## Raw data

The `Entries` tab has one row per task: date, start, end (epoch ms plus readable times), minutes, employee,
product, step, labor or printer time, units, distributor, printer and note. `minutes` is worked time:
the span from start to end minus `pausedMs` (time the timer sat paused). `pausedAt` is filled in only while a timer is paused. Deleted entries stay as rows
with `deleted = TRUE` so every device can sync the deletion. Don't edit the `id` or `updatedAt` columns by hand.
