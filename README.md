# LOE-tracker

Plan your work week, month, and year. Import timesheets, and export
reports & spreadsheets.

A forecast-first weekly ledger for tracking planned vs. actual work hours,
vacation/holiday/sick leave, LOE (Level of Effort) targets, and billable
vs. non-billable time — built for Canadian contractors/consultants working
against an assigned day-rate or hour target. See `VISION.md` for the full
product intent, `DATA_MODEL.md` for how the data works, and `PLAN.md` for
the running log of what's built and why.

Repo: git@github.com:bobbyc-brs/LOE-tracker.git

Run it with:

```
python calendar_app.py
```

Other tools in this repo:

```
python import_timesheets.py [--dry-run] [--force]   # import ./Timesheets/*.pdf into work_log.json
python fill_akkodis_timesheet.py <year> <month>      # fill a monthly xlsx template from work_log.json
```

## Monthly routine

1. Drop the new timecard PDFs into `./Timesheets/`.
2. Preview the import and check what it would add:

   ```
   python import_timesheets.py --dry-run
   ```

   Look for missing weekdays in the "Would add" list (a missing PDF, or a
   day that needs a status set in the app) and for any "Mismatches" lines.
   Already-imported dates are skipped, so re-running is safe.
3. Import for real:

   ```
   python import_timesheets.py
   ```

4. Generate the month's sheet (overwrites any existing one for that month):

   ```
   python fill_akkodis_timesheet.py <year> <month>
   ```

   A `WARNING: ... weekday(s) with no logged entries` line means workdays
   with no hours -- usually a missing timesheet.
5. If the import added dates from the previous month (a week that spans
   the month boundary), regenerate that month's sheet too.

Both are tied to one specific vendor's file format (a particular PDF
timecard export, and a particular monthly-timesheet xlsx template) — not
universal importers/exporters. Treat them as worked samples of "how to
scan a timecard PDF into `work_log.json`" and "how to fill an external
xlsx template from it," not drop-in tools for a different employer's
formats. Adapting either one to a different system means inspecting that
system's actual file layout and writing new parsing/filling logic against
it, the same way these were built (see `timesheet_extractor.py`'s
docstring for how the PDF one was reverse-engineered).

Note: `work_log.json`, `./Timesheets/`, `./Akkodis/*Timesheet-*.xlsx`
(generated output), `./data/`, and `./Screenshots/` are gitignored — they
contain real client/timesheet data, not code.
