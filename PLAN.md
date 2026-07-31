# Plan

Focus on summaries, persistence, and calculations.

## Done

- Period picker (View bar: presets + custom From/To, scopes the ledger
  display and the Forecast Summary; the per-week Bank running total stays
  full-range regardless).
- CSV export (View bar → Export CSV): exports whatever range is currently
  selected in the period picker. One row per day (Date, Day, Status,
  Planned, Worked, Variance, Entries, Notes), with a Week Total row after
  each week and a Period Total at the end.
- Print/PDF (View bar → Print / PDF): renders the selected period as a
  color-coded weekly-card HTML page and opens it in the default browser,
  so the browser's own Print / Save-as-PDF handles the rest. No new
  dependency — reuses the same view-range scoping as CSV export.
- LOE (Level of Effort) tracking (right panel, below Forecast Summary):
  raised by a manager email (2026-07-21, re: "90 days LOE 1 Apr–31 Aug").
  LOE always covers whatever period is currently selected in the picker —
  there's no separate stored LOE-period list. Enter a day target and it
  shows Days (Working=1, Half-Day=0.5, everything else=0) in that period
  vs. the target, with the shortfall/surplus color-coded red/green. Days
  are counted via each day's real/assumed status either way (past = what
  actually happened, future = the current schedule), so vacation/holiday
  days correctly reduce the count rather than the app assuming every
  weekday gets worked.
- Excel export with live formulas (View bar → Export Excel): unlike CSV,
  a real `.xlsx` can hold formulas. One block per week — day columns
  (Sat..Fri) as headers, then Planned/Worked/Status rows, with a `Total`
  column holding `=SUM(...)` over that week's day cells, and a Variance
  cell referencing the two totals. A Period Total block at the end sums
  all the weekly Total cells. Editing a raw hour value in Excel
  recalculates everything above it. Requires the `openpyxl` package
  (installed via pip). CSV stays as the simple/no-dependency option.
- Forecasted vs. worked (`F:` vs `W:`): a `WORKING` day's hours line reads
  `F:<planned>` (e.g. `F:7.5`) instead of `W:0` until it actually has a
  timesheet behind it (logged entries or an override), then it reads
  `W:<credited>` like normal. The cell's color and status letter (still
  always `W` for `WORKING`) don't change — only that one line's label and
  number do. This was the fix for a UI report where Vacation days
  (auto-credited, showing full "worked" hours) looked visually "backwards"
  next to unlogged WORKING days showing `W:0` — now the unsubmitted days
  show what they're scheduled for instead of a worked-hours count that
  hasn't happened yet. Applied consistently in the app, the print report,
  and the Excel export.
- Layout fixes from a UI report (screenshots): the print report's weekly
  Plan/Worked/Variance/Bank block now sits beside the day-cell table
  (matching the on-screen ledger) instead of below it; the sidebar was
  being clipped because the PanedWindow sash was placed before the
  sidebar's own content existed — moved to after `build()`/`_set_view()`
  so it isn't silently overridden; Bulk Edit now only appears once 2+
  days are selected instead of being shown unconditionally.
- Timesheet PDF import (`timesheet_extractor.py` + `import_timesheets.py`):
  these "Time and Expense Community" timecard PDFs (`./Timesheets/TCH-*.pdf`)
  are real text-layer PDFs, not scans, so plain extraction via `pypdf`
  works — no OCR/AI needed. `TimesheetExtractor` parses a single PDF's
  detailed Timer Session Events into `{date: [{hours, description}, ...]}`,
  falling back to the daily summary line (labeled "Imported timecard") for
  the rare timecard with no session detail, and cross-checks the two
  against each other via `verify()`. `import_timesheets.py` runs this over
  every PDF in `./Timesheets` and merges into `work_log.json` — it only
  ever fills in dates that don't already have entries, never overwrites
  existing entries/notes/status, and is safe to re-run as new PDFs show up
  (`--dry-run` previews without saving).
- Non-Billable tracking: policy is that anything from an approved/imported
  timesheet is billable by default (`import_timesheets.py` sets
  `"billable": true` explicitly), while a manually-typed entry in the day
  editor defaults to **non-billable** instead (its "Billable" checkbox
  starts unchecked — check it if the manual entry really is billable).
  Pre-existing entries with no `billable` key at all stay billable; that's
  never retroactively reclassified. `VACATION`/`PERSONAL_SICK` days also
  count their full credited hours as non-billable. Each week's summary
  (app, print report, CSV) shows an extra "Non-Billable: X" line/column
  only when that week's total is nonzero. The Excel export instead always
  includes a "Non-Billable" row per week (value `0` where it doesn't
  apply, with its own `SUM()` total and a Period Total line) rather than
  conditionally showing/hiding it — deliberately, so every week block has
  the same fixed row count and formula row-offsets stay predictable.
- Worked vs. Credited (2026-07-23): "Worked" everywhere (day cells,
  week/period summaries, CSV/print/Excel) now means literal logged
  `entries[].hours`, full stop — it no longer includes vacation/holiday/
  half-day/sick auto-credit. That auto-credit still exists internally as
  "Credited," used only for Variance/Bank, and it's policy-specific:
  Vacation/Holiday/Half-Day are planned exceptions (a manager's LOE
  already accounts for them) so they still net to zero variance; Sick
  leave is unplanned, so it's excluded from that fallback and shows a real
  variance when nothing's logged, same as an unlogged WORKING day. The
  Excel export now has an explicit "Credited" row per week (in addition to
  Worked) so its live Variance formula stays accurate rather than
  reflecting the new Worked-only figure.
- Vacation's Planned hours are now `0`, same as `HOLIDAY` — it's a day off,
  not a day that was "planned then perfectly offset." Non-Billable's
  calculation for Vacation/Sick no longer piggybacks on `planned` (which
  would now be 0 for vacation) — it uses `daily_hours` on a would-be
  workday directly instead, so vacation still correctly counts as
  non-billable time off.
- Print report polish: day-cell table now uses `table-layout: fixed` with
  an explicit pixel width so all 7 day columns render at a consistent
  width regardless of content length (a long holiday name no longer
  stretches one column); the month-cycling text color (Black → Navy →
  Burgundy) that already existed in the app is now also applied in the
  print report, which previously always rendered plain black text.
- Product-maturity pass (2026-07-23, unsupervised — user asked to "think
  away" overnight, review pending):
  - **Naming proposal** — no git repo/package name exists yet. Recommended
    `loeledger` (see README.md for the full shortlist + collision check);
    not acted on beyond documenting the recommendation, since renaming/
    initializing git is a bigger, more deliberate step than a naming
    decision alone.
  - **Welcome dialog** (`WelcomeDialog`) — shown automatically on first
    launch (tracked via the new `app_state.json`, see DATA_MODEL.md),
    re-openable anytime via the "Welcome" button (top toolbar). Explains
    the forecast-first concept, day interactions, and key terms
    (Plan/Worked/Variance/Bank/Non-Billable) in plain language.
  - **Settings dialog** (`SettingsDialog`) — edits `settings.json` (start/
    end date, daily hours, half-day hours, start of week) via the
    "Settings" button instead of hand-editing JSON. Validates input
    (numeric fields, end ≥ start) before writing; requires an app restart
    for date-range changes to take effect (documented in the dialog
    itself) rather than attempting a live rebuild of the ledger.
  - **Contextual help ("?") buttons** — next to Bulk Edit, Forecast
    Summary, and LOE Tracking headers, plus a "Help ?" menu (top toolbar)
    covering the View bar, day-cell Worked-vs-Planned semantics, and
    Non-Billable. Each popup has a "What it does" and a "Why it works this
    way" section — the rationale is condensed from the actual decisions
    logged in this file, not just a restatement of the UI.

  All four were tested directly (dialog rendering, first-launch trigger,
  settings save + validation with a stubbed messagebox, help content
  rendering) before being left running for review. Nothing here was
  requested to be committed/pushed — it's local-only, same as everything
  else in this project so far.
- Timesheet extractor parser fix (2026-07-30): the July 26 timecard PDF
  had a session with a bare "Training" marker line inserted between "Net
  Time" and "Notes" (not seen in earlier PDFs), which broke `SESSION_RE`
  and silently dropped that one session — caught by `verify()`'s
  mismatch check (daily summary said 7.0h for 2026-07-27, detailed
  sessions only summed to 4.0h). Fixed by making the "Training" line
  optional in the regex; when present, the entry's description gets a
  `[Training]` prefix for visibility, but it's still imported as
  `billable: true` like everything else from an approved timesheet —
  training isn't automatically non-billable (that's a manual call; see
  the client-training correction earlier in this log). Regression-checked
  against the two earlier PDFs before re-running the import for real.
- Filled the missing Jul 20–24 week (a timecard PDF that hadn't existed
  yet) and added `import_timesheets.py --force`, which resolves the
  July 9-week gap noted below: it replaces a date's entries even if it
  already has some, but only when the new entries' total hours match the
  existing total exactly (else it's reported as a real conflict, not
  guessed at). Used it to backfill descriptions for 2026-07-03/06-10,
  which had bare `{"hours": N}` entries with no description predating
  the importer's existence.
- Akkodis monthly timesheet (`fill_akkodis_timesheet.py`, 2026-07-30):
  fills the client-provided `./Akkodis/*.xlsx` template purely from
  `work_log.json` via openpyxl — no AI interpretation of the source
  timesheets. Billable Hours = entries not marked non-billable; Vacation/
  Sick Leave = `daily_hours` on a would-be workday with that status (the
  real hours-off figure, not the app's internal `planned`, which is 0 for
  vacation for the app's own forecasting math — a different concern);
  Breakdown of Hours Worked = one line per logged entry (`<hours>h:
  <description>`), with `wrap_text` and row height scaled to the entry
  count so multi-entry days are actually readable, not run together on
  one line. Flags any `WORKING`-status weekday with zero logged entries
  as a likely missing timesheet before writing the file, rather than
  silently submitting a day blank. Columns A-D are top-aligned (matching
  a manual fix made to the July output) rather than centered, since
  taller multi-line Breakdown rows made center-aligned Date/Hours cells
  look vertically offset against the top-aligned text next to them.
- Project named and made public (2026-07-31): repo is `LOE-tracker`
  (git@github.com:bobbyc-brs/LOE-tracker.git), public. Real client/
  timesheet data stays out of git (`.gitignore`: `work_log.json`,
  `Timesheets/`, generated `Akkodis/*Timesheet-*.xlsx`, `data/`,
  `Screenshots/`) — code, docs, and non-identifying config are tracked.
  Docs reframed to be upfront that `timesheet_extractor.py` and
  `fill_akkodis_timesheet.py` are tied to one specific vendor's PDF/xlsx
  formats and are worked samples of the scan-and-ingest pattern, not
  universal importers — there wasn't a way to make the parsing itself
  format-agnostic, so the honest framing is in the docs instead.
