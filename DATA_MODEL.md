# Data Model

All data lives in JSON files at the project root. There is no database and
no external service — `calendar_app.py` reads them on startup and writes
back after every edit. `settings.json` / `holidays_on.json` / `work_log.json`
are user data; `app_state.json` (below) is UI/session state, not config a
user would normally hand-edit.

## settings.json

```json
{
  "start_date": "2026-07-01",
  "end_date": "2026-12-31",
  "daily_hours": 7.5,
  "start_of_week": "Saturday"
}
```

Editable via the in-app **Settings** button (top toolbar) as of 2026-07-23,
as well as by hand. Date-range/hours changes need an app restart to take
effect — the ledger's week widgets are only built once at startup, and
rebuilding them live wasn't judged worth the risk for a screen used rarely.

- `start_date` / `end_date` — ISO dates bounding the forecast period.
- `daily_hours` — the standard full-day planned hours (e.g. 7.5).
- `start_of_week` — one of the weekday names (`"Monday"`...`"Sunday"`); controls
  the column order of each weekly card.
- `half_day_hours` (optional) — planned hours for a `HALF_DAY`. **Defaults to
  `4.0`** if not present. Note this is a fixed default, not `daily_hours / 2`
  — a half day is treated as a policy value in its own right, not simply
  half of the full-day figure.

## app_state.json

```json
{
  "seen_intro": true
}
```

Session/UI state, separate from `settings.json` because it isn't something
a user would go edit deliberately — currently just whether the Welcome
dialog has been dismissed with "Don't show this again." Absent entirely
until the first time it's written (e.g. a fresh checkout always shows the
Welcome dialog once).

## holidays_on.json

Maps an ISO date to a holiday name:

```json
{
  "2026-08-03": "Civic Holiday",
  "2026-09-07": "Labour Day"
}
```

Any date present here defaults to status `HOLIDAY` unless `work_log.json`
records an explicit status for that date (see below) — an explicit status
always wins.

## work_log.json

Maps an ISO date to a day record. A date with no entry at all uses the
default schedule: `WORKING` on weekdays, `OFF` on weekends, `HOLIDAY` if the
date appears in `holidays_on.json`.

```json
{
  "2026-07-07": {
    "entries": [
      {"hours": 3.5},
      {"hours": 5.5, "description": "Client call"}
    ]
  },
  "2026-11-20": {
    "status": "VACATION"
  },
  "2026-12-24": {
    "status": "HALF_DAY",
    "notes": "Office closes at noon",
    "credited_override": 4
  }
}
```

Every field on a day record is optional; only the fields that differ from
the computed default need to be present.

- `status` (string, optional) — one of `WORKING`, `VACATION`, `HOLIDAY`,
  `HALF_DAY`, `PERSONAL_SICK`, `OFF`. Set by single-clicking a day (cycles
  through this list) or via the day editor. If absent, the effective status
  is derived from `holidays_on.json` and the weekday, as above.
  Display-only nuance: a `WORKING` day is shown as `F` (forecast) instead of
  `W` until it's actually recorded (see `is_recorded` below) — this doesn't
  change the stored status, only how it's labelled/colored on screen, in the
  print report, and in exports.
- `entries` (list of objects, optional) — logged work sessions for the day,
  each `{"hours": <number>, "description": <string, optional>, "billable":
  <bool, optional>}`. A day can have any number of entries; they're summed
  to get actual worked hours. Added/removed via the day editor (double-click
  a day).

  **Billable policy**: anything from an approved/submitted timesheet is
  billable by default — `import_timesheets.py` / `TimesheetExtractor` set
  `"billable": true` explicitly on every entry it creates. A manually-typed
  entry added directly in the day editor defaults to **non-billable**
  instead (the "Billable" checkbox starts unchecked) — check it if the
  manual entry really is billable work. An entry with no `billable` key at
  all (i.e. pre-existing data from before this field existed) is treated as
  billable — that's never retroactively reclassified.

  A week's non-billable total — the sum of any non-billable entries plus
  the full credited hours of any `VACATION` or `PERSONAL_SICK` day that
  week — is shown as an extra "Non-Billable: X" line in that week's summary
  (on screen, in the print report, and as a CSV column) only when it's
  nonzero.
- `notes` (string, optional) — free-text notes for the day, edited in the
  day editor.
- `credited_override` (number, optional) — manually overrides the hours
  credited toward the plan for that day, taking precedence over both logged
  entries and the leave-type default. Used when policy credits a day
  differently than the automatic calculation (see below).

### Planned vs. credited hours

For a given day's effective status, **planned hours** are:

| Status           | Planned hours                                                            |
|-------------------|---------------------------------------------------------------------------|
| `WORKING`          | `daily_hours`                                                           |
| `HOLIDAY`          | `0`                                                                     |
| `VACATION`         | `0`                                                                     |
| `PERSONAL_SICK`    | `daily_hours` **if the date would otherwise be worked, else `0`**      |
| `HALF_DAY`         | `half_day_hours` (default `4.0`) **if the date would otherwise be worked, else `0`** |
| `OFF`              | `0`                                                                     |

A holiday or approved vacation is treated like a weekend: it's simply a
day off, not a day that was "planned then perfectly offset." Neither adds
to Plan, and neither counts as worked hours — both are `0`. `PERSONAL_SICK`
is deliberately different: it stays fully planned (see below for why —
it's the one status meant to show a real variance when unplanned).

"Would otherwise be worked" means the date is a weekday (Monday–Friday)
**and** isn't itself a designated holiday. `PERSONAL_SICK` / `HALF_DAY`
only add planned hours under that condition — applying one of these
statuses to a day that's already off (a weekend, or a holiday) is a no-op
for hours: it doesn't add to Plan and isn't credited as worked, since no
work was ever expected that day. `WORKING` is the one exception:
explicitly marking a weekend or holiday as `WORKING` means you really are
planning to work it, so it always counts as a full planned day.

There are two different hours figures per day, and they're deliberately
not the same thing:

**Worked hours** (shown as "Worked" / `W:` everywhere — day cells, week and
period summaries, CSV/print/Excel) is `entries[].hours` summed, full stop.
It never includes vacation/holiday/half-day/sick credit, and it's not
affected by `credited_override`. A vacation day always shows `W:0` unless
you actually logged real hours on top of it. (A `WORKING` day with nothing
logged yet shows `F:<planned>` — forecast — instead of `W:0`; see
`worked_line()`.)

**Credited hours** (used only for Variance/Bank, never displayed as
"Worked") are computed in this priority order:

1. `credited_override`, if set.
2. Sum of `entries[].hours`, if any entries are logged.
3. For **planned-exception** statuses (`VACATION`, `HOLIDAY`, `HALF_DAY`)
   with no logged entries, credited hours equal planned hours — these are
   known about in advance, so a manager's LOE already accounts for them,
   and they're credited in full rather than counting as a shortfall.
4. Otherwise (`WORKING`, `OFF`, or **`PERSONAL_SICK`** with no logged
   entries), credited hours are `0`. Sick leave is deliberately excluded
   from the planned-exception fallback above: it's unplanned, so with
   nothing logged it falls straight through to `0` and shows up as a real
   variance for that day, the same as an unlogged `WORKING` day would.

Variance is `credited - planned` for a day, but **only days with an actual
timesheet behind them count toward Variance and Bank.** A day counts as
"recorded" if it has logged `entries`, a `credited_override`, or an
explicit leave status (`VACATION`, `HOLIDAY`, `HALF_DAY`, `PERSONAL_SICK`
— all four, including sick, since marking a day sick is itself a
deliberate record even though it isn't auto-credited) — i.e. there's a
real, deliberate record for that date. A plain `WORKING` day with nothing
logged does **not** count, even if the date has already passed — the
calendar moving forward isn't the same thing as a timesheet being
submitted, and treating it as such would show a negative Variance for a
week simply because nobody has entered their hours yet.

A week's Variance is the sum of `credited - planned` over its recorded
days only (shown as `— (no timesheet yet)` if the week has no recorded
days at all); Bank is the running cumulative of those recorded-only
variances from `start_date`. The Forecast Summary's "-- Recorded --"
section and "Projected Bank Balance" follow the same rule, summed across
the whole period — on the assumption that any day without a timesheet yet
will eventually go exactly to plan. Note the "-- Recorded --" section's
displayed "Worked" figure is the pure Worked-hours total, but its
"Variance" is still computed from Credited — so for a period containing
vacation/holiday, the two won't arithmetically reconcile via simple
subtraction, and that's intentional (see Excel export note below).

The Excel export makes both figures visible explicitly (Planned / Worked /
Credited / Non-Billable / Status rows per week) rather than hiding
Credited, specifically so its live Variance formula can reference the
Credited total and stay accurate rather than silently diverging from the
app.
