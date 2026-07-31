"""Imports timecard PDFs from ./Timesheets into work_log.json.

By default, only fills in entries for dates that don't already have any --
never overwrites existing entries, notes, or status. Safe to re-run any
time new PDFs are added to ./Timesheets; already-imported dates are
skipped.

--force additionally replaces a date's entries when it already has some,
but ONLY if the new entries' total hours match the existing total (within
0.01h) -- this is specifically for backfilling richer PDF data (real
descriptions, billable flags) over older bare {"hours": N} entries with no
description, such as ones typed in before this importer existed. If the
totals don't match, it's left alone and reported as a real conflict
instead of guessed at.
"""
import sys

import calendar_app as ca
from timesheet_extractor import TimesheetExtractor

TIMESHEETS_DIR = ca.BASE / 'Timesheets'


def import_all(dry_run=False, force=False):
    added, enriched, skipped, conflicts, mismatches = [], [], [], [], []

    for path in sorted(TIMESHEETS_DIR.glob('*.pdf')):
        ex = TimesheetExtractor(path)
        for iso_date, expected, actual in ex.verify():
            mismatches.append((path.name, iso_date, expected, actual))
        for iso_date, day_entries in ex.entries_with_fallback().items():
            existing = ca.WORKLOG.get(iso_date, {})
            existing_entries = existing.get('entries')
            if not existing_entries:
                rec = dict(existing)
                rec['entries'] = day_entries
                ca.WORKLOG[iso_date] = rec
                added.append(iso_date)
                continue

            if not force:
                skipped.append(iso_date)
                continue

            existing_total = sum(e.get('hours', 0) for e in existing_entries)
            new_total = sum(e.get('hours', 0) for e in day_entries)
            if abs(existing_total - new_total) > 0.01:
                conflicts.append((iso_date, existing_total, new_total))
                continue
            rec = dict(existing)
            rec['entries'] = day_entries
            ca.WORKLOG[iso_date] = rec
            enriched.append(iso_date)

    if mismatches:
        print('Mismatches (daily summary vs. detailed session total):')
        for name, iso_date, expected, actual in mismatches:
            print(f'  {name}: {iso_date} summary={expected} sessions_total={actual}')
        print()

    verb = 'Would add' if dry_run else 'Added'
    print(f'{verb} entries for {len(added)} date(s):', sorted(added))
    if force:
        verb2 = 'Would enrich' if dry_run else 'Enriched'
        print(f'{verb2} {len(enriched)} date(s) (hours matched, descriptions replaced):', sorted(enriched))
        if conflicts:
            print(f'CONFLICTS -- {len(conflicts)} date(s) left untouched, hours did not match:')
            for iso_date, existing_total, new_total in conflicts:
                print(f'  {iso_date}: existing={existing_total} new={new_total}')
    else:
        print(f'Skipped {len(skipped)} date(s) that already had entries (use --force to backfill matching-hours ones):', sorted(skipped))

    if dry_run:
        print('(dry run -- nothing saved)')
    elif added or enriched:
        ca.save()
        print('Saved to', ca.LOG_FILE)
    else:
        print('Nothing new to save.')


if __name__ == '__main__':
    import_all(dry_run='--dry-run' in sys.argv, force='--force' in sys.argv)
