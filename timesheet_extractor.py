"""Parses 'Time and Expense Community' timecard PDFs (TCH-*.pdf) into
per-date work entries compatible with work_log.json's schema.

This is tied to one specific vendor's PDF export format -- there's no way
around that, since the regexes below match its exact field layout. Treat
it as a worked sample of "how to scan and ingest a timecard PDF into
work_log.json," not a universal importer. Adapting it to a different
timesheet system means writing new regexes against that system's actual
PDF text layout (print/export one and inspect it with pypdf, the same way
this one was built), not just pointing this module at a different file.

These are text-layer PDFs (not scans), so plain extraction via pypdf is
sufficient -- no OCR/AI needed. That part of the approach does generalize.
"""
import re
from datetime import date, timedelta
from pathlib import Path

import pypdf

WEEKDAY_ORDER = ['Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

DATE_RANGE_RE = re.compile(r'Start Date (\d{2}/\d{2}/\d{4}) End Date (\d{2}/\d{2}/\d{4})')
TIMECARD_ID_RE = re.compile(r'Timecard Id (\S+)')
DAILY_HOURS_RE = re.compile(r'(' + '|'.join(WEEKDAY_ORDER) + r') Hours (\d+\.\d+)')
SESSION_RE = re.compile(
    r'TMS-\d+\s*'
    r'Start Date Time(\d{2}/\d{2}/\d{4}) (\d{1,2}:\d{2} [AP]M)\s*'
    r'End Date Time(\d{2}/\d{2}/\d{4}) (\d{1,2}:\d{2} [AP]M)\s*'
    r'Net Time(\d+\.\d+)\s*'
    r'(Training\s*)?'  # some sessions have a bare "Training" marker line before Notes
    r'Notes(.*?)(?=TMS-\d+|\Z)',
    re.DOTALL,
)


def _to_iso(ddmmyyyy):
    return date(int(ddmmyyyy[6:10]), int(ddmmyyyy[3:5]), int(ddmmyyyy[0:2])).isoformat()


class TimesheetExtractor:
    """Extracts one timecard PDF's data.

    entries() -- {iso_date: [{'hours': float, 'description': str}, ...]},
    built from the detailed Timer Session Events (has notes).

    daily_totals() -- {iso_date: total_hours}, built from the weekday
    summary line ("Wednesday Hours 8.00" etc). Used only to cross-check
    entries() via verify() -- it has no notes of its own.
    """

    def __init__(self, path):
        self.path = Path(path)
        self.text = self._extract_text(self.path)

    @staticmethod
    def _extract_text(path):
        reader = pypdf.PdfReader(str(path))
        return '\n'.join(page.extract_text() for page in reader.pages)

    def timecard_id(self):
        m = TIMECARD_ID_RE.search(self.text)
        return m.group(1) if m else None

    def date_range(self):
        """(start_iso, end_iso) of the timecard period, or (None, None)."""
        m = DATE_RANGE_RE.search(self.text)
        if not m:
            return None, None
        return _to_iso(m.group(1)), _to_iso(m.group(2))

    def daily_totals(self):
        start_iso, _ = self.date_range()
        if not start_iso:
            return {}
        start = date.fromisoformat(start_iso)
        totals_by_name = dict(DAILY_HOURS_RE.findall(self.text))
        result = {}
        for i, name in enumerate(WEEKDAY_ORDER):
            hours = float(totals_by_name.get(name, 0))
            if hours:
                result[(start + timedelta(days=i)).isoformat()] = hours
        return result

    def entries(self):
        # Marked billable=True explicitly: anything in an approved/submitted
        # timesheet is billable by policy, as opposed to a manually-typed
        # entry in the app, which defaults to non-billable instead.
        # A session flagged "Training" in the source is noted in the
        # description (so it's visible for you to review) but NOT
        # auto-marked non-billable -- training time isn't automatically
        # non-billable (e.g. billable client-funded training already came
        # up once), so that classification stays a manual call.
        result = {}
        for start_date_str, _start_time, _end_date_str, _end_time, net_time_str, training_flag, notes in SESSION_RE.findall(self.text):
            iso_date = _to_iso(start_date_str)
            description = ' '.join(line.strip() for line in notes.strip().splitlines() if line.strip())
            if training_flag:
                description = f'[Training] {description}'.strip()
            result.setdefault(iso_date, []).append({'hours': float(net_time_str), 'description': description, 'billable': True})
        return result

    def entries_with_fallback(self):
        """Like entries(), but for a date that has a nonzero daily total and
        no detailed session data (some timecards only have the summary line,
        no Timer Session Events at all), falls back to a single lump entry
        labeled 'Imported timecard' -- matching how such days are already
        recorded elsewhere in this project's work_log.json."""
        result = {k: list(v) for k, v in self.entries().items()}
        for iso_date, total in self.daily_totals().items():
            if iso_date not in result:
                result[iso_date] = [{'hours': total, 'description': 'Imported timecard', 'billable': True}]
        return result

    def verify(self):
        """Cross-checks detailed session entries against the daily summary
        line. Returns a list of (date, summary_total, entries_total) for
        any date where they disagree -- empty list means it's all consistent."""
        totals = self.daily_totals()
        entries = self.entries()
        mismatches = []
        all_dates = set(totals) | set(entries)
        for iso_date in sorted(all_dates):
            expected = totals.get(iso_date, 0.0)
            actual = sum(e['hours'] for e in entries.get(iso_date, []))
            if abs(actual - expected) > 0.01:
                mismatches.append((iso_date, expected, actual))
        return mismatches


if __name__ == '__main__':
    import sys
    for path in sys.argv[1:]:
        ex = TimesheetExtractor(path)
        start, end = ex.date_range()
        print(f'{path}\n  Timecard: {ex.timecard_id()}  Period: {start} to {end}')
        for iso_date, day_entries in sorted(ex.entries().items()):
            for e in day_entries:
                print(f'    {iso_date}: {e["hours"]:g}h  {e["description"]}')
        mismatches = ex.verify()
        if mismatches:
            print('  MISMATCHES vs daily summary:', mismatches)
        print()
