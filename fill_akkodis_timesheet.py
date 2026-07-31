"""Fills the Akkodis monthly timesheet template from work_log.json.

All data comes from work_log.json (this project's system of record) via
plain Python/openpyxl -- no AI interpretation of timesheets involved.

Column mapping (Akkodis template has no separate Holiday column):
  Billable Hours -- sum of that day's entries not explicitly marked
                    non-billable (billable=False).
  Vacation       -- daily_hours if the day is VACATION and would otherwise
                    have been a workday, else blank. (Not the day's
                    internal `planned` value, which calendar_app.py treats
                    as 0 for vacation for its own forecasting purposes --
                    this sheet needs the real hours-off figure instead.)
  Sick Leave     -- same logic as Vacation, for PERSONAL_SICK.
  Breakdown      -- "<hours>h: <description>" per entry, one entry per
                    line (n entries -> n lines) so each hour range gets
                    its own remark rather than being run together.
                    work_log.json doesn't store clock-in/out times (only
                    hours + description), so this won't look like the
                    template's "8:00-12:00pm" example -- it's what's
                    actually in the system of record.
A day with none of the above (weekend, holiday, or genuinely no data) is
left blank across B/C/D/E, not zero-filled -- this is a document going to
an external party, and the existing SUM() formulas treat blank the same
as 0 either way.
"""
import calendar
import sys
from datetime import date

import openpyxl
from openpyxl.styles import Alignment

import calendar_app as ca

AKKODIS_DIR = ca.BASE / 'Akkodis'
FIRST_DATA_ROW = 8  # row for day 1
BASE_ROW_HEIGHT = 15  # openpyxl/Excel default


def find_template():
    """Locates the client-provided template in ./Akkodis by discovery
    rather than a hardcoded filename -- that file (whatever it's actually
    named) stays local-only (see .gitignore); this tracked script
    shouldn't assume or embed one particular client's exact filename."""
    candidates = [p for p in AKKODIS_DIR.glob('*.xlsx') if 'template' in p.stem.lower()]
    if not candidates:
        raise FileNotFoundError(
            f"No template found in {AKKODIS_DIR} (expected a .xlsx with 'template' in its filename)."
        )
    if len(candidates) > 1:
        raise FileNotFoundError(
            f"Multiple template candidates in {AKKODIS_DIR}: {[p.name for p in candidates]} "
            "-- pass template_path explicitly to fill_month()."
        )
    return candidates[0]


def day_breakdown_lines(rec):
    """One line per entry (n entries -> n remark lines), not run together."""
    lines = []
    for e in rec.get('entries', []):
        hours = e.get('hours', 0)
        desc = e.get('description', '')
        lines.append(f'{hours:g}h: {desc}' if desc else f'{hours:g}h')
    return lines


def missing_workdays(year, month):
    """Weekdays in the month with an effective WORKING status but no
    logged entries at all -- likely a missing timesheet, worth flagging."""
    num_days = calendar.monthrange(year, month)[1]
    missing = []
    for day in range(1, num_days + 1):
        d = date(year, month, day)
        ds = d.isoformat()
        status = ca.effective_status(ds, d)
        rec = ca.WORKLOG.get(ds, {})
        if status == 'WORKING' and not rec.get('entries'):
            missing.append(ds)
    return missing


def fill_month(year, month, template_path=None):
    wb = openpyxl.load_workbook(template_path or find_template())
    ws = wb['Timesheet']

    ws['B5'] = date(year, month, 1)

    num_days = calendar.monthrange(year, month)[1]
    for day in range(1, 32):
        row = FIRST_DATA_ROW + day - 1
        if day > num_days:
            for col in 'ABCDE':
                ws[f'{col}{row}'] = None
            continue

        d = date(year, month, day)
        ds = d.isoformat()
        rec = ca.WORKLOG.get(ds, {})
        status = ca.effective_status(ds, d)

        billable_hours = sum(
            e.get('hours', 0) for e in rec.get('entries', []) if e.get('billable') is not False
        )
        vacation_hours = ca.DAILY if (status == 'VACATION' and ca.is_scheduled_workday(ds, d)) else 0
        sick_hours = ca.DAILY if (status == 'PERSONAL_SICK' and ca.is_scheduled_workday(ds, d)) else 0
        lines = day_breakdown_lines(rec)

        ws[f'A{row}'] = day
        ws[f'B{row}'] = billable_hours or None
        ws[f'C{row}'] = vacation_hours or None
        ws[f'D{row}'] = sick_hours or None
        e_cell = ws[f'E{row}']
        e_cell.value = '\n'.join(lines) or None
        e_cell.alignment = Alignment(wrap_text=True, vertical='top')
        # Multi-line Breakdown rows are taller and top-aligned -- keep A-D
        # top-aligned too (horizontal centering stays, matching the
        # template default) so every cell in the row starts at the same
        # vertical position instead of A-D looking vertically centered
        # against a taller top-aligned E.
        for col in 'ABCD':
            cell = ws[f'{col}{row}']
            cell.alignment = Alignment(horizontal='center', vertical='top')
        if len(lines) > 1:
            ws.row_dimensions[row].height = BASE_ROW_HEIGHT * len(lines)

    out_name = f'Bobby Chawla- Akkodis Monthly Timesheet-{calendar.month_name[month]}{year}.xlsx'
    out_path = ca.BASE / 'Akkodis' / out_name
    wb.save(out_path)
    return out_path


if __name__ == '__main__':
    year = int(sys.argv[1]) if len(sys.argv) > 1 else date.today().year
    month = int(sys.argv[2]) if len(sys.argv) > 2 else date.today().month

    gaps = missing_workdays(year, month)
    if gaps:
        print(f'WARNING: {len(gaps)} weekday(s) with no logged entries at all: {gaps}')
        print('These will show as blank Billable Hours below -- likely a missing timesheet import.')
        print()

    path = fill_month(year, month)
    print('Saved to', path)
