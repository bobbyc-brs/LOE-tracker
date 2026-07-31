"""Canadian Work Hours Calendar.

Forecast-first weekly ledger. See VISION.md for the product intent.
"""
import csv
import html
import json
import tempfile
import webbrowser
from datetime import date, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).parent
SETTINGS_FILE = BASE / 'settings.json'
HOLIDAYS_FILE = BASE / 'holidays_on.json'
LOG_FILE = BASE / 'work_log.json'
APP_STATE_FILE = BASE / 'app_state.json'  # UI/session state, not user config -- e.g. "seen the intro screen"

CFG = json.load(open(SETTINGS_FILE))
HOLIDAYS = json.load(open(HOLIDAYS_FILE)) if HOLIDAYS_FILE.exists() else {}
WORKLOG = json.load(open(LOG_FILE)) if LOG_FILE.exists() else {}
APP_STATE = json.load(open(APP_STATE_FILE)) if APP_STATE_FILE.exists() else {}


def save_app_state():
    json.dump(APP_STATE, open(APP_STATE_FILE, 'w'), indent=2)

START = date.fromisoformat(CFG['start_date'])
END = date.fromisoformat(CFG['end_date'])
DAILY = CFG['daily_hours']
HALF_DAY_HOURS = CFG.get('half_day_hours', 4.0)  # default: not simply DAILY / 2 — see DATA_MODEL.md

STATUSES = ['WORKING', 'VACATION', 'HOLIDAY', 'HALF_DAY', 'PERSONAL_SICK', 'OFF']
COLORS = {
    'WORKING': '#d9f2d9',
    'VACATION': '#ead7ff',
    'HOLIDAY': '#d7ebff',
    'HALF_DAY': '#fff2b8',
    'PERSONAL_SICK': '#ffd9d9',
    'OFF': '#e0e0e0',
}
STATUS_LETTER = {'WORKING': 'W', 'VACATION': 'V', 'HOLIDAY': 'H', 'HALF_DAY': 'H/2', 'PERSONAL_SICK': 'P/S', 'OFF': 'O'}
LEAVE_STATUSES = ('VACATION', 'HOLIDAY', 'HALF_DAY', 'PERSONAL_SICK')
# Planned exceptions -- known in advance, so a manager already plans the LOE
# around them: fully credited, no variance impact. Sick leave is NOT here --
# it's unplanned, so it's meant to show up as a real variance instead.
PLANNED_LEAVE_STATUSES = ('VACATION', 'HOLIDAY', 'HALF_DAY')

MONTH_COLORS = ['#000000', '#000080', '#800020']  # Black -> Navy -> Burgundy, cycling every 3 months

WEEKDAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
SHORT_NAMES = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
START_OF_WEEK_IDX = WEEKDAY_NAMES.index(CFG.get('start_of_week', 'Saturday'))
WEEK_ORDER = [(START_OF_WEEK_IDX + i) % 7 for i in range(7)]

_HTML_PAGE = """<!doctype html>
<html><head><meta charset="utf-8">
<title>Canadian Work Hours Calendar</title>
<style>
  body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 24px; color: #111; }}
  h1 {{ font-size: 18px; margin-bottom: 4px; }}
  .period {{ color: #555; margin-bottom: 16px; }}
  .summary {{ white-space: pre-wrap; font-size: 13px; border: 1px solid #ccc; padding: 10px 14px;
              display: inline-block; margin-bottom: 20px; }}
  .week {{ margin-bottom: 14px; page-break-inside: avoid; }}
  .weekhdr {{ font-weight: bold; font-size: 12px; margin-bottom: 4px; }}
  .weekrow {{ display: flex; align-items: flex-start; gap: 16px; }}
  table.days {{ border-collapse: collapse; table-layout: fixed; width: 560px; }}
  table.days td {{ border: 1px solid #999; text-align: center; font-size: 11px;
                    padding: 4px 4px; vertical-align: top; width: 80px; overflow: hidden;
                    word-wrap: break-word; }}
  .dayname {{ font-weight: bold; }}
  .status {{ font-weight: bold; }}
  .holiday {{ font-size: 9px; }}
  .weektotals {{ font-size: 12px; font-weight: bold; white-space: nowrap; padding-top: 2px; }}
  .weektotals div {{ margin-bottom: 2px; }}
  @media print {{
    body {{ margin: 0.4in; }}
    .summary {{ page-break-inside: avoid; }}
  }}
</style></head>
<body>
  <h1>Canadian Work Hours Calendar</h1>
  <div class="period">{period}</div>
  <div class="summary">{summary}</div>
  {weeks}
</body></html>
"""


def save():
    json.dump(WORKLOG, open(LOG_FILE, 'w'), indent=2)


def is_scheduled_workday(ds, d):
    """Whether this date would be worked by default, ignoring any leave override."""
    return d.weekday() < 5 and ds not in HOLIDAYS


def planned_hours(status, ds, d):
    # A holiday or approved vacation is a day off, same as a weekend -- it
    # doesn't add to Plan, it's simply not a working day. (PERSONAL_SICK is
    # deliberately NOT here -- it stays fully planned below, so an unplanned
    # sick day still shows up as a real variance against that plan.)
    if status in ('OFF', 'HOLIDAY', 'VACATION'):
        return 0
    if status == 'WORKING':
        return DAILY
    # HALF_DAY / PERSONAL_SICK only cost planned hours on days that would
    # otherwise have been worked. Applying one on a day already off (a
    # weekend or a holiday) doesn't add hours to the plan.
    if not is_scheduled_workday(ds, d):
        return 0
    return HALF_DAY_HOURS if status == 'HALF_DAY' else DAILY


def effective_status(ds, d):
    explicit = WORKLOG.get(ds, {}).get('status')
    if explicit:
        return explicit
    if ds in HOLIDAYS:
        return 'HOLIDAY'
    return 'WORKING' if d.weekday() < 5 else 'OFF'


def day_hours(ds, d):
    """Returns (planned, credited, worked, status).

    'worked' is literal logged hours -- entries only, always, regardless of
    status or override. This is what "Worked" means everywhere it's shown
    (day cells, week/period summaries, exports): it never includes
    vacation/holiday/half-day/sick credit.

    'credited' is the policy-aware figure used for Variance/Bank: a planned
    exception (vacation/holiday/half-day) is fully credited with no
    logged entries needed, since it's known in advance and the LOE already
    accounts for it -- so it never shows as a shortfall. Sick leave is
    deliberately excluded from that fallback: it's unplanned, so with no
    entries logged it falls through to 0, same as an unlogged WORKING day,
    and shows up as a real variance.
    """
    rec = WORKLOG.get(ds, {})
    status = effective_status(ds, d)
    planned = planned_hours(status, ds, d)
    entries = rec.get('entries', [])
    worked = sum(e.get('hours', 0) for e in entries)

    override = rec.get('credited_override')
    if override is not None:
        credited = override
    elif entries:
        credited = worked
    elif status in PLANNED_LEAVE_STATUSES:
        credited = planned
    else:
        credited = 0
    return planned, credited, worked, status


def is_recorded(ds, status):
    """Whether this day has an actual timesheet behind it -- logged entries,
    a manual override, or an explicit leave status -- as opposed to just
    being an elapsed calendar date nobody has entered hours for yet."""
    rec = WORKLOG.get(ds, {})
    if rec.get('credited_override') is not None:
        return True
    if rec.get('entries'):
        return True
    return status in LEAVE_STATUSES


def non_billable_hours(ds):
    """Hours that don't count as billable client work for the week:
    Vacation/Sick leave, plus any work entry explicitly marked non-billable.
    Policy: anything that came from an approved timesheet is billable by
    default (existing entries have no 'billable' key and stay billable --
    this isn't retroactively reclassified); a manually-added entry defaults
    to non-billable unless its 'Billable' checkbox was checked."""
    d = date.fromisoformat(ds)
    status = effective_status(ds, d)
    total = 0.0
    if status in ('VACATION', 'PERSONAL_SICK'):
        # Use what a normal working day would have been -- not the day's own
        # planned value, which is 0 for vacation and doesn't reflect actual
        # non-billable time off. DAILY on a scheduled workday, else 0 (taking
        # vacation/sick on an already-off day doesn't add non-billable hours).
        total += DAILY if is_scheduled_workday(ds, d) else 0
    rec = WORKLOG.get(ds, {})
    total += sum(e.get('hours', 0) for e in rec.get('entries', []) if e.get('billable') is False)
    return total


def worked_line(ds, status, planned, worked):
    """Text for a day cell's hours line. A WORKING day with no timesheet
    behind it yet reads 'F:<planned>' (forecast) instead of 'W:0' -- showing
    what it's scheduled for, not a worked-hours count that hasn't happened.
    Otherwise it reads 'W:<worked>' -- literal logged hours only, never a
    vacation/holiday/sick credit (a vacation day reads 'W:0' unless you
    actually logged something on top of it)."""
    if status == 'WORKING' and not is_recorded(ds, status):
        return f'F:{planned:g}'
    return f'W:{worked:g}'


def loe_day_weight(status):
    """Level-of-Effort day weight: LOE is a day-rate count (scheduled/worked
    days), not an hours total. Vacation/holiday/sick days consume none of it
    -- they reduce how many days are available, not how many are 'used'."""
    if status == 'WORKING':
        return 1.0
    if status == 'HALF_DAY':
        return 0.5
    return 0.0


def compute_loe_days(view_start, view_end):
    """LOE days across a date range, using each day's effective_status --
    the real recorded status for past dates, the assumed/scheduled status
    for future dates -- so 'days committed so far' and 'days planned ahead'
    fall out of the same sum without needing to treat them separately."""
    total = 0.0
    d = view_start
    while d <= view_end:
        ds = d.isoformat()
        total += loe_day_weight(effective_status(ds, d))
        d += timedelta(days=1)
    return total


def week_start_on_or_before(d):
    while d.weekday() != START_OF_WEEK_IDX:
        d -= timedelta(days=1)
    return d


class DayEditor(tk.Toplevel):
    def __init__(self, master, ds, on_save):
        super().__init__(master)
        self.ds = ds
        self.on_save = on_save
        self.rec = WORKLOG.get(ds, {})
        self.entries = [dict(e) for e in self.rec.get('entries', [])]

        d = date.fromisoformat(ds)
        self.title(f'Edit {d.strftime("%A, %d %B %Y")}')
        self.resizable(False, False)
        self.grab_set()

        pad = {'padx': 8, 'pady': 4}

        top = ttk.Frame(self)
        top.pack(fill='x', **pad)
        ttk.Label(top, text='Status:').grid(row=0, column=0, sticky='w')
        self.status_var = tk.StringVar(value=effective_status(ds, d))
        ttk.Combobox(top, textvariable=self.status_var, values=STATUSES, state='readonly', width=12).grid(row=0, column=1, sticky='w')

        if ds in HOLIDAYS:
            ttk.Label(top, text=f'Holiday: {HOLIDAYS[ds]}', foreground='#3366cc').grid(row=0, column=2, sticky='w', padx=(12, 0))

        ttk.Label(top, text='Credited hours override:').grid(row=1, column=0, sticky='w', pady=(6, 0))
        self.override_var = tk.StringVar(value='' if self.rec.get('credited_override') is None else str(self.rec['credited_override']))
        ttk.Entry(top, textvariable=self.override_var, width=8).grid(row=1, column=1, sticky='w', pady=(6, 0))
        ttk.Label(top, text='(blank = auto-calculated)', foreground='#888888').grid(row=1, column=2, sticky='w', pady=(6, 0))

        ttk.Label(self, text='Work entries:').pack(anchor='w', padx=8)
        entry_frame = ttk.Frame(self)
        entry_frame.pack(fill='both', expand=True, padx=8)
        self.entry_tree = ttk.Treeview(entry_frame, columns=('hours', 'type', 'description'), show='headings', height=5)
        self.entry_tree.heading('hours', text='Hours')
        self.entry_tree.heading('type', text='Type')
        self.entry_tree.heading('description', text='Description')
        self.entry_tree.column('hours', width=60, anchor='center')
        self.entry_tree.column('type', width=90, anchor='center')
        self.entry_tree.column('description', width=230)
        self.entry_tree.pack(side='left', fill='both', expand=True)
        self.refresh_entries()

        add_frame = ttk.Frame(self)
        add_frame.pack(fill='x', padx=8, pady=(4, 0))
        ttk.Label(add_frame, text='Hours:').pack(side='left')
        self.new_hours_var = tk.StringVar()
        hours_entry = ttk.Entry(add_frame, textvariable=self.new_hours_var, width=6)
        hours_entry.pack(side='left', padx=(2, 8))
        ttk.Label(add_frame, text='Description:').pack(side='left')
        self.new_desc_var = tk.StringVar()
        desc_entry = ttk.Entry(add_frame, textvariable=self.new_desc_var, width=22)
        desc_entry.pack(side='left', padx=(2, 8))
        # Default unchecked: a manually-typed entry is non-billable unless
        # you say otherwise. Entries from an approved/imported timesheet are
        # billable by default instead -- see import_timesheets.py.
        self.new_billable_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(add_frame, text='Billable', variable=self.new_billable_var).pack(side='left', padx=(0, 8))
        hours_entry.bind('<Return>', lambda e: self.add_entry())
        desc_entry.bind('<Return>', lambda e: self.add_entry())
        ttk.Button(add_frame, text='Add Entry', command=self.add_entry).pack(side='left')
        ttk.Button(add_frame, text='Remove Selected', command=self.remove_entry).pack(side='left', padx=(4, 0))

        ttk.Label(self, text='Notes:').pack(anchor='w', padx=8, pady=(8, 0))
        self.notes_text = tk.Text(self, width=48, height=5)
        self.notes_text.pack(padx=8, pady=(0, 8))
        self.notes_text.insert('1.0', self.rec.get('notes', ''))

        btns = ttk.Frame(self)
        btns.pack(fill='x', padx=8, pady=(0, 8))
        ttk.Button(btns, text='Save', command=self.save_and_close).pack(side='right')
        ttk.Button(btns, text='Cancel', command=self.destroy).pack(side='right', padx=(0, 8))

    def refresh_entries(self):
        self.entry_tree.delete(*self.entry_tree.get_children())
        for i, e in enumerate(self.entries):
            entry_type = 'Non-billable' if e.get('billable') is False else 'Billable'
            self.entry_tree.insert('', 'end', iid=str(i), values=(e.get('hours', 0), entry_type, e.get('description', '')))

    def add_entry(self):
        try:
            hours = float(self.new_hours_var.get())
        except ValueError:
            messagebox.showerror('Invalid hours', 'Enter a numeric value for hours.')
            return
        entry = {'hours': hours, 'description': self.new_desc_var.get()}
        if not self.new_billable_var.get():
            entry['billable'] = False
        self.entries.append(entry)
        self.new_hours_var.set('')
        self.new_desc_var.set('')
        self.new_billable_var.set(False)
        self.refresh_entries()

    def remove_entry(self):
        sel = self.entry_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        del self.entries[idx]
        self.refresh_entries()

    def save_and_close(self):
        rec = {}
        rec['status'] = self.status_var.get()
        if self.entries:
            rec['entries'] = self.entries
        notes = self.notes_text.get('1.0', 'end').strip()
        if notes:
            rec['notes'] = notes
        override_raw = self.override_var.get().strip()
        if override_raw:
            try:
                rec['credited_override'] = float(override_raw)
            except ValueError:
                messagebox.showerror('Invalid override', 'Credited hours override must be numeric.')
                return

        if rec:
            WORKLOG[self.ds] = rec
        elif self.ds in WORKLOG:
            del WORKLOG[self.ds]
        save()
        self.on_save()
        self.destroy()


# -- Help content -------------------------------------------------------
# Each entry is (title, what_it_does, why_it_works_this_way). The "why" is
# deliberately kept close to the actual decisions made while building this
# (see PLAN.md for the full log) rather than just re-describing the UI.
HELP_TEXT = {
    'bulk_edit': (
        'Bulk Edit',
        "Select more than one day -- Ctrl-click to toggle individual days, "
        "or click-and-drag for a contiguous range -- then pick a status and "
        "hit Apply to set all of them at once.",
        "Marking a week of vacation used to mean clicking through each day "
        "individually, cycling its status every time. This panel only "
        "appears once 2+ days are selected -- for a single day, plain "
        "click/double-click is already faster, so showing it unconditionally "
        "was just noise.",
    ),
    'forecast_summary': (
        'Forecast Summary',
        "Answers the core question this app exists for: given your normal "
        "pattern, holidays, vacations, and actual hours, what does the rest "
        "of the period look like? \"Recorded\" covers days with a real "
        "timesheet behind them; \"Remainder\" looks ahead from today.",
        "A day only counts as \"Recorded\" if it actually has logged "
        "entries, an override, or an explicit leave status -- an elapsed "
        "calendar date with nothing logged yet isn't the same as a "
        "submitted timesheet, and treating it that way would show a fake "
        "deficit for a week nobody's entered hours for yet.",
    ),
    'loe': (
        'LOE Tracking',
        "LOE = Level of Effort: a day-rate budget assigned over a date "
        "range (e.g. \"90 days LOE Apr 1-Aug 31\"). Enter a target and it "
        "shows how many days (Working=1, Half-Day=0.5) fall in whatever "
        "period is currently selected in the View bar, versus that target.",
        "Raised by a manager email about an assignment tracking against its "
        "LOE. There's no separate stored list of LOE periods -- it always "
        "covers whatever's selected in the View bar, so checking a "
        "different LOE window is just a matter of changing the view.",
    ),
    'worked_vs_credited': (
        'Day Cells: Worked vs. Planned',
        "\"W:\" is literal logged hours only -- it never includes vacation/"
        "holiday/half-day/sick credit. A vacation day shows W:0 unless you "
        "actually logged real hours on top of it. A WORKING day with "
        "nothing logged yet shows \"F:<planned>\" (forecast) instead of "
        "W:0, since it hasn't happened yet, it's just scheduled.",
        "Originally \"Worked\" quietly included leave credit, which made a "
        "vacation day look like a full day of actual work. Vacation and "
        "Holiday still count as zero variance internally (they're planned "
        "exceptions a manager's LOE already accounts for), but Sick leave "
        "doesn't get that free pass -- it's unplanned, so it shows a real "
        "variance if nothing's logged, the same as an unlogged working day.",
    ),
    'non_billable': (
        'Non-Billable',
        "Tracks hours that aren't billable client work: Vacation/Sick days "
        "count in full, plus any work entry you mark non-billable "
        "(training, admin, travel time). Shown per week only when nonzero.",
        "Policy: anything from an approved/imported timesheet is billable "
        "by default. A manually-typed entry defaults to non-billable "
        "instead -- check its \"Billable\" box if it really is billable.",
    ),
    'view_bar': (
        'View / Period Picker',
        "Scopes which weeks are shown below, and what the Forecast Summary "
        "covers -- pick a preset or a custom range. Export CSV / Export "
        "Excel / Print all export whatever range is currently selected "
        "here.",
        "The per-week Bank running total is deliberately NOT rescoped by "
        "this -- it always accumulates from the true start date, since Bank "
        "answers \"am I ahead or behind overall,\" not just within "
        "whatever window happens to be on screen right now.",
    ),
}


def show_help(parent, key):
    title, what_text, why_text = HELP_TEXT[key]
    win = tk.Toplevel(parent)
    win.title(f'Help: {title}')
    win.resizable(False, False)
    win.grab_set()
    body = tk.Frame(win, padx=14, pady=12)
    body.pack()
    tk.Label(body, text=title, font=('Segoe UI', 13, 'bold')).pack(anchor='w', pady=(0, 8))
    tk.Label(body, text='What it does', font=('Segoe UI', 9, 'bold'), fg='#444444').pack(anchor='w')
    tk.Label(body, text=what_text, wraplength=380, justify='left', font=('Segoe UI', 10)).pack(anchor='w', pady=(0, 10))
    tk.Label(body, text='Why it works this way', font=('Segoe UI', 9, 'bold'), fg='#444444').pack(anchor='w')
    tk.Label(body, text=why_text, wraplength=380, justify='left', font=('Segoe UI', 10), fg='#333333').pack(anchor='w', pady=(0, 10))
    ttk.Button(body, text='Close', command=win.destroy).pack(anchor='e')


def help_button(parent, key):
    """A small '?' button that opens the help popup for `key` when clicked."""
    return tk.Button(parent, text='?', font=('Segoe UI', 8, 'bold'), width=2,
                      command=lambda: show_help(parent, key), cursor='hand2',
                      relief='solid', bd=1)


def show_help_menu(parent, button, keys):
    """A '?' button that pops up a small menu of help topics (for a spot
    with more than one relevant topic and no single natural header)."""
    menu = tk.Menu(parent, tearoff=0)
    for key in keys:
        menu.add_command(label=HELP_TEXT[key][0], command=lambda k=key: show_help(parent, k))
    x = button.winfo_rootx()
    y = button.winfo_rooty() + button.winfo_height()
    menu.tk_popup(x, y)


WELCOME_TEXT = (
    "This is a forecast-first work-hours ledger, not a timesheet system. "
    "It's built to answer: how many hours am I expected to work for the "
    "rest of the period? Am I ahead or behind? How much of my assigned "
    "LOE (Level of Effort) have I used?\n\n"
    "Working with a day:\n"
    "   - Single-click a day to cycle its status (Working -> Vacation -> "
    "Holiday -> Half-Day -> Personal/Sick -> Off).\n"
    "   - Click the \"W:\" line, or double-click anywhere on the day, to "
    "open the detailed editor -- add work entries, notes, or override "
    "credited hours.\n"
    "   - Ctrl-click, or click-and-drag across days, to select a range, "
    "then use Bulk Edit (right panel) to set them all at once.\n\n"
    "Key ideas:\n"
    "   - Plan = hours you're expected to work.\n"
    "   - Worked = actual logged hours only -- never vacation/holiday/sick "
    "credit.\n"
    "   - Variance / Bank = how you're tracking against plan. Vacation and "
    "Holiday are planned exceptions and never count against you; Sick "
    "leave is unplanned, so it does.\n"
    "   - Non-Billable = training/admin/vacation/sick time, tracked "
    "separately from billable client work.\n\n"
    "Use the View bar (top-left) to scope what you're looking at, and to "
    "export to CSV / Excel / Print. Most sections have a small \"?\" "
    "button for more detail on how they work and why."
)


class WelcomeDialog(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title('Welcome')
        self.resizable(False, False)
        self.grab_set()

        body = tk.Frame(self, padx=16, pady=14)
        body.pack()
        tk.Label(body, text='Welcome to the Canadian Work Hours Calendar',
                 font=('Segoe UI', 14, 'bold')).pack(anchor='w', pady=(0, 10))
        tk.Label(body, text=WELCOME_TEXT, wraplength=460, justify='left', font=('Segoe UI', 10)).pack(anchor='w')

        self.dont_show_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(body, text="Don't show this again", variable=self.dont_show_var).pack(anchor='w', pady=(12, 8))
        ttk.Button(body, text='Got it', command=self._close).pack(anchor='e')

    def _close(self):
        if self.dont_show_var.get():
            APP_STATE['seen_intro'] = True
            save_app_state()
        self.destroy()


class SettingsDialog(tk.Toplevel):
    """Edits settings.json. Date-range/hours changes need an app restart to
    take effect, since the ledger's week widgets are only built once at
    startup (see App.build()) -- rebuilding them live isn't worth the risk
    of a config screen that's used rarely."""

    def __init__(self, master):
        super().__init__(master)
        self.title('Settings')
        self.resizable(False, False)
        self.grab_set()

        body = tk.Frame(self, padx=14, pady=12)
        body.pack()
        tk.Label(body, text='Configuration', font=('Segoe UI', 13, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))

        tk.Label(body, text='Start date (YYYY-MM-DD):').grid(row=1, column=0, sticky='w', pady=3)
        self.start_var = tk.StringVar(value=CFG['start_date'])
        ttk.Entry(body, textvariable=self.start_var, width=14).grid(row=1, column=1, sticky='w')

        tk.Label(body, text='End date (YYYY-MM-DD):').grid(row=2, column=0, sticky='w', pady=3)
        self.end_var = tk.StringVar(value=CFG['end_date'])
        ttk.Entry(body, textvariable=self.end_var, width=14).grid(row=2, column=1, sticky='w')

        tk.Label(body, text='Daily hours:').grid(row=3, column=0, sticky='w', pady=3)
        self.daily_var = tk.StringVar(value=str(CFG['daily_hours']))
        ttk.Entry(body, textvariable=self.daily_var, width=14).grid(row=3, column=1, sticky='w')

        tk.Label(body, text='Half-day hours:').grid(row=4, column=0, sticky='w', pady=3)
        self.half_day_var = tk.StringVar(value=str(CFG.get('half_day_hours', 4.0)))
        ttk.Entry(body, textvariable=self.half_day_var, width=14).grid(row=4, column=1, sticky='w')

        tk.Label(body, text='Start of week:').grid(row=5, column=0, sticky='w', pady=3)
        self.start_of_week_var = tk.StringVar(value=CFG.get('start_of_week', 'Saturday'))
        ttk.Combobox(body, textvariable=self.start_of_week_var, values=WEEKDAY_NAMES, state='readonly', width=12).grid(row=5, column=1, sticky='w')

        tk.Label(body, text='Changes here need an app restart to take effect.',
                 font=('Segoe UI', 8), fg='#888888', wraplength=280, justify='left').grid(row=6, column=0, columnspan=2, sticky='w', pady=(10, 8))

        btns = tk.Frame(body)
        btns.grid(row=7, column=0, columnspan=2, sticky='e')
        ttk.Button(btns, text='Save', command=self._save).pack(side='right')
        ttk.Button(btns, text='Cancel', command=self.destroy).pack(side='right', padx=(0, 8))

    def _save(self):
        try:
            start = date.fromisoformat(self.start_var.get().strip())
            end = date.fromisoformat(self.end_var.get().strip())
            daily = float(self.daily_var.get().strip())
            half_day = float(self.half_day_var.get().strip())
        except ValueError:
            messagebox.showerror('Invalid value', 'Check the date and number fields.', parent=self)
            return
        if end < start:
            messagebox.showerror('Invalid range', 'End date must be on or after start date.', parent=self)
            return
        new_cfg = dict(CFG)
        new_cfg['start_date'] = start.isoformat()
        new_cfg['end_date'] = end.isoformat()
        new_cfg['daily_hours'] = daily
        new_cfg['half_day_hours'] = half_day
        new_cfg['start_of_week'] = self.start_of_week_var.get()
        json.dump(new_cfg, open(SETTINGS_FILE, 'w'), indent=2)
        messagebox.showinfo('Settings saved', 'Saved. Restart the app for changes to take effect.', parent=self)
        self.destroy()


class App:
    def __init__(self, root):
        self.root = root
        root.title('Canadian Work Hours Calendar')
        WIDTH, HEIGHT = 1150, 850
        root.geometry(f'{WIDTH}x{HEIGHT}')

        main = tk.PanedWindow(root, orient='horizontal', sashwidth=6)
        main.pack(fill='both', expand=True)

        left = tk.Frame(main)
        right = tk.Frame(main, width=300)
        main.add(left, stretch='always')
        main.add(right, minsize=280, stretch='never')
        self._main_paned = main
        self._sash_pos = int(WIDTH * 0.68)

        # Bulk Edit section only shows once a multi-day selection exists --
        # a single day is already handled by plain click/double-click, so
        # showing this all the time is just noise.
        self.bulk_edit_frame = tk.Frame(right)
        bulk_hdr = tk.Frame(self.bulk_edit_frame)
        bulk_hdr.pack(anchor='w', padx=10, pady=(8, 2), fill='x')
        tk.Label(bulk_hdr, text='Bulk Edit', font=('Segoe UI', 12, 'bold')).pack(side='left')
        help_button(bulk_hdr, 'bulk_edit').pack(side='left', padx=(6, 0))
        self.selection_info = tk.StringVar()
        tk.Label(self.bulk_edit_frame, textvariable=self.selection_info, font=('Segoe UI', 9), fg='#555555',
                 wraplength=260, justify='left').pack(anchor='w', padx=10)
        bulk_row = tk.Frame(self.bulk_edit_frame)
        bulk_row.pack(anchor='w', padx=10, pady=(4, 10))
        self.bulk_status_var = tk.StringVar(value=STATUSES[0])
        ttk.Combobox(bulk_row, textvariable=self.bulk_status_var, values=STATUSES, state='readonly', width=13).pack(side='left')
        ttk.Button(bulk_row, text='Apply', command=self.apply_bulk_status).pack(side='left', padx=(4, 0))
        ttk.Button(bulk_row, text='Clear', command=self.clear_selection).pack(side='left', padx=(4, 0))
        # Not packed here -- _update_selection_panel() packs/unpacks it based
        # on how many days are currently selected (ctrl-click multiple, or drag).
        # Re-inserted with pack(before=self.forecast_header) each time it's
        # shown so hiding/reshowing never reorders the sidebar.

        self.forecast_header = tk.Frame(right)
        self.forecast_header.pack(anchor='w', padx=10, pady=(8, 2), fill='x')
        tk.Label(self.forecast_header, text='Forecast Summary', font=('Segoe UI', 14, 'bold')).pack(side='left')
        help_button(self.forecast_header, 'forecast_summary').pack(side='left', padx=(6, 0))
        self.summary = tk.StringVar()
        tk.Label(right, textvariable=self.summary, justify='left', font=('Segoe UI', 10)).pack(anchor='w', padx=10, pady=0)

        loe_hdr = tk.Frame(right)
        loe_hdr.pack(anchor='w', padx=10, pady=(10, 2), fill='x')
        tk.Label(loe_hdr, text='LOE Tracking', font=('Segoe UI', 14, 'bold')).pack(side='left')
        help_button(loe_hdr, 'loe').pack(side='left', padx=(6, 0))
        tk.Label(right, text="Type your assigned day target for whatever period is currently"
                              " selected in the View bar (top-left) -- the days-committed"
                              " figure below it is calculated, not editable.",
                 font=('Segoe UI', 8), fg='#888888', wraplength=260, justify='left').pack(anchor='w', padx=10)
        loe_row = tk.Frame(right)
        loe_row.pack(anchor='w', padx=10, pady=(4, 0))
        tk.Label(loe_row, text='Target (days) — you enter this:').pack(side='left')
        self.loe_target_var = tk.StringVar()
        loe_entry = ttk.Entry(loe_row, textvariable=self.loe_target_var, width=8)
        loe_entry.pack(side='left', padx=(4, 0))
        loe_entry.bind('<KeyRelease>', lambda e: self._update_loe_display())
        tk.Label(right, text='Calculated:', font=('Segoe UI', 8, 'bold'), fg='#666666').pack(anchor='w', padx=10, pady=(6, 0))
        self.loe_result = tk.StringVar()
        self.loe_result_label = tk.Label(right, textvariable=self.loe_result, justify='left', font=('Segoe UI', 10))
        self.loe_result_label.pack(anchor='w', padx=10, pady=(4, 0))

        view_bar = tk.Frame(left)
        view_bar.pack(fill='x', padx=8, pady=(8, 4))
        tk.Label(view_bar, text='View:', font=('Segoe UI', 9, 'bold')).pack(side='left')
        self.view_start = START
        self.view_end = END
        self.view_preset_var = tk.StringVar(value='Full Period')
        preset_combo = ttk.Combobox(
            view_bar, textvariable=self.view_preset_var, state='readonly', width=16,
            values=['Full Period', 'This Month', 'This Quarter', 'Year to Date', 'Remainder of Year', 'Custom'],
        )
        preset_combo.pack(side='left', padx=(4, 10))
        preset_combo.bind('<<ComboboxSelected>>', self._on_preset_selected)
        tk.Label(view_bar, text='From:').pack(side='left')
        self.view_start_var = tk.StringVar(value=START.isoformat())
        ttk.Entry(view_bar, textvariable=self.view_start_var, width=11).pack(side='left', padx=(2, 8))
        tk.Label(view_bar, text='To:').pack(side='left')
        self.view_end_var = tk.StringVar(value=END.isoformat())
        ttk.Entry(view_bar, textvariable=self.view_end_var, width=11).pack(side='left', padx=(2, 8))
        ttk.Button(view_bar, text='Go', command=self._apply_custom_view).pack(side='left')
        ttk.Button(view_bar, text='Export CSV', command=self.export_csv).pack(side='left', padx=(10, 0))
        ttk.Button(view_bar, text='Export Excel', command=self.export_xlsx).pack(side='left', padx=(4, 0))
        ttk.Button(view_bar, text='Print / PDF', command=self.print_report).pack(side='left', padx=(4, 0))
        self.view_status = tk.StringVar()
        tk.Label(view_bar, textvariable=self.view_status, fg='#555555', font=('Segoe UI', 8)).pack(side='left', padx=(10, 0))

        toolbar2 = tk.Frame(left)
        toolbar2.pack(fill='x', padx=8, pady=(0, 4))
        ttk.Button(toolbar2, text='Settings', command=lambda: SettingsDialog(self.root)).pack(side='left')
        ttk.Button(toolbar2, text='Welcome', command=lambda: WelcomeDialog(self.root)).pack(side='left', padx=(4, 0))
        help_btn = ttk.Button(toolbar2, text='Help ?')
        help_btn.configure(command=lambda: show_help_menu(toolbar2, help_btn, ['view_bar', 'worked_vs_credited', 'non_billable']))
        help_btn.pack(side='left', padx=(4, 0))

        self.canvas = tk.Canvas(left)
        sb = ttk.Scrollbar(left, orient='vertical', command=self.canvas.yview)
        self.frame = tk.Frame(self.canvas)
        self.frame.columnconfigure(0, weight=1)
        self.frame.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.create_window((0, 0), window=self.frame, anchor='nw')
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        self.canvas.bind_all('<MouseWheel>', lambda e: self.canvas.yview_scroll(int(-e.delta / 120), 'units'))

        self.build()
        self._set_view(START, END)

        # Placed last, after all left- and right-pane content already exists --
        # setting this earlier gets silently overridden once later-packed
        # content forces the PanedWindow to recompute the sash itself.
        root.update_idletasks()
        self._main_paned.sash_place(0, self._sash_pos, 0)

        if not APP_STATE.get('seen_intro'):
            root.after(150, lambda: WelcomeDialog(root))

    def build(self):
        """One-time widget construction. Never called again after __init__ —
        subsequent changes go through recompute(), which mutates the existing
        widgets in place so there's no rebuild flicker and no scroll reset."""
        self.cells = {}
        self.weeks = []
        self.widget_to_ds = {}
        self.selection = set()
        self._drag_anchor = None
        self._drag_active = False
        self._all_days = [(START + timedelta(days=i)).isoformat() for i in range((END - START).days + 1)]
        cur = week_start_on_or_before(START)
        week_idx = 0

        while cur <= END:
            card = tk.LabelFrame(self.frame, text=f'Week of {cur.isoformat()}')
            # grid (not pack) so hiding/showing weeks for the period picker
            # (via grid_remove/grid) preserves chronological order.
            card.grid(row=week_idx, column=0, sticky='ew', padx=8, pady=6)

            week_days = []
            for col, wd in enumerate(WEEK_ORDER):
                d = cur + timedelta(days=(wd - cur.weekday()) % 7)
                if d < START or d > END:
                    continue
                ds = d.isoformat()
                week_days.append(ds)

                cell = tk.Frame(card, bd=1, relief='solid')
                cell.grid(row=0, column=col, padx=2, pady=2, sticky='n')
                name_lbl = tk.Label(cell, width=9, font=('Segoe UI', 9, 'bold'))
                name_lbl.pack()
                planned_lbl = tk.Label(cell)
                planned_lbl.pack()
                worked_lbl = tk.Label(cell)
                worked_lbl.pack()
                status_lbl = tk.Label(cell, font=('Segoe UI', 9, 'bold'))
                status_lbl.pack()
                holiday_lbl = tk.Label(cell, font=('Segoe UI', 7), wraplength=80)
                holiday_lbl.pack()

                for w in (cell, name_lbl, planned_lbl, status_lbl, holiday_lbl):
                    self.widget_to_ds[w] = ds
                    w.bind('<ButtonPress-1>', lambda e, dd=ds: self._on_press(dd))
                    w.bind('<B1-Motion>', self._on_drag)
                    w.bind('<ButtonRelease-1>', lambda e, dd=ds: self._on_release(dd))
                    w.bind('<Control-Button-1>', lambda e, dd=ds: self._on_ctrl_click(dd))
                    w.bind('<Double-Button-1>', lambda e, dd=ds: self.open_editor(dd))

                # The "W:" (worked hours) label is a dedicated shortcut straight
                # to the entry editor, instead of participating in status-cycle
                # / selection like the rest of the cell.
                self.widget_to_ds[worked_lbl] = ds
                worked_lbl.configure(cursor='hand2')
                worked_lbl.bind('<Button-1>', lambda e, dd=ds: self.open_editor(dd))
                worked_lbl.bind('<Double-Button-1>', lambda e, dd=ds: self.open_editor(dd))

                self.cells[ds] = {
                    'frame': cell, 'name': name_lbl, 'planned': planned_lbl,
                    'worked': worked_lbl, 'status': status_lbl, 'holiday': holiday_lbl,
                }

            tk.Frame(card, bg='#999999', width=1).grid(row=0, column=7, sticky='ns', padx=6)

            summary = tk.Frame(card)
            summary.grid(row=0, column=8, sticky='n', padx=(4, 8), pady=2)
            plan_val = tk.Label(summary, font=('Segoe UI', 9, 'bold'))
            plan_val.pack(anchor='w')
            worked_val = tk.Label(summary, font=('Segoe UI', 9, 'bold'))
            worked_val.pack(anchor='w')
            variance_val = tk.Label(summary, font=('Segoe UI', 9, 'bold'))
            variance_val.pack(anchor='w')
            bank_val = tk.Label(summary, font=('Segoe UI', 9, 'bold'))
            bank_val.pack(anchor='w')
            # Not packed here -- only shown for weeks with nonzero non-billable
            # hours, appended last so plain pack()/pack_forget() (no sibling
            # ever gets packed after it) keeps ordering stable.
            nonbillable_val = tk.Label(summary, font=('Segoe UI', 9, 'bold'), fg='#8a5a00')

            self.weeks.append({
                'days': week_days, 'card': card,
                'start': date.fromisoformat(week_days[0]), 'end': date.fromisoformat(week_days[-1]),
                'plan_val': plan_val, 'worked_val': worked_val,
                'variance_val': variance_val, 'bank_val': bank_val,
                'nonbillable_val': nonbillable_val,
            })
            cur += timedelta(days=7)
            week_idx += 1

        self.recompute()

    def recompute(self):
        """Update all cell/summary text and colors in place (no widget
        destruction), then recompute the right-hand Forecast Summary."""
        today = date.today()
        bank = 0.0

        for week in self.weeks:
            wp = ww = week_nonbillable = 0.0
            recorded_variance = 0.0
            has_recorded_day = False
            for ds in week['days']:
                d = date.fromisoformat(ds)
                planned, credited, worked, status = day_hours(ds, d)
                wp += planned
                ww += worked
                week_nonbillable += non_billable_hours(ds)
                if is_recorded(ds, status):
                    recorded_variance += credited - planned
                    has_recorded_day = True
                self._update_cell(ds, d, planned, worked, status, today)

            bank_color = '#1a7a1a' if bank >= 0 else '#b02020'
            week['plan_val'].configure(text=f'Plan: {wp:.1f}')
            week['worked_val'].configure(text=f'Worked: {ww:.1f}')
            if has_recorded_day:
                bank += recorded_variance
                variance_color = '#1a7a1a' if recorded_variance >= 0 else '#b02020'
                week['variance_val'].configure(text=f'Variance: {recorded_variance:+.1f}', fg=variance_color)
            else:
                week['variance_val'].configure(text='Variance: — (no timesheet yet)', fg='#888888')
            week['bank_val'].configure(text=f'Bank: {bank:+.1f}', fg=bank_color)

            if week_nonbillable > 0:
                week['nonbillable_val'].configure(text=f'Non-Billable: {week_nonbillable:.1f}')
                if not week['nonbillable_val'].winfo_ismapped():
                    week['nonbillable_val'].pack(anchor='w')
            elif week['nonbillable_val'].winfo_ismapped():
                week['nonbillable_val'].pack_forget()

        self.summary.set(self.compute_summary())
        self._update_loe_display()
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def _update_loe_display(self):
        committed = compute_loe_days(self.view_start, self.view_end)
        target_raw = self.loe_target_var.get().strip()
        if not target_raw:
            self.loe_result_label.configure(fg='black')
            self.loe_result.set(f'Days (Working + Half-Day) in this period: {committed:g}')
            return
        try:
            target = float(target_raw)
        except ValueError:
            self.loe_result_label.configure(fg='black')
            self.loe_result.set('Enter a numeric day target.')
            return
        delta = committed - target
        label = 'Surplus' if delta >= 0 else 'Shortfall'
        self.loe_result_label.configure(fg='#1a7a1a' if delta >= 0 else '#b02020')
        self.loe_result.set(
            f'Days in this period: {committed:g}\n'
            f'LOE Target: {target:g}\n'
            f'{label}: {delta:+.1f} days'
        )

    def _update_cell(self, ds, d, planned, worked, status, today):
        c = self.cells[ds]
        bg = COLORS[status]
        fg = MONTH_COLORS[(d.month - 1) % 3]

        name = f'{SHORT_NAMES[d.weekday()]}({d.day})'
        if d == today:
            name += ' *'

        holiday_text = HOLIDAYS.get(ds, '') if (ds in HOLIDAYS and status == 'HOLIDAY') else ''

        c['frame'].configure(bg=bg)
        c['name'].configure(text=name, bg=bg, fg=fg)
        c['planned'].configure(text=f'P:{planned:g}', bg=bg, fg=fg)
        c['worked'].configure(text=worked_line(ds, status, planned, worked), bg=bg, fg=fg)
        c['status'].configure(text=STATUS_LETTER[status], bg=bg, fg=fg)
        c['holiday'].configure(text=holiday_text, bg=bg, fg=fg)

    def compute_summary(self):
        today = date.today()

        plan_recorded = worked_recorded = credited_recorded = 0.0
        remaining_plan = 0.0
        remaining_work_days = 0
        vacation_days = 0
        holiday_days = 0
        personal_sick_days = 0

        d = self.view_start
        while d <= self.view_end:
            ds = d.isoformat()
            planned, credited, worked, status = day_hours(ds, d)

            if is_recorded(ds, status):
                plan_recorded += planned
                worked_recorded += worked
                credited_recorded += credited

            if d >= today:
                remaining_plan += planned
                if status == 'WORKING':
                    remaining_work_days += 1
                elif status == 'VACATION':
                    vacation_days += 1
                elif status == 'HOLIDAY':
                    holiday_days += 1
                elif status == 'PERSONAL_SICK':
                    personal_sick_days += 1
            d += timedelta(days=1)

        variance_recorded = credited_recorded - plan_recorded
        projected_bank = variance_recorded

        lines = [
            f'Period: {self.view_start.isoformat()} to {self.view_end.isoformat()}',
            f'Today: {today.isoformat()}',
            '-- Recorded --',
            f'Planned: {plan_recorded:.1f}',
            f'Worked: {worked_recorded:.1f}',
            f'Variance: {variance_recorded:+.1f}',
            '-- Remainder of Period --',
            f'Remaining Planned Hours: {remaining_plan:.1f}',
            f'Remaining Work Days: {remaining_work_days}',
            f'Vacation Days: {vacation_days}',
            f'Holiday Days: {holiday_days}',
            f'Personal/Sick Days: {personal_sick_days}',
            f'Projected Bank Balance: {projected_bank:+.1f}',
        ]
        return '\n'.join(lines)

    def change_status(self, ds):
        d = date.fromisoformat(ds)
        cur_status = effective_status(ds, d)
        idx = (STATUSES.index(cur_status) + 1) % len(STATUSES)
        WORKLOG.setdefault(ds, {})['status'] = STATUSES[idx]
        save()
        self.recompute()

    def open_editor(self, ds):
        DayEditor(self.root, ds, self.recompute)

    # -- CSV export -------------------------------------------------------
    # Exports whatever range is currently selected in the period picker, so
    # picking "This Month" (etc.) and then exporting gives you just that range.

    def _build_csv_rows(self):
        header = ['Date', 'Day', 'Status', 'Planned', 'Worked', 'Non-Billable', 'Variance', 'Entries', 'Notes']
        rows = []
        total_plan = total_worked = total_nonbillable = 0.0
        total_recorded_variance = 0.0
        total_has_recorded = False
        week_plan = week_worked = week_nonbillable = 0.0
        week_recorded_variance = 0.0
        week_has_recorded = False
        week_end_weekday = (START_OF_WEEK_IDX - 1) % 7

        d = self.view_start
        while d <= self.view_end:
            ds = d.isoformat()
            planned, credited, worked, status = day_hours(ds, d)
            rec = WORKLOG.get(ds, {})

            entries_desc = '; '.join(
                f"{e.get('hours', 0):g}h" + (f" ({e['description']})" if e.get('description') else '')
                for e in rec.get('entries', [])
            )
            notes = rec.get('notes', '')
            if status == 'HOLIDAY' and ds in HOLIDAYS:
                notes = HOLIDAYS[ds] + (f'; {notes}' if notes else '')

            day_nonbillable = non_billable_hours(ds)

            # Variance only counts days with an actual timesheet behind them,
            # same rule as the on-screen ledger and the print report -- a day
            # nobody's logged yet isn't "behind," it just hasn't happened.
            # It's computed from credited (policy-aware: vacation/holiday
            # nets to zero, sick doesn't), not from the Worked column itself.
            recorded = is_recorded(ds, status)
            day_variance = f'{credited - planned:+.1f}' if recorded else ''
            rows.append([ds, SHORT_NAMES[d.weekday()], status, f'{planned:g}', f'{worked:g}',
                         f'{day_nonbillable:g}' if day_nonbillable else '', day_variance, entries_desc, notes])

            week_plan += planned
            week_worked += worked
            week_nonbillable += day_nonbillable
            total_plan += planned
            total_worked += worked
            total_nonbillable += day_nonbillable
            if recorded:
                week_recorded_variance += credited - planned
                week_has_recorded = True
                total_recorded_variance += credited - planned
                total_has_recorded = True

            if d.weekday() == week_end_weekday or d == self.view_end:
                week_variance = f'{week_recorded_variance:+.1f}' if week_has_recorded else ''
                rows.append(['Week Total', '', '', f'{week_plan:g}', f'{week_worked:g}',
                             f'{week_nonbillable:g}' if week_nonbillable else '', week_variance, '', ''])
                rows.append([])
                week_plan = week_worked = week_nonbillable = 0.0
                week_recorded_variance = 0.0
                week_has_recorded = False

            d += timedelta(days=1)

        total_variance = f'{total_recorded_variance:+.1f}' if total_has_recorded else ''
        rows.append(['Period Total', '', '', f'{total_plan:g}', f'{total_worked:g}',
                     f'{total_nonbillable:g}' if total_nonbillable else '', total_variance, '', ''])
        return header, rows

    def _write_csv(self, path):
        header, rows = self._build_csv_rows()
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow([f'Canadian Work Hours Calendar: {self.view_start.isoformat()} to {self.view_end.isoformat()}'])
            writer.writerow([])
            writer.writerow(header)
            writer.writerows(rows)

    def export_csv(self):
        default_name = f'workhours_{self.view_start.isoformat()}_to_{self.view_end.isoformat()}.csv'
        path = filedialog.asksaveasfilename(
            parent=self.root, defaultextension='.csv', filetypes=[('CSV files', '*.csv'), ('All files', '*.*')],
            initialfile=default_name, title='Export to CSV',
        )
        if not path:
            return
        self._write_csv(path)
        messagebox.showinfo('Export complete', f'Saved to:\n{path}', parent=self.root)

    # -- Excel export (live formulas) --------------------------------------
    # Unlike CSV, .xlsx can hold real formulas -- weekly totals are SUM()
    # over that week's day columns, and the Period Total sums the weekly
    # totals, so editing a raw hour value in Excel recalculates everything
    # above it.

    def _write_xlsx(self, path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Work Hours'
        bold = Font(bold=True)

        ws.cell(row=1, column=1, value=f'Canadian Work Hours Calendar: {self.view_start.isoformat()} to {self.view_end.isoformat()}').font = Font(bold=True, size=13)
        ws.column_dimensions['A'].width = 20
        for col in 'BCDEFGH':
            ws.column_dimensions[col].width = 9
        ws.column_dimensions['J'].width = 10

        row = 3
        plan_total_cells = []
        worked_total_cells = []
        credited_total_cells = []
        nonbillable_total_cells = []
        cur = week_start_on_or_before(START)

        while cur <= END:
            days_in_week = []
            for i in range(7):
                d = cur + timedelta(days=i)
                if d < START or d > END:
                    continue
                ds = d.isoformat()
                planned, credited, worked, status = day_hours(ds, d)
                days_in_week.append((d, ds, planned, credited, worked, status))

            week_start_d, week_end_d = days_in_week[0][0], days_in_week[-1][0]
            if not (week_end_d >= self.view_start and week_start_d <= self.view_end):
                cur += timedelta(days=7)
                continue

            # Credited/Non-Billable are always included (as 0 where they
            # don't apply) rather than conditionally shown like the
            # app/print/CSV -- a fixed row count per week keeps formula
            # row-offsets predictable. Credited feeds the Variance formula
            # (policy-aware: vacation/holiday nets to zero, sick doesn't) --
            # Worked itself is literal logged hours only, so Variance can't
            # be derived from it directly the way it could before.
            header_row, planned_row, worked_row, credited_row, nonbillable_row, status_row = row, row + 1, row + 2, row + 3, row + 4, row + 5
            ws.cell(row=header_row, column=1, value=f'Week of {week_start_d.isoformat()}').font = bold
            ws.cell(row=header_row, column=10, value='Total').font = bold
            ws.cell(row=planned_row, column=1, value='Planned')
            ws.cell(row=worked_row, column=1, value='Worked')
            ws.cell(row=credited_row, column=1, value='Credited')
            ws.cell(row=nonbillable_row, column=1, value='Non-Billable')
            ws.cell(row=status_row, column=1, value='Status')

            for i, (d, ds, planned, credited, worked, status) in enumerate(days_in_week):
                col = 2 + i  # B=2 .. H=8
                fill = PatternFill('solid', fgColor=COLORS[status].lstrip('#'))
                header_cell = ws.cell(row=header_row, column=col, value=f'{SHORT_NAMES[d.weekday()]}({d.day})')
                header_cell.font = bold
                planned_cell = ws.cell(row=planned_row, column=col, value=planned)
                worked_cell = ws.cell(row=worked_row, column=col, value=worked)
                credited_cell = ws.cell(row=credited_row, column=col, value=credited)
                nonbillable_cell = ws.cell(row=nonbillable_row, column=col, value=non_billable_hours(ds))
                status_cell = ws.cell(row=status_row, column=col, value=STATUS_LETTER[status])
                for c in (header_cell, planned_cell, worked_cell, credited_cell, nonbillable_cell, status_cell):
                    c.fill = fill

            last_col = get_column_letter(1 + len(days_in_week))
            plan_total_cell = f'J{planned_row}'
            worked_total_cell = f'J{worked_row}'
            credited_total_cell = f'J{credited_row}'
            nonbillable_total_cell = f'J{nonbillable_row}'
            ws.cell(row=planned_row, column=10, value=f'=SUM(B{planned_row}:{last_col}{planned_row})')
            ws.cell(row=worked_row, column=10, value=f'=SUM(B{worked_row}:{last_col}{worked_row})')
            ws.cell(row=credited_row, column=10, value=f'=SUM(B{credited_row}:{last_col}{credited_row})')
            ws.cell(row=nonbillable_row, column=10, value=f'=SUM(B{nonbillable_row}:{last_col}{nonbillable_row})')
            ws.cell(row=status_row, column=10, value='Variance:')
            ws.cell(row=status_row, column=11, value=f'={credited_total_cell}-{plan_total_cell}')

            plan_total_cells.append(plan_total_cell)
            worked_total_cells.append(worked_total_cell)
            credited_total_cells.append(credited_total_cell)
            nonbillable_total_cells.append(nonbillable_total_cell)
            row = status_row + 2  # blank separator row
            cur += timedelta(days=7)

        ws.cell(row=row, column=1, value='Period Total').font = bold
        ws.cell(row=row, column=2, value='Planned')
        ws.cell(row=row, column=10, value='=SUM(' + ','.join(plan_total_cells) + ')' if plan_total_cells else 0)
        ws.cell(row=row + 1, column=2, value='Worked')
        ws.cell(row=row + 1, column=10, value='=SUM(' + ','.join(worked_total_cells) + ')' if worked_total_cells else 0)
        ws.cell(row=row + 2, column=2, value='Credited')
        ws.cell(row=row + 2, column=10, value='=SUM(' + ','.join(credited_total_cells) + ')' if credited_total_cells else 0)
        ws.cell(row=row + 3, column=2, value='Non-Billable')
        ws.cell(row=row + 3, column=10, value='=SUM(' + ','.join(nonbillable_total_cells) + ')' if nonbillable_total_cells else 0)
        ws.cell(row=row + 3, column=11, value='Variance:')
        ws.cell(row=row + 3, column=12, value=f'=J{row + 2}-J{row}')

        wb.save(path)

    def export_xlsx(self):
        default_name = f'workhours_{self.view_start.isoformat()}_to_{self.view_end.isoformat()}.xlsx'
        path = filedialog.asksaveasfilename(
            parent=self.root, defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx'), ('All files', '*.*')],
            initialfile=default_name, title='Export to Excel',
        )
        if not path:
            return
        self._write_xlsx(path)
        messagebox.showinfo('Export complete', f'Saved to:\n{path}', parent=self.root)

    # -- Print / PDF -------------------------------------------------------
    # Renders whatever range is currently selected in the period picker as a
    # standalone HTML page styled like the app's own color-coded weekly
    # cards, and opens it in the default browser -- from there the browser's
    # own Print / "Save as PDF" handles pagination, so no PDF library needed.
    # Bank is still cumulative from the true start_date even though only the
    # selected weeks are rendered, matching what's shown on screen.

    def _build_html_report(self):
        cur = week_start_on_or_before(START)
        bank = 0.0
        week_blocks = []

        while cur <= END:
            days_in_week = []
            wp = ww = week_nonbillable = 0.0
            recorded_variance = 0.0
            has_recorded_day = False
            for i in range(7):
                d = cur + timedelta(days=i)
                if d < START or d > END:
                    continue
                ds = d.isoformat()
                planned, credited, worked, status = day_hours(ds, d)
                wp += planned
                ww += worked
                week_nonbillable += non_billable_hours(ds)
                if is_recorded(ds, status):
                    recorded_variance += credited - planned
                    has_recorded_day = True
                days_in_week.append((d, ds, planned, worked, status))

            if has_recorded_day:
                bank += recorded_variance

            week_start, week_end = days_in_week[0][0], days_in_week[-1][0]
            if week_end >= self.view_start and week_start <= self.view_end:
                variance_text = f'{recorded_variance:+.1f}' if has_recorded_day else '&mdash;'
                week_blocks.append(self._render_week_html(days_in_week, wp, ww, variance_text, bank, week_nonbillable))

            cur += timedelta(days=7)

        return _HTML_PAGE.format(
            period=html.escape(f'{self.view_start.isoformat()} to {self.view_end.isoformat()}'),
            summary=html.escape(self.summary.get()),
            weeks='\n'.join(week_blocks),
        )

    def _render_week_html(self, days_in_week, wp, ww, variance_text, bank, week_nonbillable=0.0):
        cells = []
        for d, ds, planned, worked, status in days_in_week:
            name = f'{SHORT_NAMES[d.weekday()]}({d.day})'
            fg = MONTH_COLORS[(d.month - 1) % 3]
            holiday_note = f'<div class="holiday">{html.escape(HOLIDAYS[ds])}</div>' if (ds in HOLIDAYS and status == 'HOLIDAY') else ''
            cells.append(
                f'<td style="background:{COLORS[status]};color:{fg}">'
                f'<div class="dayname">{name}</div>'
                f'<div>P:{planned:g}</div><div>{worked_line(ds, status, planned, worked)}</div>'
                f'<div class="status">{STATUS_LETTER[status]}</div>{holiday_note}</td>'
            )
        nonbillable_line = f'<div>Non-Billable: {week_nonbillable:.1f}</div>' if week_nonbillable > 0 else ''
        return (
            f'<div class="week">'
            f'<div class="weekhdr">Week of {days_in_week[0][0].isoformat()}</div>'
            f'<div class="weekrow">'
            f'<table class="days"><tr>{"".join(cells)}</tr></table>'
            f'<div class="weektotals">'
            f'<div>Plan: {wp:.1f}</div><div>Worked: {ww:.1f}</div>'
            f'<div>Variance: {variance_text}</div><div>Bank: {bank:+.1f}</div>{nonbillable_line}'
            f'</div>'
            f'</div>'
            f'</div>'
        )

    def print_report(self):
        html_text = self._build_html_report()
        path = Path(tempfile.gettempdir()) / f'workhours_{self.view_start.isoformat()}_to_{self.view_end.isoformat()}.html'
        path.write_text(html_text, encoding='utf-8')
        webbrowser.open(path.as_uri())

    # -- Period picker ---------------------------------------------------
    # This only scopes which weeks are shown in the scrollable ledger. The
    # Forecast Summary and per-week Bank running total are unaffected -- they
    # always cover the full settings.json start_date..end_date period.

    def _on_preset_selected(self, event=None):
        preset = self.view_preset_var.get()
        if preset == 'Custom':
            return
        today = date.today()
        if preset == 'Full Period':
            vs, ve = START, END
        elif preset == 'This Month':
            first = today.replace(day=1)
            next_month = date(today.year + (today.month == 12), today.month % 12 + 1, 1)
            vs, ve = first, next_month - timedelta(days=1)
        elif preset == 'This Quarter':
            q_start_month = (today.month - 1) // 3 * 3 + 1
            first = today.replace(month=q_start_month, day=1)
            next_q_year = today.year + (q_start_month + 3 > 12)
            next_q_month = (q_start_month + 3 - 1) % 12 + 1
            vs, ve = first, date(next_q_year, next_q_month, 1) - timedelta(days=1)
        elif preset == 'Year to Date':
            vs, ve = date(today.year, 1, 1), today
        else:  # Remainder of Year
            vs, ve = today, date(today.year, 12, 31)
        self._set_view(vs, ve)

    def _apply_custom_view(self):
        try:
            vs = date.fromisoformat(self.view_start_var.get().strip())
            ve = date.fromisoformat(self.view_end_var.get().strip())
        except ValueError:
            messagebox.showerror('Invalid date', 'Enter dates as YYYY-MM-DD.')
            return
        self.view_preset_var.set('Custom')
        self._set_view(vs, ve)

    def _set_view(self, view_start, view_end):
        view_start = max(START, min(view_start, END))
        view_end = max(START, min(view_end, END))
        if view_start > view_end:
            view_start, view_end = view_end, view_start
        self.view_start, self.view_end = view_start, view_end
        self.view_start_var.set(view_start.isoformat())
        self.view_end_var.set(view_end.isoformat())

        for week in self.weeks:
            if week['end'] >= view_start and week['start'] <= view_end:
                week['card'].grid()
            else:
                week['card'].grid_remove()

        self.recompute()  # Forecast Summary is scoped to view_start/view_end
        self.frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))
        self.canvas.yview_moveto(0)
        self.view_status.set(f'Showing {view_start.isoformat()} to {view_end.isoformat()}')

    # -- Multi-day selection -------------------------------------------------
    # Plain click-and-release with no movement: single-day status cycle (as before).
    # Click-and-drag: selects the contiguous date range between press and release.
    # Ctrl-click: toggles one day in/out of the selection, for non-contiguous picks.

    def _on_press(self, ds):
        self._drag_anchor = ds
        self._drag_active = False

    def _on_drag(self, event):
        if self._drag_anchor is None:
            return
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        target_ds = self.widget_to_ds.get(widget)
        if target_ds is None:
            return
        if target_ds != self._drag_anchor:
            self._drag_active = True
        if self._drag_active:
            self._select_range(self._drag_anchor, target_ds)

    def _on_release(self, ds):
        anchor = self._drag_anchor
        was_dragging = self._drag_active
        self._drag_anchor = None
        self._drag_active = False
        if anchor is None:
            # This release wasn't preceded by a plain press we tracked (e.g.
            # it followed a Ctrl-click, which is handled by _on_ctrl_click).
            return
        if was_dragging:
            return
        self.clear_selection()
        self.change_status(anchor)

    def _on_ctrl_click(self, ds):
        if ds in self.selection:
            self.selection.discard(ds)
        else:
            self.selection.add(ds)
        self._update_selection_visuals()
        self._update_selection_panel()

    def _select_range(self, ds_a, ds_b):
        ia, ib = self._all_days.index(ds_a), self._all_days.index(ds_b)
        lo, hi = min(ia, ib), max(ia, ib)
        self.selection = set(self._all_days[lo:hi + 1])
        self._update_selection_visuals()
        self._update_selection_panel()

    def clear_selection(self):
        self.selection = set()
        self._update_selection_visuals()
        self._update_selection_panel()

    def _update_selection_visuals(self):
        for ds, c in self.cells.items():
            if ds in self.selection:
                c['frame'].configure(highlightthickness=3, highlightbackground='#1a56db', highlightcolor='#1a56db')
            else:
                c['frame'].configure(highlightthickness=0)

    def _update_selection_panel(self):
        n = len(self.selection)
        if n > 1:
            self.selection_info.set(f'{n} days selected. Ctrl-click or drag to adjust the selection.')
            if not self.bulk_edit_frame.winfo_ismapped():
                self.bulk_edit_frame.pack(anchor='w', fill='x', before=self.forecast_header)
        elif self.bulk_edit_frame.winfo_ismapped():
            self.bulk_edit_frame.pack_forget()

    def apply_bulk_status(self):
        if not self.selection:
            return
        status = self.bulk_status_var.get()
        for ds in self.selection:
            WORKLOG.setdefault(ds, {})['status'] = status
        save()
        self.clear_selection()
        self.recompute()


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
