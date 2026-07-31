import json
from datetime import datetime,timedelta,date
import tkinter as tk
from tkinter import ttk
from pathlib import Path

BASE=Path(__file__).parent
settings=json.load(open(BASE/'data/settings.json'))
work=json.load(open(BASE/'data/work_log.json'))

start=datetime.fromisoformat(settings['start_date']).date()
end=datetime.fromisoformat(settings['end_date']).date()

def day_hours(d):
    rec=work.get(d.isoformat(), {})
    status=rec.get('status','WORKING' if d.weekday()<5 else 'OFF')
    scheduled={'WORKING':7.5,'VACATION':7.5,'SICK':7.5,'STAT_HOLIDAY':7.5,'HALF_DAY':4.0,'OFF':0}.get(status,7.5)
    actual=sum(e.get('hours',0) for e in rec.get('entries',[]))
    if status in ('VACATION','SICK','STAT_HOLIDAY','HALF_DAY') and actual==0:
        actual=scheduled
    return scheduled,actual,status

root=tk.Tk()
root.title('Canadian Work Hours Calendar v0.6')
root.geometry('1300x700')

left=ttk.Frame(root); left.pack(side='left',fill='both',expand=True)
right=ttk.Frame(root,width=250); right.pack(side='right',fill='y')

cols=['week','sat','sun','mon','tue','wed','thu','fri','plan','worked','var']
tree=ttk.Treeview(left,columns=cols,show='headings')
for c in cols: tree.heading(c,text=c.upper())
ys=ttk.Scrollbar(left,orient='vertical',command=tree.yview)
tree.configure(yscrollcommand=ys.set)
tree.pack(side='left',fill='both',expand=True)
ys.pack(side='right',fill='y')

cur=start
while cur.weekday()!=5: cur-=timedelta(days=1)
plan_total=worked_total=0
while cur<=end:
    vals=[cur.isoformat()]
    wp=ww=0
    for i in range(7):
      d=cur+timedelta(days=i)
      s,a,st=day_hours(d)
      vals.append(f'{s}/{a}')
      wp+=s; ww+=a
    vals += [round(wp,1),round(ww,1),round(ww-wp,1)]
    plan_total+=wp; worked_total+=ww
    tree.insert('', 'end', values=vals)
    cur += timedelta(days=7)

for t in [f'Start: {settings["start_date"]}',f'End: {settings["end_date"]}',f'Scheduled Hours: {plan_total:.1f}',f'Actual Hours: {worked_total:.1f}',f'Variance: {worked_total-plan_total:.1f}']:
    ttk.Label(right,text=t).pack(anchor='w',padx=10,pady=6)
root.mainloop()
