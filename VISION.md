# Vision
2
 
3
## Canadian Work Hours Calendar
4
 
5
The Canadian Work Hours Calendar is a desktop application designed to help knowledge workers, consultants, contractors, and managers plan, forecast, and track work hours across extended periods of time.
6
 
7
The primary goal is not timesheet entry.
8
 
9
The primary goal is to answer planning questions such as:
10
 
11
- How many hours am I expected to work for the remainder of the year?
12
- How do vacation periods affect my planned availability?
13
- What is my expected year-end total?
14
- How do my actual worked hours compare to my plan?
15
- Am I ahead of or behind my expected schedule?
16
 
17
The application treats forecasting as a first-class capability and time tracking as a supporting capability.
18
 
19
---
20
 
21
# Why This Application Exists
22
 
23
Many calendar applications focus on appointments.
24
 
25
Many timecard systems focus on historical reporting.
26
 
27
Neither provides a simple way to answer:
28
 
29
> "Given my normal work pattern, holidays, vacations, and actual hours worked, what does the rest of the year look like?"
30
 
31
The purpose of this application is to bridge that gap.
32
 
33
The system automatically generates a schedule using predefined work patterns and allows the user to modify only the exceptions.
34
 
35
Instead of entering hundreds of future dates manually, users should:
36
 
37
1. Define a normal work pattern.
38
2. Generate a forecast.
39
3. Mark exceptions.
40
4. Review forecasts and totals.
41
 
42
---
43
 
44
# Design Principles
45
 
46
## Forecast First
47
 
48
The application should always answer:
49
 
50
- Remaining planned hours
51
- Remaining work days
52
- Vacation days
53
- Holiday days
54
- Projected bank balance
55
 
56
with minimal effort.
57
 
58
## Exceptions Over Manual Entry
59
 
60
The system should generate sensible defaults.
61
 
62
A user should rarely need to enter individual future dates.
63
 
64
Instead, they should modify:
65
 
66
- Vacations
67
- Holidays
68
- Half-days
69
- Special working days
70
 
71
as exceptions to the default schedule.
72
 
73
## Canadian Focus
74
 
75
The application is designed with Canadian workplaces in mind.
76
 
77
Defaults include:
78
 
79
- Saturday-start work weeks
80
- Canadian spelling
81
- Ontario holiday support
82
- 7.5-hour standard work day
83
 
84
Other provinces and territories may be added later.
85
 
86
## Daily Accuracy
87
 
88
Each day can contain:
89
 
90
- Planned hours
91
- Worked hours
92
- Multiple work entries
93
- Notes
94
- Leave status
95
 
96
Actual worked time may exceed planned hours.
97
 
98
Leave and holiday days are credited according to policy.
99
 
100
---
101
 
102
# User Interface Vision
103
 
104
The main screen is a scrollable weekly ledger.
105
 
106
Each week is displayed as a card.
107
 
108
Example:
109
 
110
Sat Sun Mon Tue Wed Thu Fri
111
----------------------------------------------------
112
Planned 0 0 7.5 7.5 7.5 7.5 7.5
113
 
114
Worked 0 0 8.0 9.0 8.0 5.5 7.0
115
 
116
Status O O W W W W W
117
 
118
Totals:
119
Plan 37.5
120
Worked 37.5
121
Variance 0.0
122
Bank +1.5
123
 
124
Users can immediately see:
125
 
126
- Planned effort
127
- Actual effort
128
- Week totals
129
- Variance
130
- Bank balance
131
 
132
without opening additional screens.
133
 
134
---
135
 
136
# Day Interaction Model
137
 
138
Single-click:
139
 
140
- Change status quickly
141
- Working
142
- Vacation
143
- Holiday
144
- Half-Day
145
- Off
146
 
147
Double-click:
148
 
149
- Open detailed editor
150
- View notes
151
- View work entries
152
- Add work sessions
153
- Edit credited hours
154
 
155
This allows both rapid planning and detailed editing.
156
 
157
---
158
 
159
# Data Ownership
160
 
161
All user data is stored locally in JSON files.
162
 
163
Core files:
164
 
165
- settings.json
166
- holidays_on.json
167
- work_log.json
168
 
169
The work log acts as the system of record.
170
 
171
It stores:
172
 
173
- Actual work entries
174
- Vacation days
175
- Holidays
176
- Notes
177
- Manual adjustments
178
 
179
No external service is required.
180
 
181
---
182
 
183
# Long-Term Vision
184
 
185
The application should evolve into a complete personal work-planning system that can:
186
 
187
- Forecast future work hours
188
- Track actual work performed
189
- Calculate variances automatically
190
- Maintain running bank balances
191
- Generate management reports
192
- Import existing timecard data
193
- Export reports to Excel and PDF
194
 
195
A manager should be able to ask:
196
 
197
"How many hours are you expected to work between now and year-end?"
198
 
199
and the application should provide the answer immediately.
