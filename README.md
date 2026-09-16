# Community Engagement and University Partnerships Dashboard

An end-to-end Excel analytics project that converts 100 synthetic university-community partnership records into an executive dashboard with formula-linked KPIs, summary tables, filters, number formats, and charts.

## Business questions

- How many partnerships, participants, and student volunteers are represented?
- How much funding is associated with the portfolio?
- Which departments and impact areas account for the most activity?
- What does satisfaction data indicate about overall program performance?

## Headline results

| Metric | Result |
|---|---:|
| Partnerships | 100 |
| Participants | 25,363 |
| Student volunteers | 3,736 |
| Total funding | $12,478,642 |
| Average satisfaction | 4.06 / 5.00 |

## Start-to-finish workflow

1. Generate a deterministic 100-record synthetic partnership dataset.
2. Preserve one row per partnership with department, partnership type, impact area, participants, volunteers, funding, and satisfaction.
3. Build formula-linked KPI totals and category summaries.
4. Apply filters, validation-friendly formatting, and appropriate number formats.
5. Add charts for funding by department and participants by impact area.
6. Reconcile all dashboard KPIs to the source table.

## Rebuild the workbook

```bash
python -m venv .venv
# Windows PowerShell: .venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python build_dashboard.py
```

The script writes `Community_Engagement_Dashboard.xlsx`. Detailed instructions are in `BUILD.md`; field definitions are in `DATA_DICTIONARY.md`; quality controls are in `VALIDATION.md`.

## Data and limitations

All records are synthetic and contain no real student, employee, university, or community-partner information. The project demonstrates Excel analytics and reporting practices and should not be interpreted as an official university report.

