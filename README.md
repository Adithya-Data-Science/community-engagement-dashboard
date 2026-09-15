# Community Engagement & University Partnerships Dashboard

An Excel analytics project that converts a 100-record synthetic partnership dataset into an executive dashboard, supporting summary tables, charts, and a concise findings section.

![Dashboard preview](screenshots/dashboard.png)

## Business questions

The workbook answers five practical questions:

1. How many partnerships, participants, and student volunteers are represented?
2. How much funding is associated with the partnerships?
3. Which university departments and partnership types receive the most activity and funding?
4. Which impact areas reach the most participants?
5. What does the satisfaction data suggest about overall program performance?

## Headline results

| Metric | Result |
| --- | ---: |
| Partnerships | 100 |
| Participants engaged | 25,363 |
| Student volunteers | 3,736 |
| Total funding | $12,478,642 |
| Average satisfaction | 4.06 / 5.00 |

Additional findings:

- Technology Access reached 4,597 participants, the largest reach among the impact areas.
- Environmental Sciences received approximately $2.06 million, the highest departmental funding total.
- Health Initiative was the most common partnership type, with 14 partnerships.

## Workbook design

The project is organized into three layers:

- `Partnership_Data`: the 100-row source table.
- `Summary_Tables`: calculated KPI and category summaries used by the dashboard.
- `Dashboard`: five headline KPI cards, two charts, and a plain-language findings section.

The workbook uses Excel PivotTable-style summaries, lookup and validation logic, structured tables, number formatting, and charts to make the analysis reviewable by nontechnical stakeholders.

## Reproduce the analysis

1. Download [`workbook/Adithya_Community_Engagement_Dashboard.xlsx`](workbook/Adithya_Community_Engagement_Dashboard.xlsx).
2. Open it in Microsoft Excel.
3. Review the source records in `Partnership_Data`.
4. Trace the totals and category summaries in `Summary_Tables`.
5. Compare those results with the KPI cards, charts, and findings on `Dashboard`.
6. Change a source record and refresh the workbook summaries to validate how the dashboard responds.

## Validation performed

- Confirmed 100 unique partnership IDs.
- Reconciled participants, volunteers, funding, and average satisfaction to the source table.
- Checked department, partnership-type, and impact-area summaries against the source categories.
- Reviewed number formats, chart labels, and dashboard layout for readability.

See [VALIDATION.md](VALIDATION.md) and [DATA_DICTIONARY.md](DATA_DICTIONARY.md) for the control record and field definitions.

## Data note

The workbook uses a synthetic portfolio dataset created for a community-engagement analysis exercise. It contains no confidential university records or personal contact information.

## Tools

Microsoft Excel, PivotTables, XLOOKUP, data validation, charts, descriptive analysis, and executive reporting.
