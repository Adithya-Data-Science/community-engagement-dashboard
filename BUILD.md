# Rebuild the workbook

The repository includes a deterministic generator for the complete Excel workbook.

```bash
python -m venv .venv
# Windows PowerShell: .venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python build_dashboard.py
```

The script writes `Community_Engagement_Dashboard.xlsx` with a dashboard sheet, a 100-row source table, formula-linked KPIs, summary tables, filters, number formats, and two charts. The generated totals reconcile to 25,363 participants, 3,736 student volunteers, $12,478,642 in funding, and 4.06 average satisfaction.

