# Revenue Operations Streamlit Dashboard

This dashboard reads the included Excel workbook and visualizes:

- Win rate by country
- Win rate by CRM
- Deal distribution by stage
- Deal count by source
- Win rate by PPC budget
- Data quality checks for closing dates, loss reasons, AQL dates, and missing CRM values

## Run locally

```powershell
pip install -r requirements.txt
streamlit run app.py
```

The app includes `data/revenue_operations.xlsx` by default. You can also upload another Excel file with the same column structure from the sidebar.

## Win rate logic

Win rate is calculated as:

```text
Closed Won / (Closed Won + Closed Lost)
```

Open pipeline stages are excluded from win-rate denominators.
