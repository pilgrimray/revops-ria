# Revenue Operations Dashboard

Interactive Streamlit dashboard created as part of a Revenue Operations analytics case study.

Features:

- Executive KPI overview
- Conversion analytics
- Data quality controls
- Business investigation findings
- Revenue Operations insights
- AI usage documentation
- Interactive filtering by country, CRM, and date range

## Run locally

```powershell
pip install -r requirements.txt
streamlit run app.py
```

The app includes `data/revenue_operations.xlsx` by default. You can also upload another Excel file with the same column structure from the sidebar.

https://perebendya.streamlit.app/

## Win rate logic

Win rate is calculated as:

```text
Closed Won / (Closed Won + Closed Lost)
```

Open pipeline stages are excluded from win-rate denominators.
