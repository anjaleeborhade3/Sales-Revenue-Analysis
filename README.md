# Sales Revenue Analysis Dashboard

An Excel dashboard and Python workflow for exploring 250 simulated sales transactions across products, categories, regions, and 2026 monthly periods.

## What the project shows

- Revenue and profit totals, order count, units sold, and average order revenue
- Revenue and profit comparisons by region and product category
- Monthly revenue trend
- A filterable transaction-level dataset with consistent currency, date, and discount formats

## Project files

| File | Description |
|---|---|
| `Sales_Revenue_Analysis.xlsx` | Formatted Excel workbook with transaction data, formula-driven summaries, and charts |
| `sales_data.csv` | Source transaction data used by the project |
| `create_project.py` | Python script used to generate the dataset and workbook |

## Workbook layout

- **Sales Data** — 250 simulated orders, with filters and a frozen header row
- **Analysis** — KPI summary, regional and category summaries, monthly revenue table, and three charts

## Tools and techniques

Python, pandas, Microsoft Excel, Excel formulas (`SUM`, `AVERAGE`, `COUNTA`, `SUMIF`, `SUMIFS`), and charts.

## Notes

The dataset is simulated for portfolio and practice use. The workbook summaries reference the transaction rows directly so they recalculate when opened in Excel.
