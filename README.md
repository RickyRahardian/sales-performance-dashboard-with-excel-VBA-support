# Sales Performance Dashboard

An interactive, macro-enabled Excel dashboard for tracking sales performance across regions, sales channels, and product categories. Built with Excel and VBA to support dynamic filtering, automated data entry, and real-time reporting.

## Overview

This dashboard provides a consolidated view of sales KPIs and trends, allowing quick analysis of revenue, orders, and product performance without manual pivot table updates.

## Key Metrics Tracked

- **Total Revenue**
- **Total Orders**
- **Product Sold (Pcs)**
- **Average Order Value**

## Features

- **Multi-level filtering** by Region, Sales Channel, Product Category, and Product
- **Time-based slicers** for Quarter, Month, and Day
- **Revenue by Month** trend chart
- **Revenue by Region** comparison
- **Top 5 Products by Revenue** ranking
- **Revenue by Sales Channel** breakdown (donut chart)
- **Revenue by Category** comparison
- **Three dedicated views**: Dashboard, Summary, and raw Data sheet

## VBA Automation

The workbook uses VBA macros to:
- Auto-generate unique IDs on new data entry
- Enable full-screen dashboard mode for presentations
- Auto-refresh pivot tables and charts when data is updated

## Files in This Repository

| File | Description |
|------|--------------|
| `Interactive_Sales_Data_Dashboard-Ricky.xlsm` | Main dashboard file (macro-enabled) |
| `Sales_Raw_Data.csv` | Source dataset used by the dashboard |
| `Dashboard.jpg` | Preview of the dashboard interface |

## How to Use

1. Download `Interactive_Sales_Data_Dashboard-Ricky.xlsm`
2. Open the file in Microsoft Excel
3. When prompted, click **Enable Content** / **Enable Macros** to activate the VBA features
4. Use the filter panel on the left (Region, Sales Channel, Product Category, Product) to explore the data
5. Use the date slicers at the top to adjust the reporting period

> **Note:** Macros must be enabled for auto-refresh, full-screen mode, and ID generation features to work.

## Tools Used

- Microsoft Excel
- VBA (Visual Basic for Applications)

## Author

**Ricky Rahardian Bimantara**
[LinkedIn](https://linkedin.com/in/rickyrahardianbimantara) · [GitHub](https://github.com/RickyRahardian)
