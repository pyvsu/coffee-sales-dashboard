# Coffee Sales Dashboard (Excel)

An interactive Excel dashboard that shows coffee bean sales by coffee type, country, and customer. One click on a timeline or slicer filters every chart at once.

<img width="1794" height="898" alt="overview" src="https://github.com/user-attachments/assets/3b34f074-eff2-4bb5-a3cc-47bc5c300d44" />


## Business Question

Which coffee types, countries, and customers drive sales, and how do sales change over time, by roast type, package size, and loyalty card status?

## What the Dashboard Shows

**Charts**
- **Total Sales Over Time:** line chart split by coffee type (Arabica, Excelsa, Liberica, Robusta)
- **Sales by Country:** bar chart for the US, Ireland, and the UK
- **Top 5 Customers:** bar chart of the five biggest customers by sales

**Filters**
- **Timeline:** filter by order date
- **Slicers:** roast type (Dark, Light, Medium), size (0.2, 0.5, 1, 2.5 kg), loyalty card (Yes, No)

## Dataset

The Excel file has three tables:

| Table | Key | Main columns |
| --- | --- | --- |
| Orders | Order ID | Order Date, Customer ID, Product ID, Quantity |
| Customers | Customer ID | Name, Email, Country, Loyalty Card |
| Products | Product ID | Coffee Type, Roast Type, Size, Unit Price |

## How It Was Built

1. **Gathered the data.** Used `XLOOKUP` to add customer name, email, country, and loyalty card to the orders. Used `INDEX` + `MATCH` to add coffee type, roast type, size, and unit price with one dynamic formula.
2. **Transformed the data.**
    - Calculated `Sales` as unit price × quantity.
    - Used nested `IF` formulas to turn codes into full names (for example Rob → Robusta, M → Medium).
    - Formatted dates, sizes (kg), and currency (USD).
    - Checked for duplicates.
3. **Converted the range to an Excel Table** so the pivot tables pick up new rows and columns on refresh.
4. **Built pivot tables and pivot charts** for sales over time, sales by country, and top 5 customers.
5. **Added a timeline and three slicers**, then styled them with a custom theme.
6. **Assembled the dashboard sheet** and connected the timeline and slicers to all charts using Report Connections.
7. **Cleaned up the view:** turned off gridlines and hid the ribbon.

## Excel Skills Used

- Lookup formulas: `XLOOKUP`, `INDEX`, `MATCH`
- Logic formulas: nested `IF`
- Absolute and mixed cell references (`$D2`, `I$1`)
- Custom number formats (dates, kg)
- Excel Tables
- Pivot tables and pivot charts (grouping, value filters, sorting)
- Timeline and slicers, with custom styles
- Dashboard layout and report connections

## How to Use

1. Download `coffee-sales-dashboard.xlsx` from this repo.
2. Open the **Dashboard** sheet.
3. Click the timeline or slicers to filter the charts. Clear a filter to see all data again.
4. If you change the data, go to **PivotTable Analyze → Refresh All**.

> Requires Excel 365 or Excel 2021 (for `XLOOKUP`).

## Repository Structure

```
├── coffee-sales-dashboard.xlsx   # Final workbook
├── images/
│   └── dashboard.png             # Dashboard screenshot
└── README.md
```

## Credits

Based on a tutorial by Mo Chen: [The ONLY EXCEL PORTFOLIO PROJECT YOU NEED](https://youtu.be/m13o5aqeCbM?si=T1NLkLE_tGnPVy46). I rebuilt the dashboard step by step and changed the color theme to make it my own.
