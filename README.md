# DMart Sales Analysis – Excel Project

## Project Overview

This project analyzes retail sales data for DMart using Microsoft Excel. It was completed as a series of practical modules covering data inspection, cleaning and structuring, formula-based calculations, logical analysis, lookup functions, PivotTables, charts, and an interactive dashboard.

The objective is to turn raw retail transaction data into useful business information, such as sales and revenue performance across regions, product categories, customers, and time.

## Project Objectives

- Inspect and understand the original retail dataset.
- Calculate discount value and net revenue.
- Audit formulas and work with Excel named ranges.
- Clean and structure data using Excel Tables.
- Highlight important values using conditional formatting.
- Classify sales performance using logical functions.
- Create date-based summaries and charts.
- Retrieve data using lookup functions.
- Summarize data with PivotTables and PivotCharts.
- Build an interactive sales dashboard using slicers and a timeline.

## Tools and Features

- **Microsoft Excel**
- Excel formulas and functions: `SUM`, `IF`, nested `IF`, `AND`, `OR`, `NOT`, `IFERROR`, `TEXT`, `SUMIF`, `VLOOKUP`, `INDEX`, and `MATCH`
- Excel Tables and named ranges
- Conditional formatting
- PivotTables and PivotCharts
- Slicers and Timeline
- **Git and GitHub** for version control and project hosting

## Dataset

The original workbook is `DMart_Sales_Analysis.xlsx`.

The source dataset contains **15,345 transaction records** (excluding the header) and includes fields such as:

- Order ID
- Region
- Category
- Sales
- Discount
- Quantity
- Order Date
- Delivery Date
- Customer Type
- Payment Mode
- Profit
- City

The later modules build on this data and add calculated fields and analysis outputs. Some module workbooks contain additional helper columns, summary sheets, charts, or dashboard components.

## Module-wise Work

| Module | Workbook | Work covered |
|---|---|---|
| Original dataset | `DMart_Sales_Analysis.xlsx` | Initial inspection of the dataset, including column headings, records, and key categorical fields. |
| Module 2 – Revenue Calculations | `DMart_Sales_Analysis_Module2.xlsx` | Calculated discount value and net revenue using Excel formulas; worked with revenue-related business metrics. |
| Module 3 – Formula Auditing and Named Ranges | `DMart_Sales_Analysis_Module3.xlsx` | Reviewed formulas, traced calculation logic, and worked with named ranges for easier formula references. |
| Module 4 – Data Cleaning and Excel Tables | `DMart_Sales_Analysis_Module4.xlsx` | Worked on structuring retail data, standardizing fields, and organizing the data in an Excel Table named `DMart_Sales_Table`. |
| Module 5 – Conditional Formatting | `DMart_Sales_Analysis_Module5.xlsx` | Applied conditional formatting to make important sales and business values easier to identify. |
| Module 6 – Logical Functions and Classification | `DMart_Sales_Analysis_Module6.xlsx` | Used `IF`, nested `IF`, `AND`, `OR`, `NOT`, and `IFERROR` to classify sales performance and evaluate business conditions. |
| Module 7 – Date Analysis and Charts | `DMart_Sales_Analysis_Module7.xlsx` | Added month-based analysis and calculated summary values using functions such as `TEXT` and `SUMIF`; created a chart for regional sales analysis. |
| Module 8 – Lookup Functions and Dashboard | `DMart_Sales_Analysis_Module8.xlsx` | Practiced lookup techniques including `VLOOKUP`, `INDEX`, and `MATCH`; created supporting analysis sheets and a sales dashboard with charts. |
| Module 9 – PivotTables and PivotCharts | `DMart_Sales_Analysis_Module9.xlsx` | Added calculated fields for revenue analysis and created PivotTables, PivotCharts, and dashboard filtering elements. |
| Module 10 – Interactive Sales Dashboard | `DMart_Sales_Analysis_Module10.xlsx` | Consolidated the analysis into an interactive dashboard with KPI cards, PivotCharts, region and category slicers, and a timeline. |

## Final Dashboard

The Module 10 workbook contains a dashboard with four key performance indicators:

- **Total Sales:** ₹3,92,62,877.52
- **Net Revenue:** ₹3,53,36,589.77
- **Top Region:** North
- **Top Category:** Electronics

It also includes three visual analyses:

1. Sales by Region
2. Sales by Category
3. Monthly Sales Trend

The dashboard uses Region and Category slicers to filter the PivotTable-based analysis, along with a Timeline for date-based filtering.

## Key Business Insights

1. **Top Performing Region:** North recorded the highest sales among all regions.
2. **Top Performing Category:** Electronics was the leading product category.
3. **Total Sales:** The business recorded total sales of ₹3,92,62,877.52.
4. **Net Revenue:** Net revenue after accounting for discounts was ₹3,53,36,589.77.
5. **Interactive Analysis:** Region and Category slicers and a Timeline allow users to explore sales performance across different segments and periods.

## Repository Structure

The repository contains the original dataset and the separate Excel workbooks for the module-wise work:

```text
dmart-sales-analysis/
│
├── README.md
├── DMart_Sales_Analysis.xlsx
├── DMart_Sales_Analysis_Module2.xlsx
├── DMart_Sales_Analysis_Module3.xlsx
├── DMart_Sales_Analysis_Module4.xlsx
├── DMart_Sales_Analysis_Module5.xlsx
├── DMart_Sales_Analysis_Module6.xlsx
├── DMart_Sales_Analysis_Module7.xlsx
├── DMart_Sales_Analysis_Module8.xlsx
├── DMart_Sales_Analysis_Module9.xlsx
└── DMart_Sales_Analysis_Module10.xlsx
```

*The names shown above are the clean repository filenames. If a workbook has a different exact filename in the repository, use that exact spelling in the file list.*

## Key Learning Outcomes

Through this project, I gained practical experience in:

- Preparing and structuring retail data in Excel.
- Applying formulas to calculate business metrics.
- Auditing formulas and using named ranges.
- Using logical and lookup functions for analysis.
- Summarizing large datasets with PivotTables.
- Presenting results through charts and an interactive dashboard.
- Maintaining project versions and sharing work through GitHub.

## How to Use

1. Download or clone this repository.
2. Open the required `.xlsx` workbook in Microsoft Excel.
3. Start with the original dataset and review the module workbooks in sequence.
4. Open `DMart_Sales_Analysis_Module10.xlsx` to explore the final dashboard.
5. Use the dashboard slicers and timeline in Excel to interact with the analysis.

## Author

**Manvi Gupta**

---

This repository is an Excel-based learning and portfolio project focused on retail sales analysis and business reporting.
