## Brief Overview
This project walks you through building a complete Excel dashboard—from raw data to interactive visualizations—using XLOOKUP, INDEX‑MATCH, PivotTables, slicers, and custom styling.

## Project Overview
- **Workflow:** Data gathering → transformation → analysis → interactive dashboard  
- **Visuals:**  
  - Line chart: total sales over time by coffee type  
  - Bar chart: sales by country (U.S., Ireland, UK)  
  - Bar chart: top‑5 customers  
  - Timeline slicer + three additional slicers (roast type, size, loyalty card)

# Coffee Sales Dashboard — Excel Analytics Project ☕📊

This project documents how I built an Excel dashboard from raw order data through to a stakeholder-ready reporting view. The focus is on a typical analytics workflow: **data understanding → cleaning → modeling → KPI definition → analysis → dashboard delivery**.

---

## Background & Objective

The business needed a simple way to monitor coffee sales performance across:
- **Time** (monthly trend + period comparisons)
- **Product** (coffee type, roast type, pack size)
- **Customer** (top customers, loyalty segmentation)
- **Geography** (sales by country)

The output is a single Excel dashboard that lets stakeholders filter results using a **timeline** and **slicers** and instantly see the impact on key metrics and charts.

---

## Dataset & Structure

The workbook is organized as a small “model” with one fact table and two dimension tables:

- **Orders** (fact table): order-level transactions  
  Keys: `Order ID`, `Customer ID`, `Product ID`, `Order Date`, `Quantity`
- **Customers** (dimension): customer attributes  
  Fields: name, contact, country, loyalty card flag
- **Products** (dimension): product attributes  
  Fields: coffee type, roast type, size, unit price, profit fields

I used the Orders sheet as the reporting backbone and enriched it with customer/product fields to make it analysis-ready.

---

## Data Understanding (What I validated first)

Before building anything, I checked:
- Whether **Order ID** behaved like a true unique key
- Whether **Customer ID** and **Product ID** were consistently populated
- Whether dates were valid and within expected ranges
- Whether quantity and pricing fields had sensible values (no negatives, no blanks where required)

This helped reduce issues later when pivots and filters depended on consistent categories.

---

## Data Cleaning & Preparation

### 1) De-duplication
I removed duplicates from the Orders dataset using Excel’s built-in **Remove Duplicates** to ensure metrics like total sales and customer ranking weren’t inflated.

### 2) Standardized formatting
To support consistent reporting:
- `Order Date` formatted as `dd mmm yyyy` (stable month labels)
- Prices and sales formatted as **US $**
- Size formatted as `0.0 kg` for readability

### 3) Convert to an Excel Table
I converted Orders into a structured table (`OrdersTable`).  
This ensures:
- Pivot sources expand automatically as new rows are added
- Formulas fill down consistently
- Structured references make maintenance easier

---

## Data Enrichment (Building an analysis-ready table)

The Orders table initially had missing descriptive fields, so I populated them using lookups:

### Customer attributes (XLOOKUP)
Pulled fields like:
- Customer Name
- Email
- Country
- Loyalty Card (added later as a refreshable field)

### Product attributes (INDEX + MATCH)
Pulled fields like:
- Coffee Type
- Roast Type
- Size
- Unit Price

I used a header-driven INDEX/MATCH approach so the same formula pattern could populate multiple columns, reducing manual work and improving consistency.

---

## KPI Definitions

To align reporting with stakeholder needs, I defined KPIs that answer common business questions:

### Primary KPIs
- **Total Sales**  
  `Sales = Quantity × Unit Price`
- **Sales Trend (MoM)**  
  Monthly aggregation of Total Sales for trend monitoring

### Supporting KPIs / Breakdowns
- **Sales by Coffee Type** (Arabica / Excelsa / Liberica / Robusta)
- **Sales by Country** (U.S. / Ireland / UK)
- **Top Customers by Sales** (Top 5)
- **Sales by Roast Type** (Dark / Light / Medium)
- **Sales by Size** (0.2kg / 0.5kg / 1kg / 2.5kg)
- **Loyalty vs Non-Loyalty** segmentation

These were chosen because they map directly to commercial levers: product mix, geographic performance, and customer value.

---

## Analysis Approach

I used PivotTables as the core analysis layer:

### 1) Time series performance
Grouped Order Date into **Years + Months** to monitor monthly sales behavior and seasonality.

### 2) Geographic performance
Aggregated sales by Country and sorted to highlight the strongest market quickly.

### 3) Customer concentration
Ranked customers by sales and filtered to Top 5 to show revenue concentration and key accounts.

---

## Dashboard Design (How it’s meant to be used)

The dashboard is built for quick decision support:
- Start with the **timeline** to select the period (e.g., last quarter, YTD, specific months)
- Use slicers to segment performance:
  - Roast Type
  - Size
  - Loyalty Card
- Charts update together so stakeholders can explore product/customer/geography in one view

### Visuals included
- Total Sales Over Time (line chart with coffee type series)
- Sales by Country (bar chart)
- Top 5 Customers (bar chart)

---

## Interactivity & Governance

### Slicer/Timestamp connections
All slicers and the timeline are connected across all pivot charts using **Report Connections**, ensuring every view is consistent and filters behave predictably.

### Refresh behavior
Because pivots are built from `OrdersTable`, updates are controlled and repeatable:
- Add rows to OrdersTable → Refresh pivots → Dashboard updates

---

## Output & Value

The final deliverable is a dashboard that:
- Reduces manual reporting effort (no manual filtering or rebuilding charts)
- Makes performance drivers visible (mix, geography, customer concentration)
- Supports faster stakeholder conversations (interactive exploration during reviews)

---

## Limitations & Next Enhancements

If this were extended for production reporting, I would add:
- KPI tiles (Total Sales, Orders, Avg Order Value, YTD vs prior period)
- Profit reporting (using product profit fields for margin analysis)
- Data quality checks (missing IDs, invalid dates, outlier detection)
- A “Reset Filters” control for easier executive use

---
