# ExcelProject-End-to-End-Dashboard
I developed a comprehensive project in excel, creating dashboards and tables to analyze data. This process involved several stages, including gata gathering, transformation, analysis and in the end an interactive dashboard.

Below is an updated README.md that now includes a clickable Table of Contents.
Just replace the existing file with this content (or insert the TOC at the top of your current file).

# 📊 Excel Portfolio Project: End‑to‑End Dashboard

## Table of Contents
- [Brief Overview](#brief-overview)
- [Project Overview](#project-overview)
- [Data Structure](#data-structure)
- [Data Gathering – Lookup Formulas](#data-gathering--lookup-formulas)
  - [XLOOKUP for Customer Info](#xlookup-for-customer-info)
  - [INDEX‑MATCH for Product Info](#index‑match-for-product-info)
- [Calculated Columns](#calculated-columns)
- [Formatting](#formatting)
- [Data Validation](#data-validation)
- [Pivot Tables & Charts](#pivot-tables--charts)
- [Interactive Dashboard Elements](#interactive-dashboard-elements)
- [Pivot Table Refinements](#pivot-table-refinements)
- [Pivot Chart Customization](#pivot-chart-customization)
- [Timeline Insertion & Styling](#timeline-insertion--styling)
- [Slicer Creation & Styling](#slicer-creation--styling)
- [Updating Data Source with Loyalty Card](#updating-data-source-with-loyalty-card)
- [Additional Pivot Tables & Charts](#additional-pivot-tables--charts)
- [Dashboard Assembly](#dashboard-assembly)
- [Key Takeaway](#key-takeaway)

---

## Brief Overview
This project walks you through building a complete Excel dashboard—from raw data to interactive visualizations—using XLOOKUP, INDEX‑MATCH, PivotTables, slicers, and custom styling.

## Project Overview
- **Workflow:** Data gathering → transformation → analysis → interactive dashboard  
- **Visuals:**  
  - Line chart: total sales over time by coffee type  
  - Bar chart: sales by country (U.S., Ireland, UK)  
  - Bar chart: top‑5 customers  
  - Timeline slicer + three additional slicers (roast type, size, loyalty card)

## Data Structure
|
 Worksheet 
|
 Primary Key 
|
 Key Columns (example) 
|
|
---------
|
-----------
|
----------------------
|
|
**
Orders
**
|
 Order ID 
|
 Order Date, Customer ID, Product ID, Quantity 
|
|
**
Customers
**
|
 Customer ID 
|
 Customer Name, Email, Phone, Address, City, Country, Postcode, Loyalty Card 
|
|
**
Products
**
|
 Product ID 
|
 Coffee Type, Roast Type, Unit Price, Price / 100 g, Profit 
|

*Columns F–M in **Orders** are initially empty and will be filled via lookups.*

## Data Gathering – Lookup Formulas
### XLOOKUP for Customer Info
```excel
=XLOOKUP(C2, Customers!A:A, Customers!B:B, "", 0)   'Customer Name in F2
C2 = Customer ID
Wrap with IF(...=0,"",…) to suppress zeros.
INDEX‑MATCH for Product Info (dynamic single formula)
=INDEX(Products!E:E, MATCH(D2, Products!A:A, 0), MATCH(I$1, Products!$1:$1, 0))
Drag right/down to populate Roast Type, Size, Unit Price, etc.
Calculated Columns
Sales (L2): =K2*L2
Full Coffee‑Type Name (N):
=IF(I2="ROB","Robusta",
   IF(I2="EXE","Excelsa",
      IF(I2="ARA","Arabica",
         IF(I2="LIB","Liberica",""))))
Roast‑Type Name (O):
=IF(J2="M","Medium",
   IF(J2="L","Light",
      IF(J2="D","Dark","")))
Formatting
Order Date: dd‑mmm‑yyyy (e.g., 05‑Sep‑2023)
Size: 0.0" kg"
Unit Price & Sales: US $ currency
Data Validation
Remove duplicates (Data → Remove Duplicates).
Convert Orders to a table (Ctrl T) → name it OrdersTable.
Pivot Tables & Charts
Insert Pivot Table (Alt → N → V → T).
Total Sales (line chart):
Rows → Order Date → Group → Years & Months.
Values → Sales (Sum).
Sales by Country (bar chart).
Top‑5 Customers (bar chart, filter to top 5).
Interactive Dashboard Elements
Slicer / Control	Field	Effect
Timeline	Order Date	Filters all visuals by period
Roast‑Type Slicer	Roast Type	Shows selected roast categories
Size Slicer	Package Size	Filters by bean package size
Loyalty‑Card Slicer	Loyalty Card	Isolates customers with/without cards
All slicers are linked to the pivot tables for a dynamic dashboard.

Pivot Table Refinements
Group dates by Years & Months (Ctrl‑click both).
Layout → Show in Tabular Form, disable Grand Totals/Subtotals.
Add Coffee Type to Columns, Sales to Values.
Number format → Thousands separator, 0 decimals.
Pivot Chart Customization
Insert Line Chart, hide field buttons.
Apply purple theme (RGB 60, 20, 100) to chart area and text.
Axis styling: white line, thicker weight.
Titles: USD (vertical), Total Sales Over Time (chart).
Series colors: Liberica → Yellow, Excelsa → Brown, Arabica → Bright Blue, Robusta → Red.
Timeline Insertion & Styling
PivotChart Analyze → Insert Timeline → Order Date.
Create a custom purple style (header white, selected block white, unselected light purple).
Slicer Creation & Styling
Insert slicers for Size, Roast‑Type, Loyalty Card.
Build a custom purple slicer style (headers white on dark purple, selected items white border, etc.).
Set column layout (e.g., Roast‑Type → 3 columns).
Updating Data Source with Loyalty Card
=XLOOKUP([@Customer_ID], Customers!A:A, Customers!I:I, "", 0)
Add Loyalty Card column (P1) in Orders.
Refresh pivot tables to include the new field.
Additional Pivot Tables & Charts
Country Sales Bar Chart: duplicate Total Sales sheet, swap Coffee Type for Country, green bars with country‑specific shades, data labels, US $ formatting.
Top‑5 Customers Bar Chart: duplicate Country chart, set row field to Customer Name, apply Top 5 filter, sort ascending, rename title.
Dashboard Assembly
New worksheet Dashboard.
Adjust column A width, row 1 height (~5).
Insert a large purple rectangle (RGB 60‑20‑100) spanning A:Z, white font, title Coffee Sales Dashboard.
Cut & paste timeline, line chart, country bar chart, top‑5 chart, and slicers onto the Dashboard.
Align with Alt‑drag for snapping.
Report Connections: link timeline and slicers to all three visuals.
UI cleanup: View → uncheck Gridlines, hide formula bar, scrollbars, sheet tabs via File → Options → Advanced if desired.
Key Takeaway
By converting raw data into a structured table, leveraging PivotTables/Charts, and applying consistent purple styling, you can build an interactive, professional‑looking dashboard that updates automatically as new orders are added.

Feel free to copy this markdown into your repository’s README.md. Would you like help adding any badges or a screenshot to the top of the file?
