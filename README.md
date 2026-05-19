# 📦 Inventory & Supply Chain Management System with Reorder Alerts

An automated Excel workbook (.xlsm) designed to track real-time warehouse inventory, monitor historical supply chain movements, and automatically flag products that have fallen below minimum stock thresholds.

---

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Workbook Structure](#workbook-structure)
- [Data Dictionary](#data-dictionary)
- [Functions & Automation](#functions--automation)
- [Data Cleaning Process](#data-cleaning-process)
- [Summary Statistics](#summary-statistics)
- [Getting Started](#getting-started)
- [Requirements](#requirements)

---

## Overview

This workbook provides an end-to-end inventory management solution for warehouse operations. It ingests raw supply chain data, cleans it, tracks live stock levels, triggers reorder alerts when stock falls below defined thresholds, and presents insights through a pivot table and dashboard.

The dataset covers **124 unique products** across **7 categories**, with a combined total stock of **55,039 units** and **453 active reorder alerts**.

---

## Features

- ✅ **Automated Reorder Alerts** — Products are automatically flagged for reorder when live stock falls below the defined reorder level
- ✅ **Live Stock Calculation** — Stock quantities are computed in real time using SUMIF across Stock In and Stock Out sheets
- ✅ **Macro-Powered Auto-Fill** — Entering a Product ID automatically populates Product Name, Supplier, and Warehouse Location
- ✅ **Data Cleaning Pipeline** — Raw data is processed and validated before analysis
- ✅ **Category-Level Summary** — Aggregated totals by product category
- ✅ **Interactive Dashboard** — Visual overview of inventory health and supply chain status
- ✅ **Pivot Table Analysis** — Flexible cross-tabulation of inventory data

---

## Workbook Structure

The workbook contains **13 sheets**, each serving a distinct purpose:

| Sheet | Description |
|---|---|
| **Welcome Page** | Title page and introduction to the workbook |
| **Content** | Table of contents listing all sheets |
| **Instructions** | Detailed guide on how to use the workbook and the functions used |
| **Cleaning Tasks** | Documentation of all data cleaning steps applied to the raw data |
| **Raw Data** | Original, unprocessed inventory and supply chain data |
| **Clean_Data** | Cleaned and validated dataset ready for analysis |
| **Products** | Master product reference table |
| **Stock In** | Records of all incoming stock/deliveries |
| **Stock Out** | Records of all outgoing stock/sales |
| **Reorder Alerts** | Live list of products that have breached their reorder threshold |
| **Summary Analysis** | Category-level aggregation of total products, stock, and reorder counts |
| **Pivot Table** | Dynamic pivot analysis of the inventory data |
| **Dashboard** | Visual summary of key inventory and supply chain metrics |

---

## Data Dictionary

### Clean Data Sheet (Primary Working Dataset)

| Column | Description |
|---|---|
| `Product_ID` | Unique identifier for each product (e.g., P001) |
| `Product_Name` | Name of the product |
| `Category` | Product category (e.g., Dairy, Seafood, Bakery) |
| `Supplier_Name` | Name of the supplying company |
| `Warehouse_Location` | Physical address of the warehouse storing the product |
| `Status` | Current product status: `Active`, `Backordered`, or `Discontinued` |
| `Order_ID` | Unique order reference number |
| `Supplier_ID` | Unique identifier for the supplier |
| `Date_Received` | Date the stock was received at the warehouse |
| `Last_Order_Date` | Date of the most recent order placed |
| `Stock_Quantity` | Raw stock count from the original dataset |
| `Live_Stock_Quantity` | Calculated live stock (Stock In minus Stock Out) |
| `Reorder_Level` | Minimum stock threshold before a reorder is triggered |
| `Reorder_Quantity` | Quantity to order when restocking |
| `Reorder_Status` | `Yes` if live stock is below reorder level, `No` otherwise |
| `Unit_Price` | Price per unit of the product |
| `Sales_Volume` | Total units sold |
| `Inventory_Turnover_Rate` | Rate at which inventory is sold and replenished |

### Product Categories

| Category | Total Products |
|---|---|
| Fruits & Vegetables | 42 |
| Dairy | 24 |
| Grains & Pulses | 20 |
| Bakery | 10 |
| Oils & Fats | 10 |
| Seafood | 10 |
| Beverages | 8 |

---

## Functions & Automation

### Excel Functions Used

| Function | Description | How It Is Used in This Workbook |
|---|---|---|
| `IF` | Tests a logical condition and returns one value if the condition is true and another if it is false | Checks whether the Live Stock Quantity is below the Reorder Level and returns `Yes` or `No` in the Reorder Status column; also used to suppress blank rows in the Stock Out and Reorder Alerts sheets |
| `SUMIF` | Adds all values in a range that match a single criterion | Calculates each product's Live Stock Quantity by summing all incoming deliveries from the Stock In sheet and subtracting all outgoing sales from the Stock Out sheet, matched by Product ID |
| `COUNTIF` | Counts the number of cells in a range that meet a single condition | Counts how many products belong to each category in the Summary Analysis sheet |
| `COUNTIFS` | Counts cells that meet multiple conditions simultaneously | Counts how many products within each category have a Reorder Status of `Yes`, giving the "Number of Reorder Products" per category in the Summary Analysis sheet |
| `SUM` | Adds all values in a specified range | Totals the category-level figures (Total Products, Total Stock, Reorder Products) into overall grand totals in the Summary Analysis sheet |
| `VLOOKUP` | Searches for a value in the first column of a range and returns a value from a specified column in the same row | Retrieves the Live Stock Quantity for each product flagged in the Reorder Alerts sheet by looking up the Product ID against the Clean_Data sheet |
| `IFERROR` | Returns a custom value if a formula produces an error, otherwise returns the formula result | Wraps VLOOKUP calls so that if a Product ID is not found, the cell displays a blank or "Not Found" instead of an error code such as `#N/A` |
| `ABS` | Returns the absolute (non-negative) value of a number | Converts any negative stock quantities in the Raw Data into positive values during the data cleaning step |
| `PROPER` | Converts a text string so that the first letter of each word is capitalised and the rest are lowercase | Standardises inconsistently cased text across product name and category columns when pulling data from Raw Data into Clean_Data |
| **XLOOKUP** | Searches a range for a match and returns the corresponding value from another range; a more flexible successor to VLOOKUP | Used in the Content, Stock Out, and Reorder Alerts sheets to auto-populate supplier and product details from Clean_Data based on a Product ID lookup |
| **Macro (VBA)** | A set of programmed instructions that run automatically in the background to perform repetitive tasks | Instantly auto-fills the Product Name, Supplier, and Warehouse Location columns the moment a Product ID is typed into the Stock In or Stock Out sheets |

---

## Data Cleaning Process

The following cleaning techniques were applied to the raw data before analysis:

| Technique | Problem Addressed | Method |
|---|---|---|
| Remove Duplicates | Repeated records in the dataset | Excel's Remove Duplicates tool (Data Ribbon) |
| Negative Values | Negative quantities in the `Stock_Quantity` column | `ABS()` function |
| Missing Values | Blank cells causing formula errors | Go To Special feature |
| Trim Extra Spaces | Leading/trailing spaces breaking VLOOKUP | `TRIM()` function |
| Standardize Text Case | Inconsistent capitalisation across columns | `PROPER()` function |
| Date Formatting | Inconsistent or unrecognised date formats | Number Format → Date |
| Remove Irrelevant Data | Columns/rows not needed for analysis | Manual deletion |
<img width="1599" height="730" alt="WhatsApp Image 2026-05-19 at 17 50 59" src="https://github.com/user-attachments/assets/e8b37821-9baa-4011-9045-4b71ed810e22" />


---

## Summary Statistics

| Category | Total Products | Total Stock | Products Needing Reorder |
|---|---|---|---|
| Fruits & Vegetables | 42 | 18,806 | 143 |
| Oils & Fats | 10 | 4,400 | 36 |
| Dairy | 24 | 10,174 | 84 |
| Grains & Pulses | 20 | 8,347 | 75 |
| Seafood | 10 | 5,256 | 40 |
| Bakery | 10 | 3,916 | 37 |
| Beverages | 8 | 4,140 | 38 |
| **Total** | **124** | **55,039** | **453** |

---

## Getting Started

1. **Enable Macros** — When opening the file, click *Enable Content* in the Excel security bar. The macro auto-fill feature will not work without this step.
2. **Review the Instructions sheet** — Start here for a walkthrough of how each sheet connects.
3. **Do not edit the Raw Data sheet** — Use the Clean_Data sheet for all analysis.
4. **Log new stock movements** — Enter incoming deliveries in the **Stock In** sheet and outgoing sales in the **Stock Out** sheet. Live stock in the Clean_Data sheet will update automatically.
5. **Monitor Reorder Alerts** — Check the **Reorder Alerts** sheet regularly for products that need restocking.
6. **Refresh the Pivot Table** — Right-click the pivot table and select *Refresh* after updating stock data to reflect the latest figures.

---

## Requirements

- **Microsoft Excel 2016 or later** (the `.xlsm` format and VBA macros require a desktop version of Excel)
- Macros must be enabled for the auto-fill functionality to work
- The workbook is **not compatible** with Google Sheets or LibreOffice Calc due to VBA macro dependencies

## Dashboard
<img width="853" height="561" alt="Screenshot 2026-05-19 173804" src="https://github.com/user-attachments/assets/23d9045e-e202-4d0b-b346-040c310dfecc" />

