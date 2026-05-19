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

| Function | Purpose |
|---|---|
| `IF` | Tests whether live stock is below the reorder level and returns `Yes` or `No` for the Reorder Status column |
| `SUMIF` | Calculates live stock by summing all incoming deliveries and subtracting all outgoing sales for a specific Product ID |
| `COUNTIF` | Counts how many unique products belong to each product category in the Summary Analysis |
| **Macro (VBA)** | Auto-fills Product Name, Supplier, and Warehouse Location as soon as a Product ID is entered in the Stock In or Stock Out sheets |

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
