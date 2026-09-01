# 📊 E-Commerce Sales & Margin Interactive Dashboard

## 📝 Project Overview

This project is an advanced Excel-based Business Intelligence dashboard built to analyse e-commerce sales performance, profitability, and customer segmentation contained in [Superstore Dataset](https://www.kaggle.com/datasets/ishanshrivastava28/superstore-sales). It demonstrates a full-stack data analysis workflow: from raw **data extraction** and **transformation (ETL)** to relational **data modelling**, **DAX** calculations, and dynamic **visual reporting**.

The main deliverable of this project is the interactive file: `Dashboard.xlsx`

## 🛠️ Tech Stack & Tools Used

* **Microsoft Excel (Advanced).**

* **Power Query:** Data cleaning, transformation, and ETL processes.

* **Power Pivot:** Relational data modeling (Star Schema).

* **DAX (Data Analysis Expressions):** Custom measures for dynamic KPI calculations.

* **Pivot Tables & Pivot Charts:** Data aggregation and visualization.

## 📁 Repository Contents

* **Dashboard.xlsx** — The final interactive dashboard file containing the data model, background calculations, and the UI frontend.

## ⚙️ Project Workflow & Methodology

### 1. Data Extraction & Transformation (Power Query)
   
* Imported raw sales data (.csv) and fixed locale-specific date formatting errors using proper Data Type Locales (e.g., UK/US formats).

* Cleaned string data by extracting text before delimiters to normalize customer names.

### 2. Data Modeling (Power Pivot)

* Loaded the clean data into the Excel Data Model to optimize performance.

* Created a separate Dim_Calendar (Date Table) to enable Time Intelligence analysis.

* Established one-to-many active relationships between the Fact Table (Sales) and Dimension Tables.

### 3. DAX Measures

Instead of relying on basic Pivot Table summarization, I built explicit DAX measures to ensure accurate, dynamic recalculations when filters are applied:

`Total Revenue := SUM([Sales])`

`Total Profit := SUM([Net Profit])`

`Profit Margin % := DIVIDE([Total Profit], [Total Revenue], 0)`

`Transactions := COUNTROWS('Sales')`

### 4. Data Visualization & UI Design

* **KPI Cards:** Custom shapes linked directly to DAX measures to display high-level metrics instantly.

* **Trend Analysis:** A Combo Chart showing Revenue (Columns) vs. Profit (Line on a secondary axis) over time.

* **Product Performance:** A horizontal Bar Chart with an automated Top-10 filter to highlight the most profitable sub-categories.

* **Interactivity:** Implemented Slicers (connected to all pivot tables via Report Connections) with unlocked object properties, allowing users to filter the entire dashboard by Year, Region, and Category seamlessly.

## 🚀 How to Use the Dashboard

1. Download the **Dashboard.xlsx** file from this repository.

2. Open the file in **Microsoft Excel** (Windows Desktop version recommended for full Power Pivot compatibility).

3. If prompted, click *"Enable Editing"* and *"Enable Content"* to activate the data model.

4. Use the **Slicers** on the main Dashboard sheet to interact with the data. All charts and KPI cards will update dynamically.

(Note: The raw data and Pivot Tables are safely hidden on a separate background sheet to keep the UI clean).
