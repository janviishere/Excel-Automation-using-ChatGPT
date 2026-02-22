# Excel-Automation-using-ChatGPT

# 📊 Pizza Sales Excel Automation using Python (Colab)

## 🚀 Project Overview

This project automates pizza sales data analysis using **Python (Pandas)** in **Google Colab**.

It reads a CSV file (`pizza_sales.csv`) and automatically:

* Calculates Total Sales
* Calculates Average Sales
* Finds Top Selling Pizza
* Generates Category & Size Reports
* Creates Monthly Sales Analysis
* Exports a Multi-Sheet Excel Report

---

## 📁 Dataset

The dataset contains pizza sales information such as:

* `order_id`
* `pizza_name`
* `pizza_category`
* `pizza_size`
* `quantity`
* `order_date`
* `total_price`

---

##  Technologies Used

* Python
* Pandas
* OpenPyXL
* Google Colab
* Excel Automation

---

##  Installation

If running locally:

```bash
pip install pandas openpyxl
```

If using Google Colab:

```python
!pip install pandas openpyxl
```

---

##  How It Works

1. Upload `pizza_sales.csv`
2. Script reads and processes data
3. Performs analysis
4. Generates `Pizza_Sales_Automation_Report.xlsx`

---

##  Features Implemented

### 1. Total & Average Sales

Calculates overall revenue and average order value.

### 2. Top Selling Pizza

Identifies pizza with highest total revenue.

### 3. Sales by Category

Aggregates revenue by pizza category.

### 4. Sales by Size

Analyzes performance by pizza size.

### 5. Monthly Sales Report

Groups revenue month-wise.

### 6. Multi-Sheet Excel Export

Automatically creates:

* Raw Data Sheet
* Summary Sheet
* Top Pizza Report
* Category Report
* Size Report
* Monthly Sales Report

---

##  Output

The script generates:

```
Pizza_Sales_Automation_Report.xlsx
```

With multiple analytical sheets inside.

---

##  Project Purpose

This project demonstrates:

* Excel Automation using Python
* Data Cleaning & Processing
* Business Sales Analysis
* Report Generation
* Automation using ChatGPT assistance

---

##  Future Improvements

* Add interactive dashboard
* Add visual charts in Excel
* Add KPI metrics (growth %, profit margin)
* Automate daily report generation
* Deploy as web app

