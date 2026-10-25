# Budgetary-Analysis
# 📊 Financial Budget vs. Actual Cost Analysis & Variance Reporting

[![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-150458?logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![Power BI](https://img.shields.io/badge/Power_BI-Interactive_Dashboard-F2C811?logo=powerbi&logoColor=black)](https://powerbi.microsoft.com/)
[![Excel](https://img.shields.io/badge/Microsoft_Excel-Data_Cleaning_&_Formulas-217346?logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

---

## 📌 Executive Summary

This project presents an end-to-end **Corporate Financial Planning & Analysis (FP&A)** project evaluating the fiscal year 2023 budget against actual business expenditures.

By combining **Microsoft Excel** for data cleaning and error-handled variance calculations (`IFERROR`), **Python (Pandas)** for automated exploratory data analysis (EDA) and numerical verification, and **Microsoft Power BI** for executive-level dashboard reporting, this repository delivers clear visibility into cost overruns, cost savings, and budget allocation efficiencies across operational departments.

---

## 📑 Table of Contents

- [Project Overview](#-project-overview)
- [Key Financial Metrics & KPIs](#-key-financial-metrics--kpis)
- [Tech Stack](#-tech-stack)
- [Data Pipeline & Workflow](#-data-pipeline--workflow)
  - [1. Excel Data Cleaning & IFERROR Variance Formulas](#1-excel-data-cleaning--iferror-variance-formulas)
  - [2. Python Exploratory Data Analysis](#2-python-exploratory-data-analysis)
  - [3. Power BI Dashboards & Visualization](#3-power-bi-dashboards--visualization)
- [Power BI Dashboard Walkthrough](#-power-bi-dashboard-walkthrough)
- [Business Insights & Financial Findings](#-business-insights--financial-findings)
- [Project Directory Structure](#-project-directory-structure)
- [How to Run / Reproduce](#-how-to-run--reproduce)
- [Author & Acknowledgments](#-author--acknowledgments)

---

## 🎯 Project Overview

Managing corporate expenses requires continuous tracking of actual costs against approved budgets. This project analyzes a 12-month expense lifecycle across six major cost centers:

1. **Development Costs** (Core product engineering & software development)
2. **Operational Costs** (Daily facility, utilities, and backend business operations)
3. **Marketing Costs** (Acquisition campaigns, lead generation, and branding)
4. **Travelling Cost** (Onsite client visits, flights, and travel per diems)
5. **Training Cost** (Upskilling programs, certifications, and workshops)
6. **Maintenance Cost** (Software licenses, IT infrastructure, and repairs)

### Primary Objectives:
- Identify monthly and department-level expenditure deviations.
- Clean and normalize raw financial tables handling zero-division risks through robust Excel formulas.
- Automate variance computation and aggregation using Python Pandas.
- Build an interactive Power BI dashboard featuring dynamic slicers, KPI tracking, and proportion benchmarking.

---

## 📈 Key Financial Metrics & KPIs

| Metric | Budgeted (FY23) | Actual Incurred (FY23) | Variance (Actual - Budget) | Status |
| :--- | :---: | :---: | :---: | :---: |
| **Total Expenditure** | **$1,762,200** | **$1,718,180** | **-$44,020** (-2.50%) | 🟢 **Favorable (Under Budget)** |
| **Development Costs** | $1,340,000 (76.05%) | $1,315,100 (76.51%) | -$24,900 (-1.86%) | 🟢 Favorable |
| **Operational Costs** | $165,000 (9.36%) | $174,191 (10.13%) | +$9,191 (+5.57%) | 🔴 Unfavorable (Over Budget) |
| **Training Cost** | $86,000 (4.88%) | $69,500 (4.06%) | -$16,500 (-19.19%) | 🟢 Favorable (Under-spent) |
| **Marketing Costs** | $75,000 (4.26%) | $68,076 (3.96%) | -$6,924 (-9.23%) | 🟢 Favorable |
| **Maintenance Cost** | $48,000 (2.72%) | $41,288 (2.40%) | -$6,712 (-13.98%) | 🟢 Favorable |
| **Travelling Cost** | $48,200 (2.73%) | $50,025 (2.91%) | +$1,825 (+3.79%) | 🟡 Slight Overrun |

---

## 🛠️ Tech Stack

- **Microsoft Excel**:
  - Raw table restructuring and comparative layout design.
  - Calculated fields for `Total Monthly Budget Expenses` and `Total Monthly Actual Expenses`.
  - Advanced error-trapping formula design: `IFERROR(...)` for clean division and null-proof variance computation.
- **Python 3.9+**:
  - `pandas`: Data ingestion, slicing (`head` / `tail`), matrix aggregation, difference and sample variance calculations.
  - `matplotlib`: Baseline visual exploration and sanity checking.
- **Power BI Desktop**:
  - Data Modeling, DAX measures, KPI scorecards.
  - Slicers, Donut Charts, Clustered Bar/Column Charts, and Area Variance Trendlines.

---

## 🔄 Data Pipeline & Workflow

```mermaid
flowchart LR
    A[Raw Excel Data\nFinancial analysis_Data Set.xlsx] --> B[Excel Data Cleaning\n& IFERROR Formulas]
    B --> C[Python Pandas\nVerification & EDA]
    B --> D[Power BI Desktop\nData Model & Visuals]
    C --> E[Variance & Difference Metrics]
    D --> F[Executive Interactive Dashboard]
```

### 1. Excel Data Cleaning & IFERROR Variance Formulas

The raw dataset stored 24 records (12 months of Budget followed by 12 months of Actual expenses) stacked in a single vertical table.

#### Cleaning Steps:
1. **Separation & Restructuring**: Split records into distinct **Budget** and **Actual** structures side-by-side for month-to-month alignment.
2. **Total Calculated Columns**:
   ```excel
   =SUM(B2:G2)
   ```
   Computed the total expenses for each month across the 6 cost categories.
3. **Monthly & Yearly Variance Calculation with `IFERROR`**:
   To prevent runtime formula errors like `#DIV/0!` when a cost category had zero allocations (e.g., Training Cost in March was 0), the calculations were guarded using `IFERROR`:
   
   - **Monthly Variance % Formula**:
     ```excel
     =IFERROR((Actual_Expense - Budget_Expense) / Budget_Expense, 0)
     ```
   - **Yearly / Category Variance % Formula**:
     ```excel
     =IFERROR((SUM(Actual_Range) - SUM(Budget_Range)) / SUM(Budget_Range), 0)
     ```
   - **Variance Difference**:
     ```excel
     =IFERROR(Actual_Expense - Budget_Expense, 0)
     ```

This ensured clean, professional tables ready for Power BI consumption without dirty data or missing value propagation.

---

### 2. Python Exploratory Data Analysis

A Python script was developed to validate Excel's totals, slice subsets, and calculate statistical variances programmatically:

```python
import pandas as pd
import matplotlib.pyplot as plt

# 1. Load the cleaned financial dataset
f = pd.read_excel("Financial analysis_Data Set.xlsx")

# 2. Select relevant financial attributes
d = f[[
    "Month", "Development Costs", "Operatinal Costs", "Marketing Costs",
    "Travelling Cost", "Training Cost", "Maintenance Cost", "Expenses"
]]

# 3. Segregate into Budget (first 12 rows) and Actual (last 12 rows)
b = d.head(12)  # Budget Table
a = d.tail(12)  # Actual Table

# 4. Sum expenses across categories
cost_cols = [
    "Development Costs", "Operatinal Costs", "Marketing Costs",
    "Travelling Cost", "Training Cost", "Maintenance Cost"
]
bc = b[cost_cols].sum()  # Total Budget per category
ac = a[cost_cols].sum()  # Total Actual per category

# 5. Compute Dollar Differences & Variance
diff = ac - bc              # Difference per cost center (Actual - Budget)
total_diff = diff.sum()     # Net FY difference (-$44,020)
total_variance = diff.var() # Cross-category variance
```

---

### 3. Power BI Dashboards & Visualization

The Power BI solution is structured into two main analytical views:

#### 📊 View 1: Expense Proportions & Monthly Breakdown
- **Slicers**: Toggle between `All`, `Budget`, and `Actual` views.
- **Proportions of Budget (Pie Chart)**:
  - Highlights baseline budget weighting (Development dominates at **76.05%**, followed by Operational at **9.36%**).
- **Proportions of Actual Expenses (Donut Chart)**:
  - Realized distribution (Development **76.51%**, Operational **10.13%**, Training **4.06%**).
- **Monthly Actual vs Budget Ranking (Horizontal Bar Charts)**:
  - Instant monthly audit showing peak spend in November ($151,154) and December ($151,290).

#### 📈 View 2: Executive Budget vs. Actual & Variance Trends
- **KPI Headline Cards**:
  - `Total Budget`: **1,762,200**
  - `Actual Expenses`: **1,718,180**
  - `Variance Index`: **0.33**
- **Budget vs. Actual Clustered Column Chart**:
  - Side-by-side cost center benchmarking.
- **Sum of Variance by Expenses (Pie Chart)**:
  - Shows where variance originated (Training Cost contributed **45.12%** of variance deviation, followed by Marketing at **30.09%**).
- **Monthly Difference & Variance Area Chart**:
  - Displays monthly spending trajectory with a noticeable surplus peak in July (+$15.2K saved) and minor overspending in April (-$5.9K).

---

## 🖼️ Dashboard Preview

*(Place your exported Power BI screenshots in an `assets/` or `images/` folder in your repo)*


| Dashboard View 1: Executive Variance Analysis |
<img width="1326" height="741" alt="Screenshot 2026-10-16 130133" src="https://github.com/user-attachments/assets/25ccb15c-94ab-41ba-a5b3-221000cb8179" /> 

| Dashboard View 2: Proportions & Monthly Totals |
<img width="1322" height="742" alt="Screenshot 2026-10-16 130158" src="https://github.com/user-attachments/assets/dc159287-e2c6-42c7-be37-7d8e6fed9ded" />

| Raw Data Structure | Cleaned Data & Calculated Columns |
| :---: | :---: |
| ![Raw Excel Data] <img width="1027" height="675" alt="Screenshot 2026-10-25 100749" src="https://github.com/user-attachments/assets/3090a241-0de4-40e9-a8be-9cd3469b8e78" />
| ![Cleaned Excel Data] <img width="1662" height="222" alt="Screenshot 2026-10-25 101013" src="https://github.com/user-attachments/assets/f863c96f-e739-47a8-9ddd-a93b13f6f06f" />
<img width="1322" height="442" alt="Screenshot 2026-10-25 101048" src="https://github.com/user-attachments/assets/94310f53-bfef-4dd3-8797-00b9ea0f8525" />
<img width="622" height="377" alt="Screenshot 2026-10-25 101103" src="https://github.com/user-attachments/assets/c84e6206-6ed5-4b5f-8b89-4f0edda2398b" />





---

## 💡 Business Insights & Financial Findings

1. **Overall Fiscal Health (Under Budget)**:
   -Figuratively **$44020** have been saved from the overall budget of **$1762200**, if we compare budget with actual expenses of **$1718180**, we can visually conclude that the actual expenses were well aligned and within the budget but if look at the the monthly allocations and variations, there are points that are needed to be considered.
   
2.**Development Costs**:
   -Yearly budget was around **$134200** where as actual expenses were around **$1341365** making a variance of around **2%** which is well versed with the allotted budget, with a slight positive variation. 
   
3. **Operational Overrun Alert**:
    -Total yearly budget allocated for operational costs was **$16500** but actual expenses exceeded by a difference of **-9082** i.e. actual expenses were **$174082** that leads to a variance of **-6%**. which makes efficient allocation of budget considerable for operational costs.
   
4. **Training, Travelling & Marketing (Under-utilization)**:
   -Budgeted costs for Marketing, Training and Travelling expenses were heavily under utilized reportedly having a variance that went up to **21%** for Training Costs which makes the allocation of budget highly considerable, For the following Expenses the difference and variance between actual and allocated budget were as follows:
   Marketing Costs Budgeted=**$75000**, Actual=**$5824** yearly, making a variance of **7%**
   Training Costs Budgeted=**$86000**, Actual=**$68000** yearly, variance = **21%**
   Travelling Costs Budgeted=**$48000**, Actual=**$41300** yearly making a variance of **14%**
   whereas allocation of maintenance costs suffered a variance of **-5%** when actual expenses exceeded the budget with **-$2487** second to operational costs which was budgeted around **$48000** for the year but actual expenses reached up to **$50487**.
    
5. **Overall Spending Dynamics**:
   -For further analysis monthly difference between budgeted and actual costs have been calculated as per monthly basis, so that the allocations can be reviewed as per monthly variations.
-As for monthly variances we can see that for the months of **July and November** the variance steep was on the two extremes one lowest and other highest for the month of July the variance shot up to a positive **10.5%** indicating high under utilization of budgets where as for the month of November the variance was at the lowest extreme a negative **-0.10%** variance FY23, Hence to conclude the yearly budget might indicate a considerable saving for the year but efficient and detailed  allocation is still crucial.
---

## 📂 Project Directory Structure

```plaintext
├── data/
│   └── Financial analysis_Data Set.xlsx    # Cleaned & original dataset with IFERROR formulas
├── notebooks/
│   └── financial_analysis.py               # Python verification and data manipulation script
├── power_bi/
│   └── Financial_Analysis_Report.pbix      # Interactive Power BI dashboard file
├── images/
│   ├── dashboard_proportions.png           # Power BI dashboard page 1
│   ├── dashboard_variance_analysis.png     # Power BI dashboard page 2
│   ├── raw_excel_data.png                  # Raw dataset preview
│   └── cleaned_excel_data.png              # Cleaned dataset with calculated fields
├── README.md                               # Project documentation & execution guide
└── requirements.txt                        # Python library dependencies
```

---

## 🚀 How to Run / Reproduce

### Prerequisites
- Python 3.8+
- Microsoft Excel 2016 or newer (or Office 365)
- Power BI Desktop (Free download from Microsoft)

### Step 1: Clone the Repository
```bash
git clone https://github.com/your-username/financial-budget-actual-analysis.git
cd financial-budget-actual-analysis
```

### Step 2: Set Up Python Environment
```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

pip install -r requirements.txt
```
*(or install directly: `pip install pandas openpyxl matplotlib`)*

### Step 3: Run the Analysis Script
```bash
python notebooks/financial_analysis.py
```

### Step 4: Open Power BI Dashboard
1. Launch **Power BI Desktop**.
2. Open `power_bi/Financial_Analysis_Report.pbix`.
3. If prompted to refresh data sources, select **Transform Data** > **Data source settings** and point to `data/Financial analysis_Data Set.xlsx`.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE) - feel free to use, modify, and distribute for personal and commercial analytics portfolio work.
