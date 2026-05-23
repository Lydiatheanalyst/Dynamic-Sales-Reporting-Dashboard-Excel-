# Dynamic-Sales-Reporting-Dashboard-Excel-
Dynamic Excel sales dashboard featuring automated KPI reporting, time-intelligence analysis, and real-time daily, monthly, and yearly sales tracking.


# 📊 Dynamic Sales Reporting Dashboard (Excel)

## 🔍 Project Overview
This project is a **dynamic sales reporting dashboard** built entirely in Microsoft Excel using formula-driven automation and time-intelligence analysis.

The objective was to create a reporting system that automatically recalculates KPIs whenever new sales records are entered, eliminating repetitive manual reporting tasks.

Unlike static dashboards that require continuous adjustments, this dashboard dynamically updates sales metrics across daily, monthly, and yearly periods in real time.

--

# 🎯 Business Objective
The dashboard was designed to:

- Automate recurring sales reporting
- Monitor business performance across time periods
- Compare current and historical sales trends
- Reduce manual KPI calculations
- Enable faster business decision-making through real-time reporting

---

# 🗂 Dataset Information

### Dataset Type
Sales Dataset

### Dataset Columns
| Column Name | Description |
|---|---|
| Date | Transaction date |
| CustomerID | Unique customer identifier |
| ProductID | Unique product identifier |
| Category | Product category |
| Quantity | Quantity sold |
| Unit Price | Price per product |
| Sales Amount | Total transaction amount |
| Payment Method | Payment type used |
| Region | Sales region |

### Data Quality
The dataset was already clean and required no preprocessing or transformation before analysis.

---

# ⚙️ Dashboard Features

The dashboard automatically tracks:

- Daily Sales
- Yesterday vs Today Sales
- Current Month Sales
- Previous Month Sales
- Month-on-Month (MoM) Analysis
- Year-to-Date Sales
- Previous Year Sales
- Year-on-Year (YoY) Analysis
- Quantity Sold by Day
- Quantity Sold by Month

---

# 🧠 Time-Intelligence Analysis

This project focuses heavily on **time-intelligence calculations** using dynamic Excel formulas.

The formulas use:
- `TODAY()`
- `SUMIFS()`
- `EOMONTH()`
- `DATE()`
- `YEAR()`

to automatically calculate performance metrics based on the current date.

This means:
> Once new sales records are entered, the dashboard updates automatically without manual recalculation.

---

# 📈 Key Insights

| Metric | Value |
|---|---|
| Current Month Sales | ₦9.74M |
| Current Year Sales | ₦55.98M |
| Month-on-Month Sales Change | -8% |
| Year-on-Year Sales Change | -40% |


---

# 🧮 Excel Formulas Used

## 💰 Sales Amount Calculations (Column I)

### Daily Sales
```excel
=SUMIFS($I$2:$I$327, $B$2:$B$327, TODAY())
```

### Yesterday’s Sales
```excel
=SUMIFS($I$2:$I$327, $B$2:$B$327, TODAY()-1)
```

### Current Month Sales
```excel
=SUMIFS($I$2:$I$327, $B$2:$B$327, ">=" & EOMONTH(TODAY(),-1)+1, $B$2:$B$327, "<=" & TODAY())
```

### Previous Month Sales
```excel
=SUMIFS($I$2:$I$327, $B$2:$B$327, ">=" & EOMONTH(TODAY(),-2)+1, $B$2:$B$327, "<=" & EOMONTH(TODAY(),-1))
```

### Year-to-Date Sales
```excel
=SUMIFS($I$2:$I$327, $B$2:$B$327, ">=" & DATE(YEAR(TODAY()),1,1), $B$2:$B$327, "<=" & TODAY())
```

### Previous Year Sales
```excel
=SUMIFS($I$2:$I$327, $B$2:$B$327, ">=" & DATE(YEAR(TODAY())-1,1,1), $B$2:$B$327, "<=" & DATE(YEAR(TODAY())-1,12,31))
```

---

# 📦 Quantity Calculations (Column F)

### Quantity Sold Today
```excel
=SUMIFS($F$2:$F$327, $B$2:$B$327, TODAY())
```

### Quantity Sold Yesterday
```excel
=SUMIFS($F$2:$F$327, $B$2:$B$327, TODAY()-1)
```

### Quantity Sold This Month
```excel
=SUMIFS($F$2:$F$327, $B$2:$B$327, ">=" & EOMONTH(TODAY(),-1)+1, $B$2:$B$327, "<=" & TODAY())
```

### Quantity Sold Previous Month
```excel
=SUMIFS($F$2:$F$327, $B$2:$B$327, ">=" & EOMONTH(TODAY(),-2)+1, $B$2:$B$327, "<=" & EOMONTH(TODAY(),-1))
```

### Quantity Sold Year-to-Date
```excel
=SUMIFS($F$2:$F$327, $B$2:$B$327, ">=" & DATE(YEAR(TODAY()),1,1), $B$2:$B$327, "<=" & TODAY())
```

### Quantity Sold Previous Year
```excel
=SUMIFS($F$2:$F$327, $B$2:$B$327, ">=" & DATE(YEAR(TODAY())-1,1,1), $B$2:$B$327, "<=" & DATE(YEAR(TODAY())-1,12,31))
```

---

# 🛠 Tools Used

- Microsoft Excel
- Formula-Based Automation
- Time-Intelligence Analysis
- KPI Reporting
- Sales Analytics

---

# 📚 Key Learnings

Through this project, I learned:

- How to build automated reporting systems in Excel
- How time-intelligence formulas improve reporting efficiency
- How to structure KPI-driven dashboards
- How dynamic calculations reduce repetitive manual work
- How businesses can monitor performance trends in real time

---


# 👤 Author

**Lydia Azide**  
Operations | Data Analyst | Building Scalable Systems with Data & Automation



