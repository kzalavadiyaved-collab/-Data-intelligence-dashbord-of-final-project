<div align="center">

# 📊 PR. Final Project — "Data Intelligence Dashboard"

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Formulas](https://img.shields.io/badge/Formulas-Dynamic_Arrays-blue?style=for-the-badge)
![Pivot](https://img.shields.io/badge/Pivot-Tables_%26_Charts-orange?style=for-the-badge)
![Dashboard](https://img.shields.io/badge/Dashboard-Interactive-success?style=for-the-badge)

**Every formula explained, with its output shown in a table.**

</div>

---

## 📑 Table of Contents

| # | Section | # | Section |
|:-:|---------|:-:|---------|
| 1 | [Assignment Brief](#1--assignment-brief) | 9 | [Most Frequent Product & Repeating Names](#9--most-frequent-product--repeating-names) |
| 2 | [Sample Dataset](#2--sample-dataset-used-in-this-document) | 10 | [Compare Two Lists](#10--compare-two-lists) |
| 3 | [Date & Time Functions](#3--date--time-functions) | 11 | [Abbreviations using TEXT Functions](#11--abbreviations-using-text-functions) |
| 4 | [FILTER Function](#4--filter-function) | 12 | [Dashboard (Pivot, Charts, Slicers, KPIs)](#12--dashboard) |
| 5 | [Conditional Formatting with Icons](#5--conditional-formatting-with-icons--arrows) | 13 | [Final Report & Storytelling](#13--final-report-with-insights--storytelling) |
| 6 | [Timestamp using NOW](#6--create-timestamp-using-now--iterative-calculation) | 14 | [Deliverables Checklist](#14--deliverables-checklist) |
| 7 | [What-If Analysis & Regression](#7--what-if-analysis--linear-regression) | 15 | [How to Use & Troubleshooting](#15--how-to-use--troubleshooting) |
| 8 | [GroupBy: High-Value Customers](#8--groupby-concept--high-value-customers) | | |

---

## 1. 📋 Assignment Brief

<table>
<tr><th align="center" style="background:#FFF2CC">PR. Final Project - "Data Intelligence Dashboard"</th></tr>
<tr><td>

**Dataset:** Use the *'Final Project Dataset'* sheet containing 200+ records of customer transactions.

### 🎯 Objective
Create a full-fledged analysis and reporting solution using Excel to analyze customer transactions, find insights, and build visualizations and dashboards.

### 🧩 Key Topics to Implement
- Date & Time Functions: TODAY, NOW, DATEDIF, EOMONTH
- FILTER Function for multi-value returns
- Conditional Formatting with Icons/Arrows
- Create Timestamp (using NOW and iterative calculations)
- WHAT-IF Analysis: Scenario Manager, Goal Seek
- Linear Regression using Data Analysis Toolpak
- Data Analysis & Storytelling from the dataset
- Use GroupBy concepts to find High-Value Customers using formulas
- Create Dashboard using:
  - Pivot Tables
  - Pivot Charts (Bar, Line, Pie)
  - Slicers and Timelines
  - KPI Indicators with Conditional Formatting
- Find most frequently purchased product or repeating names using formulas
- Compare two lists and extract matching names in a new column
- Create Abbreviations using TEXT functions
- Build a final report with insights and storytelling

### 📦 Deliverable
- Well-structured workbook with multiple sheets (Raw Data, Analysis, Visualizations, Dashboard)
- Summary Insights with key metrics: Total Revenue, Top Customers, Most Sold Products, Regional Sales
- Use of all required formulas and functions as listed above

</td></tr>
</table>

### 🗂️ Workbook Structure

| # | Sheet | Contains |
|:-:|-------|----------|
| 1 | 📥 **Raw Data** | Excel Table `tblSales` (200+ records) + helper columns |
| 2 | 🧮 **Analysis** | FILTER, GroupBy, frequency, list comparison, timestamp |
| 3 | 🔮 **What-If** | Scenario Manager, Goal Seek, Regression output |
| 4 | 📈 **Visualizations** | Pivot tables and pivot charts |
| 5 | 🖥️ **Dashboard** | KPIs, charts, slicers, timeline |
| 6 | 📝 **Final Report** | Insights and storytelling |

> 💡 **Why it is "dynamic":** The data is converted into an **Excel Table** named `tblSales` (`Ctrl + T`). All formulas use *structured references* like `tblSales[Revenue]`, so they **update automatically** when you add or change records. Pivot tables refresh with **Data → Refresh All**.

---

## 2. 🧾 Sample Dataset Used in This Document

> ⚠️ The real dataset has **200+ records**. To keep every output easy to verify, this document uses a **10-row sample**. Your results will change with the full data, but the formulas stay exactly the same.

**Sheet:** `Raw Data` — **Table name:** `tblSales`

| Transaction ID | Date | Customer | Product | Region | Quantity | Unit Price | Revenue |
|:--:|:--:|----------|---------|:--:|:--:|--:|--:|
| T001 | 05-Jan-2026 | Amit Shah | Laptop | North | 2 | 50,000 | 100,000 |
| T002 | 18-Jan-2026 | Riya Patel | Mouse | West | 5 | 500 | 2,500 |
| T003 | 02-Feb-2026 | Amit Shah | Keyboard | North | 3 | 1,200 | 3,600 |
| T004 | 14-Feb-2026 | Neha Singh | Laptop | East | 1 | 50,000 | 50,000 |
| T005 | 09-Mar-2026 | Riya Patel | Mouse | West | 10 | 500 | 5,000 |
| T006 | 21-Mar-2026 | Karan Mehta | Monitor | South | 2 | 12,000 | 24,000 |
| T007 | 11-Apr-2026 | Neha Singh | Mouse | East | 4 | 500 | 2,000 |
| T008 | 25-Apr-2026 | Amit Shah | Monitor | North | 1 | 12,000 | 12,000 |
| T009 | 07-May-2026 | Karan Mehta | Mouse | South | 6 | 500 | 3,000 |
| T010 | 19-May-2026 | Riya Patel | Laptop | West | 1 | 50,000 | 50,000 |

**Revenue column formula (inside the table):**
```excel
=[@Quantity]*[@[Unit Price]]
```

| Metric | Formula | Output |
|--------|---------|-------:|
| Total Revenue | `=SUM(tblSales[Revenue])` | **₹ 2,52,100** |
| Total Transactions | `=COUNTA(tblSales[Transaction ID])` | **10** |
| Total Units Sold | `=SUM(tblSales[Quantity])` | **35** |

---

## 3. 📅 Date & Time Functions

**Topic:** `TODAY`, `NOW`, `DATEDIF`, `EOMONTH`

### 🧮 Formulas

| Function | Purpose | Formula |
|----------|---------|---------|
| `TODAY()` | Current date (updates daily) | `=TODAY()` |
| `NOW()` | Current date and time | `=NOW()` |
| `DATEDIF()` | Days / months between two dates | `=DATEDIF([@Date],TODAY(),"D")` |
| `EOMONTH()` | Last day of the transaction month | `=EOMONTH([@Date],0)` |

### 📊 Output (as on **29-Oct-2026**)

| Cell | Formula | Output |
|------|---------|--------|
| Today | `=TODAY()` | 29-Oct-2026 |
| Now | `=NOW()` | 29-Oct-2026 10:30 AM *(changes every recalculation)* |

| Transaction | Date | Days Since Purchase `DATEDIF(...,"D")` | Months Since `DATEDIF(...,"M")` | Month End `EOMONTH(...,0)` |
|:--:|:--:|:--:|:--:|:--:|
| T001 | 05-Jan-2026 | 297 | 9 | 31-Jan-2026 |
| T002 | 18-Jan-2026 | 284 | 9 | 31-Jan-2026 |
| T003 | 02-Feb-2026 | 269 | 8 | 28-Feb-2026 |
| T004 | 14-Feb-2026 | 257 | 8 | 28-Feb-2026 |

> 📝 **Note:** `DATEDIF` is a hidden function in Excel. It will not show in the formula suggestions, but it works when typed fully. Units: `"D"` = days, `"M"` = months, `"Y"` = years.

---

## 4. 🔎 FILTER Function

**Topic:** Return multiple values with one formula (dynamic array).

### 🧮 Formula

```excel
=FILTER(tblSales, tblSales[Region]="North", "No records found")
```

| Part | Meaning |
|------|---------|
| `tblSales` | Array to return (all columns) |
| `tblSales[Region]="North"` | Condition |
| `"No records found"` | Shown if nothing matches |

### 📊 Output (Region = North)

| Transaction ID | Date | Customer | Product | Region | Quantity | Unit Price | Revenue |
|:--:|:--:|----------|---------|:--:|:--:|--:|--:|
| T001 | 05-Jan-2026 | Amit Shah | Laptop | North | 2 | 50,000 | 100,000 |
| T003 | 02-Feb-2026 | Amit Shah | Keyboard | North | 3 | 1,200 | 3,600 |
| T008 | 25-Apr-2026 | Amit Shah | Monitor | North | 1 | 12,000 | 12,000 |

### 🔁 Multiple Conditions

| Logic | Formula | Output |
|-------|---------|--------|
| **AND** (Region = West **and** Product = Mouse) | `=FILTER(tblSales, (tblSales[Region]="West")*(tblSales[Product]="Mouse"))` | T002, T005 |
| **OR** (Product = Laptop **or** Monitor) | `=FILTER(tblSales, (tblSales[Product]="Laptop")+(tblSales[Product]="Monitor"))` | T001, T004, T006, T008, T010 |
| **Revenue > 20,000** | `=FILTER(tblSales[[Customer]:[Revenue]], tblSales[Revenue]>20000)` | 4 rows (T001, T004, T006, T010) |

> 💡 **Dynamic tip:** Put the region in a cell (e.g. `B1`) or use a **Data Validation dropdown**, and write `tblSales[Region]=$B$1`. The result changes instantly when you change the dropdown.

---

## 5. 🚦 Conditional Formatting with Icons / Arrows

**Topic:** Show Revenue performance visually with icon sets.

### 🪜 Steps

1. Select the `Revenue` column data.
2. Go to **Home → Conditional Formatting → Icon Sets → 3 Arrows (Colored)**.
3. Open **Manage Rules → Edit Rule**, set the type to **Percent**:

| Icon | Rule | Meaning |
|:--:|------|---------|
| 🟢 ⬆️ | Value ≥ 67 percent | High revenue |
| 🟡 ➡️ | Value ≥ 33 percent | Medium revenue |
| 🔴 ⬇️ | Value < 33 percent | Low revenue |

### 📊 Output

| Transaction | Revenue | Icon | Rating |
|:--:|--:|:--:|:--:|
| T001 | 1,00,000 | 🟢 ⬆️ | High |
| T004 | 50,000 | 🟢 ⬆️ | High |
| T010 | 50,000 | 🟢 ⬆️ | High |
| T006 | 24,000 | 🟡 ➡️ | Medium |
| T008 | 12,000 | 🟡 ➡️ | Medium |
| T005 | 5,000 | 🟡 ➡️ | Medium |
| T003 | 3,600 | 🟡 ➡️ | Medium |
| T009 | 3,000 | 🔴 ⬇️ | Low |
| T002 | 2,500 | 🔴 ⬇️ | Low |
| T007 | 2,000 | 🔴 ⬇️ | Low |

### 🧮 Text version of the rating (helper column)

```excel
=IF([@Revenue]>=PERCENTILE.INC([Revenue],0.67),"High",
 IF([@Revenue]>=PERCENTILE.INC([Revenue],0.33),"Medium","Low"))
```

---

## 6. ⏱️ Create Timestamp using NOW & Iterative Calculation

**Topic:** Record the time **once** when data is entered, and keep it fixed.

### 🪜 Setup

1. **File → Options → Formulas → ✅ Enable iterative calculation** (Max Iterations: `1`).
2. In the `Analysis` sheet, enter this formula in cell `B2` and copy it down:

```excel
=IF(A2<>"", IF(B2="", NOW(), B2), "")
```

| Part | Meaning |
|------|---------|
| `A2<>""` | Checks if an entry was typed in column A |
| `B2=""` | Checks if the timestamp is still empty |
| `NOW()` | Records the current date and time |
| `B2` | Keeps the old value so it does not change |

> The formula refers to its own cell (`B2`), which is a **circular reference**. That is why iterative calculation must be enabled.

### 📊 Output

| Cell A (Entry) | Cell B (Timestamp) | Behavior |
|----------------|--------------------|----------|
| *(empty)* | *(blank)* | No entry, no time |
| Order received | 29-Oct-2026 10:30:12 | Time recorded on entry |
| Payment done | 29-Oct-2026 10:45:40 | Fixed, does **not** change later |

**Format the column:** `Ctrl + 1 → Custom → dd-mmm-yyyy hh:mm:ss`

---

## 7. 🔮 What-If Analysis & Linear Regression

### 7.1 Scenario Manager

**Model (sheet `What-If`):**

| Cell | Label | Value / Formula |
|:--:|-------|-----------------|
| B2 | Laptop Quantity | 4 |
| B3 | Unit Price | 50,000 |
| B4 | **Revenue** | `=B2*B3` |

**Path:** *Data → What-If Analysis → Scenario Manager → Add*
**Changing cells:** `B2:B3`

| Scenario | Quantity (B2) | Unit Price (B3) | Revenue (B4) |
|----------|:--:|--:|--:|
| 🔴 Worst Case | 3 | 45,000 | **1,35,000** |
| 🟡 Expected (Base) | 4 | 50,000 | **2,00,000** |
| 🟢 Best Case | 6 | 55,000 | **3,30,000** |

Click **Summary** to generate the *Scenario Summary* sheet automatically.

### 7.2 Goal Seek

**Question:** *How many laptops must be sold to reach a revenue of ₹3,00,000?*

**Path:** *Data → What-If Analysis → Goal Seek*

| Goal Seek Field | Value |
|-----------------|-------|
| Set cell | `B4` |
| To value | `300000` |
| By changing cell | `B2` |

| | Before | After Goal Seek |
|--|:--:|:--:|
| Quantity (B2) | 4 | **6** |
| Unit Price (B3) | 50,000 | 50,000 |
| Revenue (B4) | 2,00,000 | **3,00,000** ✅ |

### 7.3 Linear Regression (Data Analysis ToolPak)

**Enable:** *File → Options → Add-ins → Manage Excel Add-ins → Go → ✅ Analysis ToolPak*
**Run:** *Data → Data Analysis → Regression*

| Regression Input | Selection |
|------------------|-----------|
| Input Y Range | `Revenue` column |
| Input X Range | `Unit Price` column |
| Labels | ✅ Checked |
| Confidence Level | 95% |
| Output | New Worksheet |

**📊 Output (computed on the 10-row sample; approximate):**

**Regression Statistics**

| Statistic | Value |
|-----------|------:|
| Multiple R | 0.9027 |
| R Square | 0.8149 |
| Adjusted R Square | 0.7918 |
| Standard Error | 14,764 |
| Observations | 10 |

**ANOVA**

| | df | SS | MS | F | Significance F |
|--|:--:|--:|--:|--:|--:|
| Regression | 1 | 7.678 × 10⁹ | 7.678 × 10⁹ | 35.22 | 0.0003 |
| Residual | 8 | 1.744 × 10⁹ | 2.180 × 10⁸ | | |
| Total | 9 | 9.422 × 10⁹ | | | |

**Coefficients**

| | Coefficient | Standard Error | t Stat | P-value |
|--|--:|--:|--:|--:|
| Intercept | 2,442 | 6,043 | 0.40 | 0.697 |
| Unit Price | 1.2849 | 0.2165 | 5.94 | 0.0003 |

**🧠 How to read it**

| Item | Interpretation |
|------|----------------|
| **R² = 0.81** | About 81% of the change in revenue is explained by unit price |
| **Coefficient = 1.28** | On average, each ₹1 rise in unit price goes with ₹1.28 more revenue |
| **Significance F < 0.05** | The model is statistically significant |
| **Equation** | `Revenue = 2,442 + 1.2849 × Unit Price` |

**Prediction with formulas:**

| Purpose | Formula | Output (Price = 20,000) |
|---------|---------|--:|
| Slope | `=SLOPE(tblSales[Revenue],tblSales[Unit Price])` | 1.2849 |
| Intercept | `=INTERCEPT(tblSales[Revenue],tblSales[Unit Price])` | 2,442 |
| R² | `=RSQ(tblSales[Revenue],tblSales[Unit Price])` | 0.8149 |
| Forecast | `=FORECAST.LINEAR(20000,tblSales[Revenue],tblSales[Unit Price])` | ≈ 28,139 |

---

## 8. 👑 GroupBy Concept — High-Value Customers

**Topic:** Group transactions by customer using formulas.

### 🧮 Formulas

| Step | Formula | Spill / Use |
|------|---------|-------------|
| 1. Unique customers | `=UNIQUE(tblSales[Customer])` | Spills down from `A2` |
| 2. Total revenue per customer | `=SUMIFS(tblSales[Revenue],tblSales[Customer],A2)` | Column B |
| 3. Number of orders | `=COUNTIFS(tblSales[Customer],A2)` | Column C |
| 4. Rank | `=RANK.EQ(B2,$B$2:$B$5,0)` | Column D |
| 5. High-value flag | `=IF(B2>=AVERAGE($B$2:$B$5),"⭐ High Value","Regular")` | Column E |

### 📊 Output

| Customer | Total Revenue | Orders | Rank | Customer Type |
|----------|--:|:--:|:--:|:--:|
| Amit Shah | 1,15,600 | 3 | 1 | ⭐ High Value |
| Riya Patel | 57,500 | 3 | 2 | Regular |
| Neha Singh | 52,000 | 2 | 3 | Regular |
| Karan Mehta | 27,000 | 2 | 4 | Regular |

> The average customer revenue is **63,025** (₹2,52,100 ÷ 4). Customers above this average are marked **High Value**.

### 🏆 Top 3 Customers (single dynamic formula)

```excel
=TAKE(SORTBY(UNIQUE(tblSales[Customer]),
      SUMIFS(tblSales[Revenue],tblSales[Customer],UNIQUE(tblSales[Customer])),-1),3)
```

| Rank | Customer |
|:--:|----------|
| 1 | Amit Shah |
| 2 | Riya Patel |
| 3 | Neha Singh |

### 🌍 GroupBy Region and Product

| Group | Formula | Output |
|-------|---------|--------|
| North revenue | `=SUMIFS(tblSales[Revenue],tblSales[Region],"North")` | 1,15,600 |
| West revenue | `=SUMIFS(tblSales[Revenue],tblSales[Region],"West")` | 57,500 |
| East revenue | `=SUMIFS(tblSales[Revenue],tblSales[Region],"East")` | 52,000 |
| South revenue | `=SUMIFS(tblSales[Revenue],tblSales[Region],"South")` | 27,000 |

---

## 9. 🔁 Most Frequently Purchased Product & Repeating Names

### 9.1 Most Frequent Product

```excel
=INDEX(tblSales[Product], MODE(MATCH(tblSales[Product], tblSales[Product], 0)))
```

**How it works**

| Step | What happens | Result |
|:--:|--------------|--------|
| 1 | `MATCH(Product, Product, 0)` gives the first position of each product | 1, 2, 3, 1, 2, 6, 2, 6, 2, 1 |
| 2 | `MODE(...)` picks the most repeated position | 2 |
| 3 | `INDEX(Product, 2)` returns the product at position 2 | **Mouse** |

**Dynamic array alternative (with counts):**

```excel
=SORTBY(UNIQUE(tblSales[Product]), COUNTIF(tblSales[Product],UNIQUE(tblSales[Product])), -1)
```

### 📊 Output — Product Frequency

| Product | Times Purchased `COUNTIF` | Units Sold `SUMIFS` | Revenue |
|---------|:--:|:--:|--:|
| 🥇 **Mouse** | **4** | **25** | 12,500 |
| Laptop | 3 | 4 | 2,00,000 |
| Monitor | 2 | 3 | 36,000 |
| Keyboard | 1 | 3 | 3,600 |

| Question | Formula | Output |
|----------|---------|--------|
| Most frequent product | `=INDEX(tblSales[Product],MODE(MATCH(tblSales[Product],tblSales[Product],0)))` | **Mouse** |
| Times it was purchased | `=COUNTIF(tblSales[Product],"Mouse")` | **4** |
| Units sold of Mouse | `=SUMIFS(tblSales[Quantity],tblSales[Product],"Mouse")` | **25** |

### 9.2 Repeating Customer Names

**Helper column in the table:**

```excel
=IF(COUNTIF(INDEX([Customer],1):[@Customer],[@Customer])>1,"Repeat","New")
```

| Transaction | Customer | Status |
|:--:|----------|:--:|
| T001 | Amit Shah | New |
| T002 | Riya Patel | New |
| T003 | Amit Shah | 🔁 Repeat |
| T004 | Neha Singh | New |
| T005 | Riya Patel | 🔁 Repeat |
| T006 | Karan Mehta | New |
| T007 | Neha Singh | 🔁 Repeat |
| T008 | Amit Shah | 🔁 Repeat |
| T009 | Karan Mehta | 🔁 Repeat |
| T010 | Riya Patel | 🔁 Repeat |

**Customers who purchased more than once:**

```excel
=FILTER(UNIQUE(tblSales[Customer]), COUNTIF(tblSales[Customer],UNIQUE(tblSales[Customer]))>1)
```

| Output |
|--------|
| Amit Shah |
| Riya Patel |
| Neha Singh |
| Karan Mehta |

---

## 10. 🔗 Compare Two Lists

**Goal:** Compare *All Customers* with *Loyalty Members* and extract the matching names in a new column.

| Column A: All Customers (List 1) | Column D: Loyalty Members (List 2) |
|----------|----------|
| Amit Shah | Riya Patel |
| Riya Patel | Karan Mehta |
| Neha Singh | Sunil Rao |
| Karan Mehta | Amit Shah |
| Pooja Verma | |

### 🧮 Formulas (Column B and C)

| Column | Purpose | Formula |
|:--:|---------|---------|
| B | Match status | `=IF(COUNTIF($D$2:$D$5,A2)>0,"✅ Match","❌ No Match")` |
| C | Extract matching name | `=IF(COUNTIF($D$2:$D$5,A2)>0,A2,"")` |
| F | One-formula alternative | `=FILTER(A2:A6, COUNTIF(D2:D5,A2:A6))` |

### 📊 Output

| List 1 Name | Status (B) | Matching Name (C) |
|-------------|:--:|-------------|
| Amit Shah | ✅ Match | Amit Shah |
| Riya Patel | ✅ Match | Riya Patel |
| Neha Singh | ❌ No Match | |
| Karan Mehta | ✅ Match | Karan Mehta |
| Pooja Verma | ❌ No Match | |

**Names only in List 2 (not in customers):**

```excel
=FILTER(D2:D5, COUNTIF(A2:A6,D2:D5)=0)
```

| Output |
|--------|
| Sunil Rao |

---

## 11. 🔤 Abbreviations using TEXT Functions

### 🧮 Formulas

| Purpose | Formula |
|---------|---------|
| Customer initials | `=LEFT([@Customer],1)&MID([@Customer],FIND(" ",[@Customer])+1,1)` |
| Product code | `=UPPER(LEFT([@Product],3))` |
| Month-Year label | `=TEXT([@Date],"mmm-yy")` |
| Formatted revenue | `=TEXT([@Revenue],"₹#,##0")` |
| Order code | `=UPPER(LEFT([@Product],3))&"-"&TEXT([@Quantity],"00")` |

### 📊 Output

| Customer | Initials | Product | Code | Date | Month-Year | Revenue | Formatted | Order Code |
|----------|:--:|---------|:--:|:--:|:--:|--:|--:|:--:|
| Amit Shah | **AS** | Laptop | **LAP** | 05-Jan-2026 | Jan-26 | 100000 | ₹100,000 | LAP-02 |
| Riya Patel | **RP** | Mouse | **MOU** | 18-Jan-2026 | Jan-26 | 2500 | ₹2,500 | MOU-05 |
| Amit Shah | **AS** | Keyboard | **KEY** | 02-Feb-2026 | Feb-26 | 3600 | ₹3,600 | KEY-03 |
| Neha Singh | **NS** | Laptop | **LAP** | 14-Feb-2026 | Feb-26 | 50000 | ₹50,000 | LAP-01 |
| Karan Mehta | **KM** | Monitor | **MON** | 21-Mar-2026 | Mar-26 | 24000 | ₹24,000 | MON-02 |

---

## 12. 🖥️ Dashboard

### 12.1 Pivot Tables (sheet `Visualizations`)

**Create:** *Click inside the table → Insert → PivotTable → New Worksheet*

**📊 Pivot 1 — Revenue by Region** *(Rows: Region, Values: Sum of Revenue)*

| Region | Sum of Revenue | % of Total |
|--------|--:|--:|
| North | 1,15,600 | 45.9% |
| West | 57,500 | 22.8% |
| East | 52,000 | 20.6% |
| South | 27,000 | 10.7% |
| **Grand Total** | **2,52,100** | **100%** |

**📊 Pivot 2 — Revenue by Product** *(Rows: Product, Values: Sum of Revenue)*

| Product | Sum of Revenue | % of Total |
|---------|--:|--:|
| Laptop | 2,00,000 | 79.3% |
| Monitor | 36,000 | 14.3% |
| Mouse | 12,500 | 5.0% |
| Keyboard | 3,600 | 1.4% |
| **Grand Total** | **2,52,100** | **100%** |

**📊 Pivot 3 — Monthly Trend** *(Rows: Date grouped by Months, Values: Sum of Revenue)*

| Month | Sum of Revenue |
|-------|--:|
| Jan | 1,02,500 |
| Feb | 53,600 |
| Mar | 29,000 |
| Apr | 14,000 |
| May | 53,000 |
| **Grand Total** | **2,52,100** |

### 12.2 Pivot Charts

| Chart | Based On | Insert Path | Purpose |
|-------|----------|-------------|---------|
| 📊 **Bar** | Pivot 1 (Region) | *PivotTable Analyze → PivotChart → Clustered Bar* | Compare regional sales |
| 📈 **Line** | Pivot 3 (Month) | *PivotChart → Line with Markers* | Show the monthly trend |
| 🥧 **Pie** | Pivot 2 (Product) | *PivotChart → Pie* | Show product share |

### 12.3 Slicers and Timeline

**Insert:** *PivotTable Analyze → Insert Slicer* (Region, Product, Customer) and *Insert Timeline* (Date).
**Connect all pivots:** *Right-click slicer → Report Connections → ✅ select all pivot tables.*

| Action | Total Revenue | Transactions |
|--------|--:|:--:|
| No filter | 2,52,100 | 10 |
| Slicer: Region = **North** | 1,15,600 | 3 |
| Slicer: Product = **Mouse** | 12,500 | 4 |
| Slicer: Customer = **Riya Patel** | 57,500 | 3 |
| Timeline: **Feb – Mar 2026** | 82,600 | 4 |

### 12.4 KPI Indicators with Conditional Formatting

**KPI formulas (sheet `Dashboard`):**

| KPI | Formula |
|-----|---------|
| Total Revenue | `=SUM(tblSales[Revenue])` |
| Transactions | `=COUNTA(tblSales[Transaction ID])` |
| Unique Customers | `=COUNTA(UNIQUE(tblSales[Customer]))` |
| Average Order Value | `=AVERAGE(tblSales[Revenue])` |
| Achievement % | `=Actual/Target` |

**Conditional formatting rule on Achievement %** (*Home → Conditional Formatting → Icon Sets → 3 Traffic Lights → Number*):

| Icon | Rule |
|:--:|------|
| 🟢 | ≥ 100% |
| 🟡 | ≥ 80% |
| 🔴 | < 80% |

**📊 Output:**

| KPI | Actual | Target | Achievement | Status |
|-----|--:|--:|:--:|:--:|
| 💰 Total Revenue | 2,52,100 | 2,50,000 | 100.8% | 🟢 |
| 🧾 Transactions | 10 | 12 | 83.3% | 🟡 |
| 👤 Unique Customers | 4 | 5 | 80.0% | 🟡 |
| 📦 Average Order Value | 25,210 | 30,000 | 84.0% | 🟡 |

### 12.5 Dashboard Layout

```
┌────────────────────────────────────────────────────────────────┐
│                📊 DATA INTELLIGENCE DASHBOARD                  │
├───────────────┬───────────────┬───────────────┬────────────────┤
│ 💰 REVENUE    │ 🧾 TRANSACTNS │ 👤 CUSTOMERS  │ 📦 AVG ORDER   │
│ 2,52,100 🟢   │ 10 🟡         │ 4 🟡          │ 25,210 🟡      │
├───────────────┴───────┬───────┴───────────────┴────────────────┤
│ 📊 Revenue by Region  │ 📈 Monthly Revenue Trend               │
│      (Bar Chart)      │       (Line Chart)                     │
├───────────────────────┼────────────────────────────────────────┤
│ 🥧 Product Share      │ 🏆 Top 3 Customers                     │
│      (Pie Chart)      │                                        │
├───────────────────────┴────────────────────────────────────────┤
│ 🎛️ Slicers: Region | Product | Customer     ⏱️ Timeline: Date  │
└────────────────────────────────────────────────────────────────┘
```

| Design Element | Choice |
|----------------|--------|
| Colors | Navy `#1F3864`, Blue `#2E75B6`, Orange `#ED7D31` |
| Gridlines | Hidden (*View → uncheck Gridlines*) |
| Font | Calibri / Segoe UI |

---

## 13. 📝 Final Report with Insights & Storytelling

### 📈 Summary Insights

| Metric | Result | Formula |
|--------|-------:|---------|
| 💰 **Total Revenue** | ₹ 2,52,100 | `=SUM(tblSales[Revenue])` |
| 👑 **Top Customer** | Amit Shah (₹1,15,600, 45.9%) | `=INDEX(A2:A5,MATCH(MAX(B2:B5),B2:B5,0))` |
| 🥇 **Most Sold Product (units)** | Mouse (25 units) | `=INDEX(Products,MATCH(MAX(Units),Units,0))` |
| 💎 **Highest Revenue Product** | Laptop (₹2,00,000) | `=INDEX(Products,MATCH(MAX(Rev),Rev,0))` |
| 🌍 **Top Region** | North (₹1,15,600) | `=INDEX(Regions,MATCH(MAX(Rev),Rev,0))` |
| 📅 **Best Month** | January (₹1,02,500) | Pivot 3 |
| 🔁 **Repeat Customers** | 4 of 4 | Section 9.2 |
| 📉 **Regression R²** | 0.81 | `=RSQ(...)` |

### 📖 Data Storytelling

| Stage | Story |
|-------|-------|
| **1. What happened?** | Total revenue is ₹2,52,100 from 10 transactions. January was the strongest month, and revenue is driven mainly by North. |
| **2. Why?** | Laptops give **79.3%** of revenue with only 4 units, while Mouse is sold the most (25 units) but earns just 5%. One customer, Amit Shah, brings **45.9%** of revenue. |
| **3. So what?** | Revenue depends heavily on a few high-priced sales and one customer. Unit price strongly explains revenue (R² = 0.81). |
| **4. Now what?** | Retain high-value customers, upsell accessories with laptops, and grow the East and South regions. |

> ✏️ **For your final submission:** replace the sample outputs in this document with the numbers from your 200+ record dataset. The formulas stay the same.

---

## 14. ✅ Deliverables Checklist

| # | Requirement | Section | Status |
|:-:|-------------|:--:|:--:|
| 1 | Date & Time: TODAY, NOW, DATEDIF, EOMONTH | 3 | ✅ |
| 2 | FILTER function | 4 | ✅ |
| 3 | Conditional Formatting with Icons/Arrows | 5 | ✅ |
| 4 | Timestamp using NOW and iterative calculation | 6 | ✅ |
| 5 | What-If: Scenario Manager, Goal Seek | 7 | ✅ |
| 6 | Linear Regression using Data Analysis ToolPak | 7 | ✅ |
| 7 | GroupBy concept for High-Value Customers | 8 | ✅ |
| 8 | Pivot Tables, Pivot Charts (Bar, Line, Pie) | 12 | ✅ |
| 9 | Slicers and Timelines | 12 | ✅ |
| 10 | KPI Indicators with Conditional Formatting | 12 | ✅ |
| 11 | Most frequent product / repeating names | 9 | ✅ |
| 12 | Compare two lists, extract matching names | 10 | ✅ |
| 13 | Abbreviations using TEXT functions | 11 | ✅ |
| 14 | Final report with insights and storytelling | 13 | ✅ |
| 15 | Multi-sheet workbook (Raw Data, Analysis, Visualizations, Dashboard) | 1 | ✅ |

---

## 15. 🛠️ How to Use & Troubleshooting

### ▶️ How to Use

1. Open `YourName_DataIntelligenceDashboard.xlsx` in **Excel 365 / 2021 or later**.
2. Go to the **Dashboard** sheet and use the **Slicers** and **Timeline** to filter.
3. Add new records at the bottom of `tblSales`. The table grows and all formulas update.
4. Press **Data → Refresh All** to refresh the pivot tables and charts.
5. Keep **Iterative Calculation** enabled for the timestamp feature.

### ❓ Common Errors

| Problem | Cause | Fix |
|---------|-------|-----|
| `#NAME?` on `FILTER` / `UNIQUE` / `TAKE` | Older Excel version | Use Excel 365 / 2021+ |
| `#SPILL!` | Cells below the formula are not empty | Clear the cells in the spill range |
| `#CALC!` in FILTER | No matching data | Add the 3rd argument `"No records"` |
| Timestamp keeps changing | Iterative calculation is off, or the formula is wrong | Re-enable it and recheck the formula |
| Data Analysis not visible | ToolPak not enabled | *File → Options → Add-ins → Analysis ToolPak* |
| Pivot shows old data | Not refreshed | **Data → Refresh All** |
| `DATEDIF` not suggested | Hidden function | Type it fully, it still works |

### 🧰 Tools & Skills

`Excel Tables` • `Structured References` • `Dynamic Arrays` • `Pivot Tables` • `Slicers & Timelines` • `What-If Analysis` • `Regression` • `Data Storytelling`

---

<div align="center">

### 👤 Author

**Your Name** • 📧 your.email@example.com
📅 *Submitted: October 2026*

</div>
