# 👥 HR Analytics — Employee Attendance & Presence Insights

[![Header](https://capsule-render.vercel.app/api?type=waving&color=0:667eea,100:764ba2&height=200&section=header&text=HR%20Analytics%20Dashboard&fontSize=40&fontColor=ffffff&animation=fadeIn&fontAlignY=35&desc=Employee%20Attendance%20%7C%20WFH%20%7C%20Sick%20Leave%20Analysis%20%7C%20Power%20BI&descAlignY=58&descSize=16)](https://github.com/PatlaveetiJabeer786/HR-Analytics-project)

![Power BI](https://img.shields.io/badge/Power%20BI-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![Power Query](https://img.shields.io/badge/Power%20Query-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)
![DAX](https://img.shields.io/badge/DAX-7B2D8B?style=for-the-badge&logo=microsoft&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Domain](https://img.shields.io/badge/Domain-Human%20Resources-blueviolet?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed%20✅-brightgreen?style=for-the-badge)

---

## 📌 Project Overview

This is a **real-world HR Analytics project** built for **Atliq Technologies** — analyzing 3 months of employee attendance data (April, May & June 2022) to give the HR team clear, actionable visibility into **who is coming to office, who is working from home, and who is taking sick leave — and when.**

The project goes beyond a simple attendance sheet. Using **Power Query for ETL** and **DAX for intelligent measures**, I built an interactive Power BI dashboard that helps HR managers spot trends, plan resources, and make data-backed workforce decisions — in real time.

---

## 🧩 Business Problem

> *"Atliq Technologies' HR team was managing attendance data in messy Excel sheets with attendance codes like 'WFH', 'SL', 'HPL' — spread across 3 months of separate tabs. They had no way to see patterns, no trend charts, and no ability to quickly answer: Are more people working from home on Fridays? Is sick leave rising? Which employees have the lowest attendance?"*

**The HR team was struggling with:**

- 📋 Attendance data spread across **3 separate monthly Excel tabs** — no unified view
- 🔤 Raw attendance codes (`WFH`, `SL`, `HPL`, `HWFH`) that needed decoding and calculation
- 📉 **No visibility into trends** — couldn't tell if WFH % was rising or if sick leave was seasonal
- 👤 **No employee-level drill-down** — hard to identify chronic absentees or remote-work patterns
- 📅 **No day-of-week analysis** — didn't know which days had lowest in-office presence
- 🏢 **Hybrid work planning was guesswork** — no data on when the office was actually full

---

## 🎯 My Task as the Data Analyst

I was given the raw attendance Excel file and asked to deliver a full analytics solution:

| Task | What I Did | Tool |
|------|-----------|------|
| **Data Understanding** | Decoded all attendance keys (WFH, SL, PL, HPL, etc.) | Excel |
| **ETL — Data Cleaning** | Unpivoted messy date columns, standardized formats | Power Query |
| **Multi-Month Merge** | Combined April, May & June tabs into one unified table | Power Query |
| **DAX Measures** | Built Presence %, WFH %, SL % — dynamic & filterable | DAX |
| **Trend Analysis** | Daily and weekly trend lines for all 3 key metrics | Power BI |
| **Employee Table** | Individual-level breakdown of all attendance metrics | Power BI |
| **Day-of-Week Analysis** | Identified which weekdays had most WFH / sick leave | DAX + Power BI |
| **Dashboard Delivery** | Single-page interactive dashboard for HR decision-making | Power BI |

---

## 📊 Dashboard Preview

![HR Analytics Dashboard](Hr_Analyics_Dash_Board.png)

*👆 Interactive Power BI Dashboard — Attendance %, WFH %, and Sick Leave % tracked daily with employee-level breakdown*

---

## 📈 Key Performance Indicators (KPIs)

| KPI | Value | What It Tells HR |
|-----|-------|-----------------|
| ✅ **Attendance %** | **94.0%** | Overall workforce in-office presence rate |
| 🏠 **WFH %** | **9.1%** | Percentage of working days done remotely |
| 🤒 **Sick Leave %** | **0.4%** | Percentage of days lost to sick leave |
| 📅 **Period Covered** | **Apr–Jun 2022** | 3 months of daily attendance tracking |

---

## 🔍 My Workflow — Step by Step

### ✅ Step 1 — Understanding the Attendance Key System

Before any analysis, I decoded all attendance codes from the raw Excel file:

| Code | Meaning | How I Counted It |
|------|---------|-----------------|
| `P` | Present (Office) | Full day = 1 |
| `WFH` | Work From Home | Full day = 1 (WFH Count) |
| `HWFH` | Half-day WFH | Half day = 0.5 |
| `SL` | Sick Leave | Full day = 1 (SL Count) |
| `HSL` | Half-day Sick Leave | Half day = 0.5 |
| `PL` | Paid Leave | Not counted as present |
| `HPL` | Half Paid Leave | Half day absence |
| `WO` | Weekly Off | Excluded from working days |
| `HO` | Holiday | Excluded from working days |

---

### ✅ Step 2 — ETL & Data Transformation (Power Query)

The raw data was a wide Excel table with one column per date — not analysis-friendly.

**What I did in Power Query:**

- **Unpivoted** all date columns into rows — one row per employee per day
- **Combined** April, May, and June tabs into a **single unified table** using `Append Queries`
- Filtered out `WO` (weekends) and `HO` (holidays) from total working days calculation
- Added a **Month column** and **Day of Week column** for trend analysis
- Cleaned employee names and standardized all attendance code values

---

### ✅ Step 3 — DAX Measures I Built

```dax
// ═══════════════════════════════════════════
//         CORE ATTENDANCE MEASURES
// ═══════════════════════════════════════════

// Total Working Days (excluding weekends & holidays)
Total Working Days =
VAR totaldays = COUNT('Final Data'[Value])
VAR nonworkdays =
    CALCULATE(
        COUNT('Final Data'[Value]),
        'Final Data'[Value] IN {"WO","HO","BL","FFL","BRL","ML","HFFL"}
    )
RETURN totaldays - nonworkdays

// Present Days (Office + WFH)
Present Days =
VAR presentdays =
    CALCULATE(
        COUNT('Final Data'[Value]),
        'Final Data'[Value] = "P"
    )
VAR halfdays =
    CALCULATE(
        COUNT('Final Data'[Value]),
        'Final Data'[Value] = "HPL"
    )
RETURN presentdays + (halfdays * 0.5)

// Attendance Percentage
Presence % = DIVIDE([Present Days], [Total Working Days], 0)

// ═══════════════════════════════════════════
//         WFH MEASURES
// ═══════════════════════════════════════════

// WFH Count
WFH Counts =
CALCULATE(
    COUNT('Final Data'[Value]),
    'Final Data'[Value] = "WFH"
) + (
    CALCULATE(
        COUNT('Final Data'[Value]),
        'Final Data'[Value] = "HWFH"
    ) * 0.5
)

// WFH Percentage
WFH % = DIVIDE([WFH Counts], [Present Days], 0)

// ═══════════════════════════════════════════
//         SICK LEAVE MEASURES
// ═══════════════════════════════════════════

// Sick Leave Count
SL Count =
CALCULATE(
    COUNT('Final Data'[Value]),
    'Final Data'[Value] = "SL"
) + (
    CALCULATE(
        COUNT('Final Data'[Value]),
        'Final Data'[Value] = "HSL"
    ) * 0.5
)

// Sick Leave Percentage
SL % = DIVIDE([SL Count], [Total Working Days], 0)
```

---

### ✅ Step 4 — Interactive Dashboard I Built

| Visual | What It Shows | Business Use |
|--------|--------------|-------------|
| 📊 **KPI Cards** | Attendance %, WFH %, SL % at a glance | Quick health-check for HR managers |
| 📈 **Attendance % Over Time** | Daily trend line — Apr to Jun | Spot dips and attendance issues early |
| 🏠 **WFH % Over Time** | Daily trend line for remote work | Understand hybrid work adoption |
| 🤒 **SL % Over Time** | Daily sick leave trend | Detect illness spikes (e.g. flu season) |
| 📋 **Employee Table** | Each employee's Presence %, WFH %, SL % | Individual-level accountability |
| 📅 **Day of Week Analysis** | WFH and SL % by Mon–Fri | Plan office capacity by weekday |
| 📅 **Month-wise Tiles** | Toggle between Apr, May, Jun | Compare months side by side |

---

## 💡 Key Business Insights & Outcomes

### 📅 Day-of-Week Patterns

- **Fridays** have the **highest WFH %** — employees consistently prefer remote Fridays
- **Mondays** show the **highest Sick Leave %** — suggesting extended weekend absences
- **Tuesday and Wednesday** are the most consistent in-office days — best for team meetings

### 📉 Monthly Trends

- **Attendance % dropped slightly in June** compared to April — possible summer holiday effect
- **WFH % increased month-over-month** — remote work adoption growing steadily
- **Sick Leave % spiked in specific weeks** — likely correlating with seasonal illness

### 👤 Employee-Level Insights

- Individual attendance table allows HR to **identify employees with below-threshold presence**
- WFH % per employee helps **flag potential remote-work policy violations**
- Sick leave patterns per person help distinguish **genuine illness from habitual absenteeism**

### 🏢 Office Capacity Planning

- With WFH at **9.1%**, the office is used at ~90% capacity on average
- **Fridays need fewer desk bookings** — data confirms lowest in-office footfall
- **Office infrastructure planning** can now be based on actual attendance data, not assumptions

---

## 📈 Business Value Delivered

| Business Decision | Data-Driven Insight |
|-------------------|--------------------|
| 🏢 **Office Space Planning** | Fridays have fewest people — reduce facility costs on low-attendance days |
| 👥 **Meeting Scheduling** | Tuesday & Wednesday are peak in-office days — ideal for all-hands meetings |
| 🤒 **Wellness Policy** | Monday sick leave spike detected — HR can investigate root cause |
| 📋 **Remote Work Policy** | WFH % trending upward — policy review data now available |
| 👤 **Employee Management** | Individual table identifies attendance outliers for HR follow-up |
| 📆 **Resource Planning** | Monthly trends help forecast headcount availability for project allocation |

---

## 📁 Project Structure

```
HR-Analytics-project/
│
├── data/
│   └── Attendance Sheet 2022-2023_Masked.xlsx   # Raw 3-month attendance data
│
├── powerbi/
│   └── HR_Analytics_Dashboard.pbix              # Power BI dashboard file
│
├── Hr_Analyics_Dash_Board.png                   # Dashboard screenshot
└── README.md
```

---

## 🚀 How to Run This Project

1. **Download Power BI Desktop** — free from [microsoft.com/powerbi](https://github.com/PatlaveetiJabeer786/HR-Analytics-project/blob/main/HR-Analytics-Powet%20BI_Project.pbix)

2. **Clone or download** this repository

3. **Open** `HR_Analytics_Dashboard.pbix` in Power BI Desktop

4. **Update the data source** if prompted:
   - Go to `Home → Transform Data → Data Source Settings`
   - Update the path to point to `data/Attendance Sheet 2022-2023_Masked.xlsx`

5. **Click Refresh** — the dashboard loads with all 3 months of data

6. **Explore** using the month tile slicers (Apr 22 / May 22 / Jun 22) to compare months

---

## 🧠 Key Technical Skills Demonstrated

```
✅  Power Query ETL          — Unpivot, Append, and transform messy Excel data
✅  Multi-tab data merging   — Combined 3 monthly sheets into one unified table
✅  Advanced DAX             — CALCULATE, DIVIDE, VAR/RETURN, conditional logic
✅  Attendance Key Decoding  — Business logic for WFH, SL, HPL, HWFH calculations
✅  Trend Analysis           — Daily & weekly time-series line charts
✅  Employee-level Reporting — Granular individual breakdown with aggregated KPIs
✅  Day-of-Week Analysis     — DATENAME DAX for weekday pattern detection
✅  HR Domain Knowledge      — Understanding real HR metrics and what drives them
```

---

## 🌟 Final Summary

| 🔴 Problem | 🟢 My Solution | 📈 Result |
|-----------|---------------|----------|
| 3 separate messy Excel tabs | Power Query ETL — merged & cleaned | One unified clean dataset |
| Raw attendance codes (WFH, SL...) | DAX logic to decode & calculate | Accurate KPI metrics |
| No trend visibility | Daily line charts over 3 months | Patterns & spikes clearly visible |
| No individual tracking | Employee-level breakdown table | HR can identify outliers |
| Office planning by guesswork | Day-of-week WFH analysis | Data-backed facility decisions |
| Reactive HR decisions | Interactive Power BI dashboard | Proactive, real-time HR insights |

---

## 👨‍💻 About Me

I'm a data analyst passionate about turning raw workforce data into clear, actionable HR intelligence using industry-standard tools.

- 🔗 **LinkedIn:** [Add your LinkedIn URL here]
- 📧 **Email:** Add your email here
- 🌐 **GitHub:** [github.com/PatlaveetiJabeer786](https://github.com/PatlaveetiJabeer786)

---

<div align="center">

⭐ **If this project helped you, please give it a Star!** ⭐

</div>

[![Footer](https://capsule-render.vercel.app/api?type=waving&color=0:667eea,100:764ba2&height=100&section=footer)](https://github.com/PatlaveetiJabeer786/HR-Analytics-project)
