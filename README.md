# SLA & KPI Reporting Automation

An end-to-end data engineering and business reporting project designed to automate daily, weekly, and monthly Service Level Agreement (SLA) and Key Performance Indicator (KPI) reporting for a Service Operations / IT Support organization.

---

## 🎯 Project Overview & Business Problem

In service operations (IT Service Desk, Customer Operations, Technical Support), managing ticket turnaround time is critical for customer satisfaction and operational governance. Service Level Agreements (SLAs) define target resolution timelines based on ticket severity (P1-P4).

### Key Operational Challenges Addressed:
1. **Manual Reporting Overhead**: Analysts spend hours manually copying CSV data into Excel spreadsheets every week.
2. **Data Quality Flaws**: Raw ticketing export logs contain duplicate rows, missing resolution timestamps, invalid dates, and inconsistent category text formatting.
3. **Lack of Automated Governance**: Executive leadership needs real-time visibility into overall SLA compliance %, ticket volume trends, open backlogs, and agent performance.

This project automates the entire reporting workflow—from raw data ingestion and Pandas cleaning to SQLite analytical query execution, dynamic Excel report generation (`SLA_KPI_Report.xlsx`), and weekly summary exports.

---

## 🛠 Tech Stack

- **Data Processing**: Python (`pandas`, `numpy`)
- **Database Layer**: SQLite (`sqlite3`), SQL (CTEs, `GROUP BY`, `CASE WHEN`, `LAG()` Window Function)
- **Excel Reporting**: `openpyxl`, Excel Formulas (`COUNTIFS`, `AVERAGEIFS`, `IFERROR`), Conditional Formatting, Native Excel Charts
- **Pipeline Orchestration**: Python (`subprocess`, `logging`, `time`)

---

## 📁 Folder Structure

```
sla_reporting_project/
├── data/
│   ├── raw_tickets.csv           # Raw synthetic export with 5,000 tickets & data quality issues
│   ├── cleaned_tickets.csv       # Audited, standardized, and clean dataset
│   └── sla_reporting.db          # SQLite Database storing clean tickets table
├── scripts/
│   ├── 01_generate_data.py       # Data generator with fixed seed & SLA rules
│   ├── 02_clean_data.py          # Data cleaning pipeline & audit logging
│   ├── 03_load_to_sql.py         # SQLite loader & SQL analytical engine
│   ├── 04_build_reports.py       # Excel builder (7 worksheets, formulas, charts)
│   └── run_all.py                # Master orchestrator pipeline script
├── sql/
│   └── queries.sql               # 7 analytical SQL queries with detailed comments
├── output/
│   ├── SLA_KPI_Report.xlsx       # Executive Excel report with dynamic formulas & charts
│   ├── weekly_summary.csv        # Weekly operational rollup CSV
│   ├── sql_kpi_summary.csv       # Consolidated SQL KPI summary
│   └── sql_results/              # Individual CSV outputs for each SQL query
├── README.md                     # Project documentation & business insights
└── INTERVIEW_QA.md               # 15 technical & business interview questions & answers
```

---

## 📊 Business Rules & KPI Definitions

### Standard SLA Definition Rule
- **SLA Compliance % Formula**:  
  $$\text{SLA Compliance \%} = \frac{\text{Tickets Met SLA}}{\text{Total Resolved Tickets}} \times 100$$
  *(Where Total Resolved Tickets includes tickets in 'Resolved' or 'Closed' status. Open and Escalated backlog tickets are excluded from SLA resolution percentages since their resolution time is not yet finalized).*

### 1. SLA Targets by Priority
- **P1 (Critical)**: Resolution Target = **4 Hours**
- **P2 (High)**: Resolution Target = **8 Hours**
- **P3 (Medium)**: Resolution Target = **24 Hours**
- **P4 (Low)**: Resolution Target = **48 Hours**

### 2. Primary Key Performance Indicators (KPIs)
- **Total Ticket Intake Volume**: Count of all support tickets logged (5,000 tickets).
- **Total Resolved Tickets**: Count of completed tickets (4,438 tickets).
- **SLA Compliance %**: `(3,629 Tickets Met SLA / 4,438 Resolved Tickets) * 100 = 81.77%`
- **Average Resolution Time (Hours)**: Mean turnaround time from creation to resolution timestamp (20.35 hours).
- **Open Backlog Count**: Count of active tickets in `Open` or `Escalated` status requiring action (562 tickets).
- **Average Customer Satisfaction (CSAT)**: Mean customer rating score on a 1–5 scale (4.00 / 5.00).
- **SLA Breach Count**: Count of resolved tickets exceeding target hours (809 tickets).

---

## 🔍 Key Data Insights (Extracted from Pipeline Execution)

Based on the automated execution of our 5,000 ticket dataset across 2024:

1. **Overall SLA Compliance**: The service desk achieved an overall SLA compliance rate of **81.77%** across **4,438 resolved tickets** (3,629 met SLA out of 4,438 resolved), meeting the benchmark governance threshold (>80%).
2. **P1 Critical Ticket Risk**: Critical **P1 tickets** experienced the lowest SLA compliance at **78.29%** (339 met SLA out of 433 resolved), breaching the 80% governance threshold, whereas P2 (82.83%), P3 (81.83%), and P4 (82.00%) maintained steady adherence.
3. **Category Bottlenecks**: **Network** issues took the longest average resolution time (**21.29 hours**), followed by **Hardware** (**20.85 hours**) and **Software** (**20.53 hours**), indicating infrastructure dependency delays.
4. **Volume Volatility**: Ticket intake peaked in **August 2024 (443 tickets)** with SLA compliance reaching **86.90%**, but experienced a severe **-12.42% intake drop in September**, followed by compliance dips in **October (78.71%)** and **December (77.84%)**.
5. **Top Performing Agent**: **Carlos Mendez (Application Support)** achieved the highest SLA compliance (**89.53%**) across 86 resolved tickets with a CSAT score of **4.16/5.00**.

---

## ⚡ How to Run the Project

### Prerequisites
Ensure Python 3.9+ is installed along with required packages:
```bash
pip install pandas numpy openpyxl matplotlib xlsxwriter
```

### Running the Complete Automated Pipeline
Execute `run_all.py` from the project directory:
```bash
python scripts/run_all.py
```

This single command will sequentially generate raw data, execute cleaning scripts, populate SQLite, run 7 SQL queries, generate the multi-tab Excel dashboard, and output weekly summary files.

---

## 💬 How I Would Explain This Project in an Interview (60-Second Pitch)

> *"I built an automated SLA and KPI reporting pipeline to solve a common operational challenge in service operations: manual Excel reporting on dirty ticket data.*
>
> *I created a synthetic dataset of 5,000 support tickets across 12 months with deliberate flaws like duplicates, text casing errors, missing timestamps, and invalid dates. Using Python and Pandas, I audited and cleaned the data, standardizing dates and categories.*
>
> *Next, I loaded the cleaned data into SQLite and wrote 7 analytical SQL queries using CTEs, CASE WHEN statements, and window functions like `LAG()` to analyze monthly SLA compliance, backlog by team, and MoM volume trends.*
>
> *Finally, I automated an executive 7-sheet Excel report (`SLA_KPI_Report.xlsx`) using `openpyxl`, featuring dynamic Excel formulas like `COUNTIFS` and `AVERAGEIFS` based on the standard SLA definition (Tickets Met SLA / Resolved Tickets = 81.77%), conditional formatting, and embedded charts, along with a master script `run_all.py` that runs the entire workflow end-to-end in under 15 seconds.*
>
> *This project demonstrated how data automation eliminates manual effort, ensures 100% data quality accuracy, and provides business leaders with actionable SLA governance insights."*
