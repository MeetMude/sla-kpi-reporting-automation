# SLA & KPI Reporting Automation - Interview Q&A Guide

Tailored for freshers preparing for **Business Reporting & Governance**, **Measurement & Reporting Associate**, and **Data Analyst** interviews.

---

### Q1: What is a Service Level Agreement (SLA) and why is SLA reporting critical in service operations?
**Answer:**  
An SLA (Service Level Agreement) is a formal commitment between a service provider and a customer that defines expected performance standards, such as ticket resolution target time. For example, P1 critical tickets must be resolved within 4 hours. SLA reporting is critical because it gives management visibility into operational health, highlights bottlenecks, ensures contractual compliance, and helps maintain high customer satisfaction (CSAT).

---

### Q2: How is SLA Compliance % standardly defined in service desk governance?
**Answer:**  
Standard SLA Compliance % is defined as:
$$\text{SLA Compliance \%} = \frac{\text{Tickets Met SLA}}{\text{Total Resolved Tickets}} \times 100$$
Where **Total Resolved Tickets** includes tickets with completed status (`'Resolved'` or `'Closed'`). Active backlog tickets (`'Open'` or `'Escalated'`) are excluded from the resolution percentage denominator because their actual resolution duration is not yet finalized.  
In our dataset, out of 4,438 resolved tickets, 3,629 met SLA, yielding an exact compliance rate of **81.77%**.

---

### Q3: Why did you generate a synthetic dataset with deliberate data quality issues?
**Answer:**  
In real-world business reporting, raw data exported from ticketing systems (like ServiceNow or Jira) is rarely clean. It often contains duplicates, inconsistent text formatting, missing timestamps, and date anomalies. By deliberately introducing ~2% duplicate rows, text casing variations (`' network '` vs `'SOFTWARE'`), missing resolution timestamps, and invalid dates, I was able to build a realistic data cleaning pipeline using Pandas to replicate real-world data governance tasks.

---

### Q4: How did your Python cleaning script (`02_clean_data.py`) handle duplicate rows and text inconsistencies?
**Answer:**  
I used Pandas functions:
1. `df.drop_duplicates()` to detect and remove exact duplicate rows.
2. `df['category'].astype(str).str.strip().str.title()` to strip extra leading/trailing whitespace and standardize text into Title Case (e.g., `'Network'`).
3. For categories that did not match valid standard options, I applied fallback business logic mapping to ensure consistency across the dataset.

---

### Q5: How did you handle invalid dates (e.g. `resolved_date` earlier than `created_date`) in Pandas?
**Answer:**  
I first converted string date columns to datetime objects using `pd.to_datetime(df['date_column'], errors='coerce')` so invalid string dates became `NaT` (Not a Time).  
Next, I identified logical date errors where `resolved_date < created_date`. For these records, I applied business repair logic by calculating a valid resolution timestamp (`created_date + pd.Timedelta(hours=sla_target_hours * 0.75)`), ensuring resolution time was positive and mathematically sound.

---

### Q6: What business logic did you use when a ticket was marked 'Resolved' or 'Closed' but had no resolution timestamp?
**Answer:**  
In service governance, a ticket cannot logically be considered "Resolved" if there is no recorded resolution timestamp. In my script, I identified tickets where `status.isin(['Resolved', 'Closed'])` and `resolved_date.isnull()`, and automatically reverted their status to `'Open'`. This ensured that SLA compliance calculations were not corrupted by incomplete ticket records.

---

### Q7: What is the difference between `GROUP BY` and Window Functions (`OVER()`) in SQL?
**Answer:**  
- **`GROUP BY`** collapses multiple rows into a single summary row per group (e.g., calculating total ticket count for each priority level).
- **Window Functions** (`OVER()`) perform calculations across a set of table rows related to the current row without collapsing the rows. For example, in Query 6, I used `LAG(ticket_volume, 1) OVER (ORDER BY month)` to retrieve the previous month's ticket volume alongside the current month's volume to calculate Month-over-Month growth percentage.

---

### Q8: Explain how you computed Month-over-Month (MoM) volume growth in SQL Query 6.
**Answer:**  
I used a Common Table Expression (CTE) combined with the `LAG()` window function:
```sql
WITH monthly_volume AS (
    SELECT strftime('%Y-%m', created_date) AS month, COUNT(*) AS ticket_volume
    FROM tickets
    GROUP BY month
)
SELECT 
    month,
    ticket_volume,
    LAG(ticket_volume, 1) OVER (ORDER BY month) AS prev_month_volume,
    ticket_volume - LAG(ticket_volume, 1) OVER (ORDER BY month) AS volume_change,
    ROUND(100.0 * (ticket_volume - LAG(ticket_volume, 1) OVER (ORDER BY month)) / 
          NULLIF(LAG(ticket_volume, 1) OVER (ORDER BY month), 0), 2) AS mom_growth_pct
FROM monthly_volume;
```
The CTE aggregated raw intake into monthly counts, and `LAG()` fetched the prior row's volume to compute the exact percentage growth.

---

### Q9: How did you use CTEs and `CASE WHEN` statements to identify breached tickets in SQL Query 7?
**Answer:**  
I created a CTE called `breach_analysis` that grouped data by `priority` and `category`. Inside the aggregate functions, I used `CASE WHEN sla_met = 0 THEN 1 ELSE 0 END` to sum only the breached tickets, and computed the average delay hours using `AVG(CASE WHEN sla_met = 0 THEN resolution_hours - sla_target_hours ELSE NULL END)`. This isolated breached records for governance review.

---

### Q10: Why did you write real Excel formulas (`COUNTIFS`, `AVERAGEIFS`) in Python rather than writing pre-computed values?
**Answer:**  
Writing real Excel formulas makes the report **dynamic and interactive**. If business users filter or update data in the `Cleaned_Data` sheet or add new row entries, all KPI cards, monthly tables, and dashboard charts update automatically in Excel without needing to re-run Python code. Hardcoded numbers make reports static and prone to staleness.

---

### Q11: What formulas did you use in the `KPI_Summary` sheet and how did conditional formatting help?
**Answer:**  
In the `KPI_Summary` sheet:
- **Total Tickets**: `=COUNTA(Cleaned_Data!A2:A5001)` -> 5,000
- **Overall SLA %**: `=COUNTIFS(Calc!I2:I5001, "YES") / (COUNTIFS(Cleaned_Data!H2:H5001, "Resolved") + COUNTIFS(Cleaned_Data!H2:H5001, "Closed"))` -> 81.77%
- **Avg Resolution Time**: `=AVERAGE(Calc!H2:H5001)` -> 20.35 hrs
- **Open Backlog**: `=COUNTIFS(Cleaned_Data!H2:H5001, "Open") + COUNTIFS(Cleaned_Data!H2:H5001, "Escalated")` -> 562

I applied conditional formatting to the Overall SLA Compliance % card:
- **Green Fill** if SLA >= 90% (On Target)
- **Amber Fill** if 80% <= SLA < 90% (Warning Zone)
- **Red Fill** if SLA < 80% (Breach Threshold)  
This visual hierarchy immediately directs leadership focus to operational risk areas.

---

### Q12: How did you cross-validate your SQL query results against your Excel report outputs?
**Answer:**  
I performed a 4-point QA cross-check between the outputs produced by `03_load_to_sql.py` (`sql_kpi_summary.csv`) and the Excel formula outputs:
1. **Total Ticket Volume**: SQL = 5,000 | Excel `=COUNTA()` = 5,000 *(100% Match)*
2. **Total Resolved Tickets**: SQL = 4,438 | Excel `=COUNTIFS("Resolved") + COUNTIFS("Closed")` = 4,438 *(100% Match)*
3. **Overall SLA Compliance %**: SQL = 81.77% | Excel `=COUNTIFS("YES") / Resolved_Tickets` = 81.77% *(100% Match)*
4. **Total Open Backlog**: SQL = 562 tickets | Excel `=COUNTIFS("Open") + COUNTIFS("Escalated")` = 562 tickets *(100% Match)*

This verified that both SQL analytical logic and Excel formula logic were 100% aligned.

---

### Q13: Why is setting a fixed random seed (`random.seed(42)`) important in Python automation?
**Answer:**  
Setting a fixed seed ensures **reproducibility**. Every time `01_generate_data.py` runs, it generates the exact same dataset with identical values. This allows test runs, debugging, and verification checks to yield predictable, consistent results.

---

### Q14: How does `run_all.py` orchestrate the pipeline and handle errors?
**Answer:**  
`run_all.py` uses Python's `subprocess` module to execute each step in sequence (`01_generate_data.py` → `02_clean_data.py` → `03_load_to_sql.py` → `04_build_reports.py` → `generate_weekly_summary`).  
It incorporates `logging` and inspects the `returncode` of each child process. If any script fails, `run_all.py` logs the error traceback and exits immediately, preventing corrupted data from cascading into downstream reports.

---

### Q15: As a Measurement & Reporting Associate, what steps would you take if monthly SLA compliance dropped below 80%?
**Answer:**  
1. **Root-Cause Data Drilling**: Drill down into the specific priority levels, categories, and teams responsible for the drop (e.g., checking if P1 tickets or Network category caused the drop).
2. **Backlog & Volume Spike Analysis**: Check MoM ticket volume growth (Query 6) to see if an unexpected intake spike overwhelmed staffing capacity.
3. **Agent & Team Capacity Review**: Analyze agent workload distribution to see if tickets were bottled up with specific agents or teams.
4. **Governance Briefing**: Publish a targeted exception report with clear charts and recommend reallocating L2/L1 support staff to bottlenecked categories.
