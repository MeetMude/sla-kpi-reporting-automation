-- ============================================================================
-- SLA & KPI REPORTING AUTOMATION - SQL ANALYTICAL QUERIES
-- Database: SQLite (data/sla_reporting.db)
-- Table Name: tickets
-- ============================================================================

-- ----------------------------------------------------------------------------
-- QUERY 1: Monthly SLA Compliance Percentage
-- Purpose: Analyzes how SLA performance trends month-by-month over the 12-month period.
-- Technique: GROUP BY month using strftime, conditional aggregation with CASE WHEN.
-- ----------------------------------------------------------------------------
SELECT 
    strftime('%Y-%m', created_date) AS month,
    COUNT(*) AS total_tickets,
    SUM(CASE WHEN status IN ('Resolved', 'Closed') THEN 1 ELSE 0 END) AS resolved_tickets,
    SUM(CASE WHEN sla_met = 1 THEN 1 ELSE 0 END) AS tickets_met_sla,
    ROUND(100.0 * SUM(CASE WHEN sla_met = 1 THEN 1 ELSE 0 END) / 
          NULLIF(SUM(CASE WHEN status IN ('Resolved', 'Closed') THEN 1 ELSE 0 END), 0), 2) AS sla_compliance_pct
FROM tickets
GROUP BY month
ORDER BY month ASC;


-- ----------------------------------------------------------------------------
-- QUERY 2: SLA Compliance Percentage by Priority
-- Purpose: Evaluates SLA breach risk across priority levels (P1 Urgent vs P4 Low).
-- Technique: GROUP BY priority, conditional ratio calculation.
-- ----------------------------------------------------------------------------
SELECT 
    priority,
    COUNT(*) AS total_tickets,
    SUM(CASE WHEN status IN ('Resolved', 'Closed') THEN 1 ELSE 0 END) AS resolved_tickets,
    SUM(CASE WHEN sla_met = 1 THEN 1 ELSE 0 END) AS tickets_met_sla,
    ROUND(100.0 * SUM(CASE WHEN sla_met = 1 THEN 1 ELSE 0 END) / 
          NULLIF(SUM(CASE WHEN status IN ('Resolved', 'Closed') THEN 1 ELSE 0 END), 0), 2) AS sla_compliance_pct,
    ROUND(AVG(resolution_hours), 2) AS avg_resolution_hours
FROM tickets
GROUP BY priority
ORDER BY priority ASC;


-- ----------------------------------------------------------------------------
-- QUERY 3: Average Resolution Time and CSAT by Category
-- Purpose: Identifies high-friction operational categories (e.g. Hardware vs Software).
-- Technique: GROUP BY category, aggregate functions AVG() and COUNT().
-- ----------------------------------------------------------------------------
SELECT 
    category,
    COUNT(*) AS total_resolved_tickets,
    ROUND(AVG(resolution_hours), 2) AS avg_resolution_hours,
    ROUND(AVG(sla_target_hours), 2) AS avg_sla_target_hours,
    ROUND(AVG(customer_satisfaction), 2) AS avg_csat_score
FROM tickets
WHERE status IN ('Resolved', 'Closed')
GROUP BY category
ORDER BY avg_resolution_hours DESC;


-- ----------------------------------------------------------------------------
-- QUERY 4: Open Ticket Backlog Breakdown by Support Team
-- Purpose: Measures current workload distribution and bottlenecks across support teams.
-- Technique: GROUP BY team, conditional counting for Open vs Escalated status.
-- ----------------------------------------------------------------------------
SELECT 
    team,
    COUNT(*) AS total_backlog,
    SUM(CASE WHEN status = 'Open' THEN 1 ELSE 0 END) AS open_tickets,
    SUM(CASE WHEN status = 'Escalated' THEN 1 ELSE 0 END) AS escalated_tickets
FROM tickets
WHERE status IN ('Open', 'Escalated')
GROUP BY team
ORDER BY total_backlog DESC;


-- ----------------------------------------------------------------------------
-- QUERY 5: Top 5 High-Performing Agents by SLA Adherence
-- Purpose: Recognizes top tier support agents based on resolved volume and SLA compliance.
-- Technique: GROUP BY agent and team, ORDER BY count and percentage, LIMIT 5.
-- ----------------------------------------------------------------------------
SELECT 
    agent,
    team,
    COUNT(*) AS total_resolved,
    SUM(CASE WHEN sla_met = 1 THEN 1 ELSE 0 END) AS tickets_met_sla,
    ROUND(100.0 * SUM(CASE WHEN sla_met = 1 THEN 1 ELSE 0 END) / COUNT(*), 2) AS sla_compliance_pct,
    ROUND(AVG(customer_satisfaction), 2) AS avg_csat
FROM tickets
WHERE status IN ('Resolved', 'Closed')
GROUP BY agent, team
ORDER BY tickets_met_sla DESC, sla_compliance_pct DESC
LIMIT 5;


-- ----------------------------------------------------------------------------
-- QUERY 6: Month-over-Month (MoM) Ticket Volume Change %
-- Purpose: Tracks ticket intake growth rates across months to assist workforce planning.
-- Technique: Common Table Expression (CTE) + Window Function (LAG).
-- ----------------------------------------------------------------------------
WITH monthly_volume AS (
    SELECT 
        strftime('%Y-%m', created_date) AS month,
        COUNT(*) AS ticket_volume
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
FROM monthly_volume
ORDER BY month ASC;


-- ----------------------------------------------------------------------------
-- QUERY 7: Breached Tickets Detailed Breakdown by Priority and Category
-- Purpose: Pinpoints exact areas of SLA failure for governance review.
-- Technique: CTE + CASE WHEN filtering for breached tickets (sla_met = 0).
-- ----------------------------------------------------------------------------
WITH breach_analysis AS (
    SELECT 
        priority,
        category,
        COUNT(*) AS total_tickets,
        SUM(CASE WHEN sla_met = 0 THEN 1 ELSE 0 END) AS breached_tickets,
        ROUND(AVG(CASE WHEN sla_met = 0 THEN resolution_hours - sla_target_hours ELSE NULL END), 2) AS avg_delay_hours
    FROM tickets
    WHERE status IN ('Resolved', 'Closed')
    GROUP BY priority, category
)
SELECT 
    priority,
    category,
    total_tickets,
    breached_tickets,
    ROUND(100.0 * breached_tickets / total_tickets, 2) AS breach_rate_pct,
    COALESCE(avg_delay_hours, 0) AS avg_delay_hours
FROM breach_analysis
WHERE breached_tickets > 0
ORDER BY breached_tickets DESC, breach_rate_pct DESC;
