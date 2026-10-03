"""
03_load_to_sql.py
-----------------
Purpose: Loads cleaned_tickets.csv into SQLite DB (data/sla_reporting.db),
executes 7 analytical SQL queries from sql/queries.sql, prints formatted output,
and exports results to CSV files for cross-validation with Excel reports.
"""

import os
import sqlite3
import pandas as pd

def main():
    print("=" * 60)
    print("STEP 3: LOADING TO SQLITE & RUNNING ANALYTICAL SQL QUERIES")
    print("=" * 60)

    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    cleaned_csv_path = os.path.join(base_dir, 'data', 'cleaned_tickets.csv')
    db_path = os.path.join(base_dir, 'data', 'sla_reporting.db')
    sql_file_path = os.path.join(base_dir, 'sql', 'queries.sql')
    output_dir = os.path.join(base_dir, 'output', 'sql_results')

    os.makedirs(output_dir, exist_ok=True)

    if not os.path.exists(cleaned_csv_path):
        raise FileNotFoundError(f"Cleaned dataset not found at {cleaned_csv_path}. Run 02_clean_data.py first.")

    # 1. Load Cleaned Dataset into pandas
    df = pd.read_csv(cleaned_csv_path)

    # 2. Connect to SQLite and write to table 'tickets'
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Drop existing table if present and insert fresh data
    cursor.execute("DROP TABLE IF EXISTS tickets;")
    df.to_sql('tickets', conn, index=False, if_exists='replace')
    print(f"Successfully loaded {len(df)} records into SQLite database at: {db_path}")

    # 3. Parse and Execute SQL Queries from sql/queries.sql
    if not os.path.exists(sql_file_path):
        raise FileNotFoundError(f"SQL file not found at {sql_file_path}")

    with open(sql_file_path, 'r', encoding='utf-8') as f:
        sql_script = f.read()

    # Split script into individual queries by semicolon
    raw_queries = [q.strip() for q in sql_script.split(';') if q.strip()]

    query_titles = [
        "1_monthly_sla_compliance",
        "2_priority_sla_compliance",
        "3_category_avg_resolution",
        "4_team_backlog",
        "5_top_5_agents",
        "6_mom_volume_growth",
        "7_breached_tickets_breakdown"
    ]

    sql_results = {}

    for idx, (title, query_sql) in enumerate(zip(query_titles, raw_queries), 1):
        print(f"\n--- Running Query {idx}: {title} ---")
        query_df = pd.read_sql_query(query_sql, conn)
        print(query_df.to_string(index=False))

        # Save to CSV
        csv_out_path = os.path.join(output_dir, f"{title}.csv")
        query_df.to_csv(csv_out_path, index=False)
        sql_results[title] = query_df

    # Also generate a consolidated SQL summary CSV for QA verification
    overall_kpi_query = """
    SELECT 
        COUNT(*) AS total_tickets,
        SUM(CASE WHEN status IN ('Resolved', 'Closed') THEN 1 ELSE 0 END) AS total_resolved,
        SUM(CASE WHEN status IN ('Open', 'Escalated') THEN 1 ELSE 0 END) AS total_backlog,
        SUM(CASE WHEN sla_met = 1 THEN 1 ELSE 0 END) AS total_met_sla,
        SUM(CASE WHEN sla_met = 0 THEN 1 ELSE 0 END) AS total_breached,
        ROUND(100.0 * SUM(CASE WHEN sla_met = 1 THEN 1 ELSE 0 END) / 
              NULLIF(SUM(CASE WHEN status IN ('Resolved', 'Closed') THEN 1 ELSE 0 END), 0), 2) AS overall_sla_pct,
        ROUND(AVG(resolution_hours), 2) AS avg_resolution_hours,
        ROUND(AVG(customer_satisfaction), 2) AS avg_csat
    FROM tickets;
    """
    overall_df = pd.read_sql_query(overall_kpi_query, conn)
    summary_path = os.path.join(base_dir, 'output', 'sql_kpi_summary.csv')
    overall_df.to_csv(summary_path, index=False)
    
    print("\n" + "=" * 40)
    print("OVERALL SQL KPI SUMMARY")
    print("=" * 40)
    print(overall_df.to_string(index=False))
    print("=" * 40 + "\n")

    conn.close()
    print(f"Success! All SQL query results exported to: {output_dir}")
    print("-" * 60)

if __name__ == '__main__':
    main()
