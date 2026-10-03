"""
run_all.py
----------
Master Orchestrator Script for SLA & KPI Reporting Automation.

Executes all pipeline steps in sequence:
1. 01_generate_data.py : Generates synthetic raw ticket data
2. 02_clean_data.py     : Cleans, audits & standardizes data
3. 03_load_to_sql.py    : Loads clean data to SQLite & executes 7 SQL queries
4. 04_build_reports.py  : Generates multi-tab Excel SLA/KPI Report with formulas & charts
5. Weekly Operational Export : Exports weekly summary CSV

Includes step timing, error handling, and pipeline logs.
"""

import os
import sys
import time
import logging
import subprocess
import pandas as pd

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

def generate_weekly_summary(base_dir):
    logging.info("STARTING Step 5: Generating Weekly Operational Summary CSV...")
    cleaned_csv = os.path.join(base_dir, 'data', 'cleaned_tickets.csv')
    weekly_csv = os.path.join(base_dir, 'output', 'weekly_summary.csv')

    if not os.path.exists(cleaned_csv):
        raise FileNotFoundError(f"Cleaned dataset missing: {cleaned_csv}")

    df = pd.read_csv(cleaned_csv)
    df['created_date'] = pd.to_datetime(df['created_date'])
    df['week_number'] = df['created_date'].dt.isocalendar().week

    weekly_df = df.groupby('week_number').agg(
        total_tickets=('ticket_id', 'count'),
        resolved_tickets=('status', lambda x: x.isin(['Resolved', 'Closed']).sum()),
        backlog_tickets=('status', lambda x: x.isin(['Open', 'Escalated']).sum()),
        sla_met_count=('sla_met', lambda x: (x == 1).sum()),
        avg_resolution_hrs=('resolution_hours', 'mean'),
        avg_csat=('customer_satisfaction', 'mean')
    ).reset_index()

    weekly_df['sla_compliance_pct'] = (100.0 * weekly_df['sla_met_count'] / weekly_df['resolved_tickets'].replace(0, 1)).round(2)
    weekly_df['avg_resolution_hrs'] = weekly_df['avg_resolution_hrs'].round(2)
    weekly_df['avg_csat'] = weekly_df['avg_csat'].round(2)

    weekly_df.to_csv(weekly_csv, index=False)
    logging.info(f"Weekly operational summary exported to: {weekly_csv}\n")

def main():
    t0 = time.time()
    print("\n" + "=" * 70)
    print("      SLA & KPI REPORTING AUTOMATION - FULL PIPELINE EXECUTION      ")
    print("=" * 70 + "\n")

    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    scripts_dir = os.path.join(base_dir, 'scripts')

    steps = [
        ("Step 1: Data Generation", os.path.join(scripts_dir, "01_generate_data.py")),
        ("Step 2: Data Cleaning", os.path.join(scripts_dir, "02_clean_data.py")),
        ("Step 3: SQL Database & Analytical Queries", os.path.join(scripts_dir, "03_load_to_sql.py")),
        ("Step 4: Excel Report Builder", os.path.join(scripts_dir, "04_build_reports.py")),
    ]

    for step_name, script_path in steps:
        logging.info(f"Executing {step_name}...")
        step_start = time.time()
        
        result = subprocess.run([sys.executable, script_path], capture_output=True, text=True)
        
        if result.returncode != 0:
            logging.error(f"Error in {step_name}:\n{result.stderr}")
            sys.exit(1)
        else:
            print(result.stdout)
            elapsed = time.time() - step_start
            logging.info(f"Finished {step_name} in {elapsed:.2f} seconds.\n")

    # Step 5: Operational Summary
    generate_weekly_summary(base_dir)

    total_time = time.time() - t0
    print("=" * 70)
    print(f"PIPELINE COMPLETED SUCCESSFULLY IN {total_time:.2f} SECONDS!")
    print("Generated Outputs:")
    print(f"  [1] Data Files   : data/raw_tickets.csv, data/cleaned_tickets.csv, data/sla_reporting.db")
    print(f"  [2] SQL Outputs  : output/sql_kpi_summary.csv, output/sql_results/*.csv")
    print(f"  [3] Excel Report : output/SLA_KPI_Report.xlsx (7 worksheets)")
    print(f"  [4] Summary CSV  : output/weekly_summary.csv")
    print("=" * 70 + "\n")

if __name__ == '__main__':
    main()
