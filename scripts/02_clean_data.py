"""
02_clean_data.py
----------------
Purpose: Audits, cleans, and standardizes raw ticket dataset using pandas.

Data Quality Tasks:
1. Detect & drop exact duplicate rows
2. Clean text formatting (strip whitespace, title case categories)
3. Parse datetime strings safely
4. Detect and repair invalid dates (e.g., resolved_date before created_date)
5. Re-classify tickets marked 'Resolved'/'Closed' with missing resolution timestamps to 'Open'
6. Compute resolution_hours and sla_met flag
7. Output comprehensive Data Quality Audit Log
"""

import os
import pandas as pd
import numpy as np

def main():
    print("=" * 60)
    print("STEP 2: CLEANING & AUDITING TICKET DATASET")
    print("=" * 60)

    # File paths
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    raw_path = os.path.join(base_dir, 'data', 'raw_tickets.csv')
    cleaned_path = os.path.join(base_dir, 'data', 'cleaned_tickets.csv')

    if not os.path.exists(raw_path):
        raise FileNotFoundError(f"Raw data file not found at {raw_path}. Run 01_generate_data.py first.")

    df_raw = pd.read_csv(raw_path)
    initial_rows = len(df_raw)
    print(f"Loaded raw dataset: {initial_rows} rows")

    # 1. Remove Duplicate Rows
    df_clean = df_raw.drop_duplicates()
    duplicates_removed = initial_rows - len(df_clean)

    # 2. Text Casing and Whitespace Cleaning
    df_clean['category'] = df_clean['category'].astype(str).str.strip().str.title()
    df_clean['team'] = df_clean['team'].astype(str).str.strip()
    df_clean['priority'] = df_clean['priority'].astype(str).str.strip().str.upper()
    df_clean['status'] = df_clean['status'].astype(str).str.strip().str.title()
    df_clean['agent'] = df_clean['agent'].astype(str).str.strip()

    # Standardize category typos/variations if any
    valid_categories = {'Network', 'Software', 'Hardware', 'Access', 'Billing'}
    df_clean['category'] = df_clean['category'].apply(lambda x: x if x in valid_categories else 'Software')

    # 3. Datetime Parsing
    df_clean['created_date'] = pd.to_datetime(df_clean['created_date'], errors='coerce')
    df_clean['resolved_date'] = pd.to_datetime(df_clean['resolved_date'], errors='coerce')

    # 4. Handle Invalid Resolution Timestamps (Resolved date < Created date)
    invalid_dates_mask = (df_clean['resolved_date'].notnull()) & (df_clean['resolved_date'] < df_clean['created_date'])
    invalid_dates_count = invalid_dates_mask.sum()
    
    # Fix invalid dates by setting resolved_date = created_date + (sla_target_hours * 0.75)
    for idx in df_clean[invalid_dates_mask].index:
        sla_hrs = df_clean.loc[idx, 'sla_target_hours']
        df_clean.loc[idx, 'resolved_date'] = df_clean.loc[idx, 'created_date'] + pd.Timedelta(hours=sla_hrs * 0.75)

    # 5. Handle missing resolved_date for Resolved/Closed tickets
    missing_res_mask = df_clean['status'].isin(['Resolved', 'Closed']) & df_clean['resolved_date'].isnull()
    missing_res_count = missing_res_mask.sum()
    # Business logic rule: if a ticket is marked Resolved/Closed but has no timestamp, revert status to 'Open'
    df_clean.loc[missing_res_mask, 'status'] = 'Open'

    # Ensure Open/Escalated tickets have null resolved_date and null CSAT
    open_mask = df_clean['status'].isin(['Open', 'Escalated'])
    df_clean.loc[open_mask, 'resolved_date'] = None
    df_clean.loc[open_mask, 'customer_satisfaction'] = None

    # 6. Compute Metrics: resolution_hours and sla_met
    df_clean['resolution_hours'] = np.nan
    df_clean['sla_met'] = np.nan

    resolved_mask = df_clean['resolved_date'].notnull()
    
    # Resolution hours = (resolved_date - created_date) in hours
    res_seconds = (df_clean.loc[resolved_mask, 'resolved_date'] - df_clean.loc[resolved_mask, 'created_date']).dt.total_seconds()
    df_clean.loc[resolved_mask, 'resolution_hours'] = (res_seconds / 3600.0).round(2)

    # SLA met flag: 1 if resolution_hours <= sla_target_hours, else 0
    sla_met_condition = df_clean.loc[resolved_mask, 'resolution_hours'] <= df_clean.loc[resolved_mask, 'sla_target_hours']
    df_clean.loc[resolved_mask, 'sla_met'] = sla_met_condition.astype(int)

    final_rows = len(df_clean)

    # -------------------------------------------------------------
    # DATA QUALITY AUDIT LOG PRINTING
    # -------------------------------------------------------------
    print("\n" + "-" * 40)
    print("DATA QUALITY AUDIT REPORT")
    print("-" * 40)
    print(f"Initial Rows Loaded          : {initial_rows}")
    print(f"Duplicate Rows Removed      : {duplicates_removed}")
    print(f"Invalid Resolution Dates Fix: {invalid_dates_count}")
    print(f"Missing Timestamp Status Fix: {missing_res_count}")
    print(f"Final Cleaned Rows           : {final_rows}")
    print(f"Resolved Tickets Count       : {resolved_mask.sum()}")
    print(f"Backlog (Open/Escalated)     : {open_mask.sum()}")
    print("-" * 40 + "\n")

    # Format datetimes back to standard strings for clean CSV output
    df_export = df_clean.copy()
    df_export['created_date'] = df_export['created_date'].dt.strftime('%Y-%m-%d %H:%M:%S')
    df_export['resolved_date'] = df_export['resolved_date'].dt.strftime('%Y-%m-%d %H:%M:%S')

    df_export.to_csv(cleaned_path, index=False)
    print(f"Success! Cleaned dataset saved to: {cleaned_path}")
    print("-" * 60)

if __name__ == '__main__':
    main()
