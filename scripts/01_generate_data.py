"""
01_generate_data.py
-------------------
Purpose: Generates a realistic dataset of ~5,000 support tickets covering 12 months (2024)
for an IT/Customer Support Service Operations team.

Includes deliberate data quality flaws:
- ~2% duplicate rows
- Inconsistent text casing & extra leading/trailing whitespace
- Missing resolved dates for some resolved/closed tickets
- Invalid dates (e.g. resolution date prior to creation date or bad strings)
- Outlier resolution times

Seed fixed for 100% reproducibility.
"""

import os
import random
import numpy as np
import pandas as pd
from datetime import datetime, timedelta

def main():
    print("=" * 60)
    print("STEP 1: GENERATING SYNTHETIC SUPPORT TICKET DATASET")
    print("=" * 60)

    # Set random seeds for reproducibility
    random.seed(42)
    np.random.seed(42)

    # Base configuration
    num_tickets = 5000
    start_date = datetime(2024, 1, 1, 8, 0, 0)
    end_date = datetime(2024, 12, 31, 20, 0, 0)
    total_seconds_span = int((end_date - start_date).total_seconds())

    priorities = ['P1', 'P2', 'P3', 'P4']
    priority_weights = [0.10, 0.25, 0.45, 0.20]  # P3 is most common, P1 is rare urgent
    sla_target_map = {'P1': 4, 'P2': 8, 'P3': 24, 'P4': 48}  # Hours

    categories = ['Network', 'Software', 'Hardware', 'Access', 'Billing']
    category_weights = [0.20, 0.35, 0.15, 0.20, 0.10]

    teams = [
        'L1 Support', 
        'L2 Support', 
        'Infrastructure', 
        'Application Support', 
        'Billing Operations'
    ]

    agents = [
        'Alex Mercer', 'Sarah Connor', 'John Doe', 'Emily Watson', 'Michael Scott',
        'Pam Beesly', 'Jim Halpert', 'Dwight Schrute', 'Ravi Patel', 'Anita Sharma',
        'Carlos Mendez', 'Priya Nair'
    ]

    statuses = ['Resolved', 'Closed', 'Open', 'Escalated']
    status_weights = [0.55, 0.35, 0.06, 0.04]  # ~90% completed, 10% backlog

    # Generate ticket timestamps and records
    records = []
    
    for i in range(1, num_tickets + 1):
        ticket_id = f"TCK-{10000 + i}"
        
        # Pick random creation timestamp throughout 2024
        random_offset_sec = random.randint(0, total_seconds_span)
        created_dt = start_date + timedelta(seconds=random_offset_sec)

        priority = random.choices(priorities, weights=priority_weights)[0]
        category = random.choices(categories, weights=category_weights)[0]
        team = random.choice(teams)
        agent = random.choice(agents)
        status = random.choices(statuses, weights=status_weights)[0]
        sla_target = sla_target_map[priority]

        # Resolution Date calculation
        resolved_dt = None
        csat = None

        if status in ['Resolved', 'Closed']:
            # 82% of tickets meet SLA target, 18% breach target
            meets_sla = random.random() < 0.82

            if meets_sla:
                # Resolution time between 10% and 95% of SLA target
                actual_hours = random.uniform(0.1 * sla_target, 0.95 * sla_target)
            else:
                # Breached tickets take 105% to 300% of SLA target
                actual_hours = random.uniform(1.05 * sla_target, 3.0 * sla_target)
            
            resolved_dt = created_dt + timedelta(hours=actual_hours)

            # Assign CSAT score (1-5), slightly higher if SLA met
            if meets_sla:
                csat = random.choices([3, 4, 5], weights=[0.1, 0.4, 0.5])[0]
            else:
                csat = random.choices([1, 2, 3, 4], weights=[0.3, 0.4, 0.2, 0.1])[0]

        record = {
            'ticket_id': ticket_id,
            'created_date': created_dt.strftime('%Y-%m-%d %H:%M:%S'),
            'resolved_date': resolved_dt.strftime('%Y-%m-%d %H:%M:%S') if resolved_dt else None,
            'priority': priority,
            'category': category,
            'team': team,
            'agent': agent,
            'status': status,
            'customer_satisfaction': csat,
            'sla_target_hours': sla_target
        }
        records.append(record)

    df = pd.DataFrame(records)

    # -------------------------------------------------------------
    # INTRODUCE DELIBERATE DATA QUALITY PROBLEMS (FOR STEP 2 DEMO)
    # -------------------------------------------------------------
    print("Introducing deliberate data quality issues for data cleaning practice...")

    # 1. Inconsistent Category Casing & Whitespace (~8% of rows)
    casing_variations = [' network ', 'SOFTWARE', 'Hardware ', ' access ', 'BILLING', '  Software  ']
    dirty_indices = df.sample(frac=0.08, random_state=42).index
    for idx in dirty_indices:
        df.loc[idx, 'category'] = random.choice(casing_variations)

    # 2. Missing resolved_date for resolved/closed tickets (~1.5% of rows)
    missing_res_idx = df[df['status'].isin(['Resolved', 'Closed'])].sample(frac=0.015, random_state=42).index
    df.loc[missing_res_idx, 'resolved_date'] = None

    # 3. Invalid dates / Negative resolution times (resolved before creation) (~1% of rows)
    invalid_date_idx = df.sample(n=25, random_state=42).index
    for idx in invalid_date_idx:
        created_val = pd.to_datetime(df.loc[idx, 'created_date'])
        # Set resolved date 5 hours BEFORE created date
        df.loc[idx, 'resolved_date'] = (created_val - timedelta(hours=5)).strftime('%Y-%m-%d %H:%M:%S')

    # 4. Severe Resolution Outliers (e.g., 600-1000 hours resolution) (~0.5% of rows)
    outlier_idx = df[df['status'].isin(['Resolved', 'Closed'])].sample(n=15, random_state=42).index
    for idx in outlier_idx:
        created_val = pd.to_datetime(df.loc[idx, 'created_date'])
        df.loc[idx, 'resolved_date'] = (created_val + timedelta(hours=random.randint(500, 900))).strftime('%Y-%m-%d %H:%M:%S')

    # 5. Exact Duplicate Rows (~2% of rows)
    duplicates = df.sample(frac=0.02, random_state=42)
    df = pd.concat([df, duplicates], ignore_index=True)

    # Save to data/raw_tickets.csv
    output_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'data'))
    os.makedirs(output_dir, exist_ok=True)
    raw_path = os.path.join(output_dir, 'raw_tickets.csv')

    df.to_csv(raw_path, index=False)
    print(f"Success! Generated {len(df)} raw ticket records saved to: {raw_path}")
    print("-" * 60)

if __name__ == '__main__':
    main()
