"""
generate_data.py
Generates realistic synthetic CRM and sales transaction data.
"""

import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

random.seed(42)
np.random.seed(42)

# --- Configuration ---
N_CUSTOMERS = 120
N_TRANSACTIONS = 800
START_DATE = datetime(2024, 1, 1)
END_DATE = datetime(2024, 12, 31)

SEGMENTS = ['SME', 'Enterprise', 'Startup']
SEGMENT_WEIGHTS = [0.5, 0.25, 0.25]

REGIONS = ['Dublin', 'Cork', 'Galway', 'Limerick', 'Other']
REGION_WEIGHTS = [0.45, 0.20, 0.15, 0.10, 0.10]

PRODUCTS = [
    'Analytics Pro', 'Reporting Suite', 'Data Connector',
    'Dashboard Builder', 'CRM Integration', 'Support Package'
]

REPS = ['Sarah M.', 'James O.', 'Aoife B.', 'Conor D.', 'Niamh R.']

STAGES = ['Won', 'Lost', 'Pending']
STAGE_WEIGHTS = [0.60, 0.25, 0.15]

INDUSTRY = ['Technology', 'Retail', 'Finance', 'Healthcare', 'Hospitality', 'Education']


def random_date(start, end):
    return start + timedelta(days=random.randint(0, (end - start).days))


def generate_customers():
    customers = []
    for i in range(1, N_CUSTOMERS + 1):
        segment = random.choices(SEGMENTS, weights=SEGMENT_WEIGHTS)[0]
        region = random.choices(REGIONS, weights=REGION_WEIGHTS)[0]
        acq_date = random_date(START_DATE, END_DATE - timedelta(days=30))
        customers.append({
            'customer_id': f'C{i:04d}',
            'company_name': f'Company {i:03d} Ltd',
            'segment': segment,
            'region': region,
            'industry': random.choice(INDUSTRY),
            'acquisition_date': acq_date.strftime('%Y-%m-%d'),
        })
    return pd.DataFrame(customers)


def generate_transactions(customers_df):
    transactions = []
    customer_ids = customers_df['customer_id'].tolist()
    customer_segments = customers_df.set_index('customer_id')['segment'].to_dict()

    for i in range(1, N_TRANSACTIONS + 1):
        customer_id = random.choice(customer_ids)
        segment = customer_segments[customer_id]

        # Deal size varies by segment
        if segment == 'Enterprise':
            amount = round(random.uniform(8000, 40000), 2)
        elif segment == 'SME':
            amount = round(random.uniform(1500, 12000), 2)
        else:  # Startup
            amount = round(random.uniform(500, 5000), 2)

        stage = random.choices(STAGES, weights=STAGE_WEIGHTS)[0]
        product = random.choice(PRODUCTS)
        rep = random.choice(REPS)
        t_date = random_date(START_DATE, END_DATE)

        # Add some growth trend — later months slightly more revenue
        month_boost = 1 + (t_date.month - 1) * 0.012
        amount = round(amount * month_boost, 2)

        transactions.append({
            'transaction_id': f'T{i:05d}',
            'customer_id': customer_id,
            'product': product,
            'amount': amount,
            'date': t_date.strftime('%Y-%m-%d'),
            'month': t_date.strftime('%Y-%m'),
            'month_name': t_date.strftime('%b %Y'),
            'stage': stage,
            'sales_rep': rep,
        })

    return pd.DataFrame(transactions)


def generate_and_save():
    customers = generate_customers()
    transactions = generate_transactions(customers)

    # Merge segment/region onto transactions
    transactions = transactions.merge(
        customers[['customer_id', 'segment', 'region', 'industry']],
        on='customer_id', how='left'
    )

    customers.to_csv('data/customers.csv', index=False)
    transactions.to_csv('data/transactions.csv', index=False)
    print(f"Generated {len(customers)} customers and {len(transactions)} transactions.")
    return customers, transactions


if __name__ == '__main__':
    import os
    os.makedirs('data', exist_ok=True)
    generate_and_save()
