"""
STEP 01 — Data Cleaning & Orders Fact Table Builder
====================================================
Purpose:
  - Load all raw CSV files
  - Clean and validate data
  - Join tables to create orders_fact (the master analytics table)
  - Export cleaned data to /data/processed/

Run: python3 python_analysis/01_data_cleaning.py
"""

import pandas as pd
import numpy as np
import os

# ─────────────────────────────────────────────
# PATHS
# ─────────────────────────────────────────────
BASE        = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW_DIR     = os.path.join(BASE, "data", "raw")
PROC_DIR    = os.path.join(BASE, "data", "processed")
CSV_OUT     = os.path.join(BASE, "outputs", "csv_exports")

os.makedirs(PROC_DIR, exist_ok=True)
os.makedirs(CSV_OUT, exist_ok=True)

print("=" * 60)
print("STEP 01 — Data Cleaning & Fact Table Builder")
print("=" * 60)

# ─────────────────────────────────────────────
# 1. LOAD RAW DATA
# ─────────────────────────────────────────────
print("\n[1] Loading raw data files...")

customers   = pd.read_csv(f"{RAW_DIR}/customers.csv")
orders      = pd.read_csv(f"{RAW_DIR}/orders.csv",      parse_dates=["order_purchase_timestamp",
                                                                       "order_delivered_customer_date"])
order_items = pd.read_csv(f"{RAW_DIR}/order_items.csv")
products    = pd.read_csv(f"{RAW_DIR}/products.csv")
payments    = pd.read_csv(f"{RAW_DIR}/payments.csv")
reviews     = pd.read_csv(f"{RAW_DIR}/reviews.csv")

tables = {
    "customers":   customers,
    "orders":      orders,
    "order_items": order_items,
    "products":    products,
    "payments":    payments,
    "reviews":     reviews,
}

for name, df in tables.items():
    print(f"  {name:15s}: {len(df):>8,} rows  |  {df.shape[1]} columns  "
          f"|  nulls: {df.isnull().sum().sum()}")

# ─────────────────────────────────────────────
# 2. CLEAN EACH TABLE
# ─────────────────────────────────────────────
print("\n[2] Cleaning tables...")

# ── ORDERS ──────────────────────────────────
before = len(orders)
orders.drop_duplicates(subset="order_id", inplace=True)
orders = orders[orders["order_status"] != "cancelled"]   # drop cancelled
after  = len(orders)
print(f"  orders:      {before:,} → {after:,} (removed {before-after:,} duplicates/cancelled)")

# Parse timestamps properly
orders["order_purchase_timestamp"]    = pd.to_datetime(orders["order_purchase_timestamp"])
orders["order_delivered_customer_date"] = pd.to_datetime(orders["order_delivered_customer_date"])

# Date features
orders["order_year"]       = orders["order_purchase_timestamp"].dt.year
orders["order_month"]      = orders["order_purchase_timestamp"].dt.month
orders["order_year_month"] = orders["order_purchase_timestamp"].dt.to_period("M").astype(str)
orders["order_quarter"]    = orders["order_purchase_timestamp"].dt.quarter
orders["order_day_name"]   = orders["order_purchase_timestamp"].dt.day_name()

# Delivery time (days)
orders["delivery_days"] = (
    orders["order_delivered_customer_date"] - orders["order_purchase_timestamp"]
).dt.days.clip(lower=0)

# ── ORDER ITEMS ──────────────────────────────
before = len(order_items)
order_items.drop_duplicates(inplace=True)
# Remove negative prices
order_items = order_items[(order_items["price"] > 0) & (order_items["freight_value"] >= 0)]
order_items["profit_estimate"] = (order_items["price"] - order_items["freight_value"]).round(2)
order_items["margin_pct"]      = (order_items["profit_estimate"] / order_items["price"] * 100).round(2)
after = len(order_items)
print(f"  order_items: {before:,} → {after:,}")

# ── PRODUCTS ─────────────────────────────────
before = len(products)
products.drop_duplicates(subset="product_id", inplace=True)
products["product_category_name"].fillna("unknown", inplace=True)
after = len(products)
print(f"  products:    {before:,} → {after:,}")

# ── CUSTOMERS ────────────────────────────────
customers.drop_duplicates(subset="customer_id", inplace=True)
print(f"  customers:   {len(customers):,}")

# ── PAYMENTS ─────────────────────────────────
payments.drop_duplicates(subset="order_id", inplace=True)
print(f"  payments:    {len(payments):,}")

# ── REVIEWS ──────────────────────────────────
reviews.drop_duplicates(subset="order_id", inplace=True)
reviews["review_score"] = reviews["review_score"].clip(1, 5)
print(f"  reviews:     {len(reviews):,}")

# ─────────────────────────────────────────────
# 3. BUILD ORDERS FACT TABLE
# ─────────────────────────────────────────────
print("\n[3] Building orders_fact table (joining all tables)...")

fact = (
    orders
    .merge(order_items,  on="order_id",   how="inner")
    .merge(products,     on="product_id", how="left")
    .merge(customers,    on="customer_id",how="left")
    .merge(payments,     on="order_id",   how="left")
    .merge(reviews[["order_id","review_score"]], on="order_id", how="left")
)

# Fill missing review scores with 0
fact["review_score"].fillna(0, inplace=True)
fact["payment_type"].fillna("unknown", inplace=True)

print(f"  orders_fact: {len(fact):,} rows | {fact.shape[1]} columns")

# ─────────────────────────────────────────────
# 4. DATA QUALITY REPORT
# ─────────────────────────────────────────────
print("\n[4] Data Quality Report:")
print(f"  Date range   : {fact['order_purchase_timestamp'].min().date()} → "
      f"{fact['order_purchase_timestamp'].max().date()}")
print(f"  Unique orders    : {fact['order_id'].nunique():,}")
print(f"  Unique customers : {fact['customer_id'].nunique():,}")
print(f"  Unique products  : {fact['product_id'].nunique():,}")
print(f"  Categories       : {fact['product_category_name'].nunique()}")
print(f"  States covered   : {fact['customer_state'].nunique()}")
print(f"  Total Revenue    : R$ {fact['price'].sum():,.2f}")
print(f"  Total Freight    : R$ {fact['freight_value'].sum():,.2f}")
print(f"  Total Profit Est : R$ {fact['profit_estimate'].sum():,.2f}")
print(f"  Avg Order Value  : R$ {fact.groupby('order_id')['price'].sum().mean():,.2f}")
print(f"  Avg Review Score : {fact[fact['review_score'] > 0]['review_score'].mean():.2f}")
print(f"  Null count       : {fact.isnull().sum().sum()}")

# ─────────────────────────────────────────────
# 5. EXPORT
# ─────────────────────────────────────────────
print("\n[5] Exporting cleaned files...")

fact.to_csv(f"{PROC_DIR}/orders_fact.csv", index=False)
print(f"  Saved: data/processed/orders_fact.csv")

# Also save a quick summary
summary = pd.DataFrame({
    "Metric": [
        "Total Orders", "Unique Customers", "Unique Products", "Categories",
        "Total Revenue (R$)", "Total Freight (R$)", "Total Profit Est (R$)",
        "Avg Order Value (R$)", "Avg Review Score"
    ],
    "Value": [
        f"{fact['order_id'].nunique():,}",
        f"{fact['customer_id'].nunique():,}",
        f"{fact['product_id'].nunique():,}",
        f"{fact['product_category_name'].nunique()}",
        f"{fact['price'].sum():,.2f}",
        f"{fact['freight_value'].sum():,.2f}",
        f"{fact['profit_estimate'].sum():,.2f}",
        f"{fact.groupby('order_id')['price'].sum().mean():,.2f}",
        f"{fact[fact['review_score']>0]['review_score'].mean():.2f}",
    ]
})
summary.to_csv(f"{CSV_OUT}/data_quality_summary.csv", index=False)
print(f"  Saved: outputs/csv_exports/data_quality_summary.csv")

print("\n✅ STEP 01 COMPLETE\n")
