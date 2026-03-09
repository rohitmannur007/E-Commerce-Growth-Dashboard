"""
Data Generator for E-commerce Growth & Profit Optimization Analytics
Generates synthetic Olist-style Brazilian e-commerce data (100k orders)
"""

import pandas as pd
import numpy as np
import os

np.random.seed(42)

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
N_CUSTOMERS  = 10000
N_ORDERS     = 100000
N_PRODUCTS   = 3000
RAW_DIR      = os.path.join(os.path.dirname(__file__), "raw")

CATEGORIES = [
    "electronics",       "clothing_accessories", "home_decor",
    "sports_leisure",    "beauty_perfumery",      "furniture",
    "books",             "toys_games",            "kitchen_utilities",
    "health_wellness",   "automotive",            "garden_tools",
    "baby_products",     "stationery",            "pet_supplies",
]

STATES = ["SP","RJ","MG","RS","PR","BA","SC","GO","PE","CE",
          "AM","PA","MT","MS","ES","PB","MA","RN","AL","SE"]

# ─────────────────────────────────────────────
# 1. CUSTOMERS
# ─────────────────────────────────────────────
print("Generating customers...")
customers = pd.DataFrame({
    "customer_id":       [f"CUST_{i:06d}" for i in range(N_CUSTOMERS)],
    "customer_city":     np.random.choice(
                             ["Sao Paulo","Rio de Janeiro","Belo Horizonte",
                              "Curitiba","Porto Alegre","Salvador","Fortaleza",
                              "Brasilia","Manaus","Recife"], N_CUSTOMERS),
    "customer_state":    np.random.choice(STATES, N_CUSTOMERS),
    "customer_zip_code": np.random.randint(10000, 99999, N_CUSTOMERS),
})
customers.to_csv(f"{RAW_DIR}/customers.csv", index=False)
print(f"  customers.csv  → {len(customers):,} rows")

# ─────────────────────────────────────────────
# 2. PRODUCTS
# ─────────────────────────────────────────────
print("Generating products...")
products = pd.DataFrame({
    "product_id":            [f"PROD_{i:05d}" for i in range(N_PRODUCTS)],
    "product_category_name": np.random.choice(CATEGORIES, N_PRODUCTS,
                                 p=[0.18,0.14,0.10,0.09,0.08,0.07,0.06,
                                    0.05,0.05,0.04,0.04,0.03,0.03,0.02,0.02]),
    "product_weight_g":      np.random.randint(100, 5000, N_PRODUCTS),
    "product_length_cm":     np.random.randint(10, 80, N_PRODUCTS),
    "product_width_cm":      np.random.randint(5, 60, N_PRODUCTS),
    "product_height_cm":     np.random.randint(2, 50, N_PRODUCTS),
})
products.to_csv(f"{RAW_DIR}/products.csv", index=False)
print(f"  products.csv   → {len(products):,} rows")

# ─────────────────────────────────────────────
# 3. ORDERS  (with realistic time distribution)
# ─────────────────────────────────────────────
print("Generating orders...")

# Simulate 2 years: Jan 2022 – Dec 2023
# More orders toward end (growth trend)
dates = pd.date_range("2022-01-01", "2023-12-31", freq="h")
weights = np.linspace(0.5, 1.5, len(dates))   # growing trend
weights /= weights.sum()
order_timestamps = np.random.choice(dates, size=N_ORDERS, replace=True, p=weights)

STATUS_CHOICES = ["delivered","delivered","delivered","delivered",
                  "shipped","processing","cancelled"]

orders = pd.DataFrame({
    "order_id":                   [f"ORD_{i:07d}" for i in range(N_ORDERS)],
    "customer_id":                np.random.choice(customers["customer_id"], N_ORDERS),
    "order_status":               np.random.choice(STATUS_CHOICES, N_ORDERS),
    "order_purchase_timestamp":   order_timestamps,
    "order_delivered_customer_date": pd.to_datetime(order_timestamps) +
                                      pd.to_timedelta(np.random.randint(3,20,N_ORDERS), unit='d'),
})
orders.to_csv(f"{RAW_DIR}/orders.csv", index=False)
print(f"  orders.csv     → {len(orders):,} rows")

# ─────────────────────────────────────────────
# 4. ORDER ITEMS
# ─────────────────────────────────────────────
print("Generating order_items...")

# Category-based price ranges (realistic)
cat_price = {
    "electronics":           (150, 2500),
    "clothing_accessories":  (30,  400),
    "home_decor":            (25,  600),
    "sports_leisure":        (40,  800),
    "beauty_perfumery":      (20,  300),
    "furniture":             (200,3000),
    "books":                 (15,  150),
    "toys_games":            (20,  350),
    "kitchen_utilities":     (30,  500),
    "health_wellness":       (25,  400),
    "automotive":            (80, 1200),
    "garden_tools":          (35,  600),
    "baby_products":         (20,  300),
    "stationery":            (10,  100),
    "pet_supplies":          (15,  250),
}

order_items_list = []
for _, row in orders.iterrows():
    n_items = np.random.choice([1,2,3,4], p=[0.65,0.20,0.10,0.05])
    prods   = products.sample(n_items)
    for item_no, (_, prod) in enumerate(prods.iterrows(), 1):
        cat = prod["product_category_name"]
        lo, hi = cat_price.get(cat, (30, 500))
        price  = round(np.random.uniform(lo, hi), 2)
        # Freight: 5-20% of price, but minimum 8
        freight = round(max(8, price * np.random.uniform(0.05, 0.20)), 2)
        order_items_list.append({
            "order_id":      row["order_id"],
            "order_item_id": item_no,
            "product_id":    prod["product_id"],
            "price":         price,
            "freight_value": freight,
        })

order_items = pd.DataFrame(order_items_list)
order_items.to_csv(f"{RAW_DIR}/order_items.csv", index=False)
print(f"  order_items.csv → {len(order_items):,} rows")

# ─────────────────────────────────────────────
# 5. PAYMENTS
# ─────────────────────────────────────────────
print("Generating payments...")
# Aggregate per order
pay_agg = order_items.groupby("order_id").agg(
    payment_value=("price","sum")).reset_index()
pay_agg["payment_type"]         = np.random.choice(
    ["credit_card","boleto","debit_card","voucher"],
    len(pay_agg), p=[0.60,0.19,0.13,0.08])
pay_agg["payment_installments"] = np.random.choice(
    [1,2,3,6,10,12], len(pay_agg), p=[0.40,0.15,0.15,0.15,0.10,0.05])
pay_agg.to_csv(f"{RAW_DIR}/payments.csv", index=False)
print(f"  payments.csv   → {len(pay_agg):,} rows")

# ─────────────────────────────────────────────
# 6. REVIEWS
# ─────────────────────────────────────────────
print("Generating reviews...")
# ~70% of orders have reviews
review_orders = orders.sample(frac=0.70)
reviews = pd.DataFrame({
    "review_id":    [f"REV_{i:07d}" for i in range(len(review_orders))],
    "order_id":     review_orders["order_id"].values,
    "review_score": np.random.choice([1,2,3,4,5],
                        len(review_orders), p=[0.05,0.07,0.12,0.28,0.48]),
})
reviews.to_csv(f"{RAW_DIR}/reviews.csv", index=False)
print(f"  reviews.csv    → {len(reviews):,} rows")

print("\nAll raw data files generated successfully!")
