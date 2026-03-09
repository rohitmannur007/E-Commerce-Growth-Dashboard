"""
STEP 04 — Cohort Retention Analysis
=====================================
Purpose:
  - Build cohort retention matrix (what % of customers come back month-over-month)
  - Visualise as heatmap
  - Calculate average retention rates across cohorts
  - Export retention table

Run: python3 python_analysis/04_cohort_analysis.py
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
warnings.filterwarnings("ignore")
import os

BASE     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROC_DIR = os.path.join(BASE, "data", "processed")
CHARTS   = os.path.join(BASE, "outputs", "charts")
CSV_OUT  = os.path.join(BASE, "outputs", "csv_exports")

print("=" * 60)
print("STEP 04 — Cohort Retention Analysis")
print("=" * 60)

fact = pd.read_csv(f"{PROC_DIR}/orders_fact.csv",
                   parse_dates=["order_purchase_timestamp"])

# ─────────────────────────────────────────────
# 1. BUILD COHORT ASSIGNMENTS
# ─────────────────────────────────────────────
print("\n[1] Assigning cohort months...")

df = fact[["customer_id","order_id","order_purchase_timestamp","price"]].drop_duplicates()

# Cohort = first purchase month per customer
df["order_month"] = df["order_purchase_timestamp"].dt.to_period("M")
first_purchase    = df.groupby("customer_id")["order_month"].min().rename("cohort_month")
df = df.merge(first_purchase, on="customer_id")

# Month index (0 = acquisition month)
df["month_index"] = (df["order_month"] - df["cohort_month"]).apply(lambda x: x.n)

print(f"  Cohorts identified: {df['cohort_month'].nunique()}")
print(f"  Max month index   : {df['month_index'].max()}")

# ─────────────────────────────────────────────
# 2. RETENTION MATRIX
# ─────────────────────────────────────────────
print("\n[2] Building retention matrix...")

cohort_data = (
    df.groupby(["cohort_month","month_index"])["customer_id"]
    .nunique()
    .reset_index(name="customers")
)

cohort_pivot = cohort_data.pivot_table(
    index="cohort_month", columns="month_index", values="customers"
)

# Cohort size = month 0 column
cohort_sizes = cohort_pivot[0]

# Retention rates
retention = cohort_pivot.divide(cohort_sizes, axis=0) * 100
retention = retention.round(1)

# Limit to first 12 months for readability
retention = retention.iloc[:, :13]

print(f"  Retention matrix shape: {retention.shape}")
retention.to_csv(f"{CSV_OUT}/cohort_retention_matrix.csv")
print("  Saved: outputs/csv_exports/cohort_retention_matrix.csv")

# ─────────────────────────────────────────────
# 3. HEATMAP
# ─────────────────────────────────────────────
print("\n[3] Generating cohort heatmap...")

fig, ax = plt.subplots(figsize=(16, 10))

# Display only cohorts with enough data (at least 3 months)
display_ret = retention.dropna(thresh=4)

sns.heatmap(
    display_ret,
    annot=True,
    fmt=".0f",
    cmap="YlOrRd_r",
    vmin=0,
    vmax=100,
    linewidths=0.5,
    ax=ax,
    cbar_kws={"label": "Retention %"},
    annot_kws={"size": 8}
)

ax.set_title("Cohort Retention Matrix — % of Customers Returning Each Month",
             fontsize=14, fontweight="bold", pad=15)
ax.set_xlabel("Month Index (0 = Acquisition Month)", fontsize=11)
ax.set_ylabel("Cohort (First Purchase Month)", fontsize=11)

# Format y-axis labels
yticklabels = [str(l) for l in display_ret.index]
ax.set_yticklabels(yticklabels, rotation=0, fontsize=8)

plt.tight_layout()
plt.savefig(f"{CHARTS}/04a_cohort_heatmap.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/04a_cohort_heatmap.png")

# ─────────────────────────────────────────────
# 4. AVERAGE RETENTION CURVE
# ─────────────────────────────────────────────
print("\n[4] Average retention curve...")

avg_retention = retention.mean(axis=0)

fig, ax = plt.subplots(figsize=(12, 6))
ax.plot(avg_retention.index, avg_retention.values,
        marker="o", color="#2E86AB", linewidth=2.5, markersize=7)
ax.fill_between(avg_retention.index, avg_retention.values, alpha=0.2, color="#2E86AB")

for x, y in zip(avg_retention.index, avg_retention.values):
    ax.annotate(f"{y:.0f}%", (x, y), textcoords="offset points",
                xytext=(0, 10), ha="center", fontsize=9, color="#2E86AB")

ax.set_xlabel("Months After First Purchase", fontsize=11)
ax.set_ylabel("Average Retention %", fontsize=11)
ax.set_title("Average Customer Retention Curve", fontsize=14, fontweight="bold")
ax.set_ylim(0, 110)
ax.set_xticks(avg_retention.index)
ax.set_facecolor("#F8F9FA")
ax.grid(alpha=0.4)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)

plt.tight_layout()
plt.savefig(f"{CHARTS}/04b_retention_curve.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/04b_retention_curve.png")

# ─────────────────────────────────────────────
# 5. COHORT REVENUE
# ─────────────────────────────────────────────
print("\n[5] Cohort revenue analysis...")

cohort_rev = (
    df.groupby(["cohort_month","month_index"])["price"]
    .sum()
    .reset_index()
)
cohort_rev_pivot = cohort_rev.pivot_table(
    index="cohort_month", columns="month_index", values="price"
).iloc[:, :13]
cohort_rev_pivot.to_csv(f"{CSV_OUT}/cohort_revenue_matrix.csv")
print("  Saved: outputs/csv_exports/cohort_revenue_matrix.csv")

# ─────────────────────────────────────────────
# 6. KEY INSIGHTS
# ─────────────────────────────────────────────
print("\n[6] Retention insights:")
m1_ret = avg_retention.get(1, np.nan)
m3_ret = avg_retention.get(3, np.nan)
m6_ret = avg_retention.get(6, np.nan)
print(f"  Month 0 (acquisition)    : 100.0%")
print(f"  Month 1 retention        : {m1_ret:.1f}%")
print(f"  Month 3 retention        : {m3_ret:.1f}%")
print(f"  Month 6 retention        : {m6_ret:.1f}%")
if not np.isnan(m1_ret):
    print(f"  Drop-off Month 0→1       : {100 - m1_ret:.1f}% of customers lost after first purchase")

print("\n✅ STEP 04 COMPLETE\n")
