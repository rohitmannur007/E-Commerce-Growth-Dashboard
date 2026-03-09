"""
STEP 03 — Customer Segmentation & RFM Analysis
===============================================
Purpose:
  - Segment customers by order count (New / Active / Loyal)
  - Build full RFM (Recency, Frequency, Monetary) scoring
  - Pareto analysis (what % of customers drive 80% of revenue)
  - Customer Lifetime Value (CLV) estimation
  - Generate charts

Run: python3 python_analysis/03_customer_segmentation.py
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings("ignore")
import os

BASE     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROC_DIR = os.path.join(BASE, "data", "processed")
CHARTS   = os.path.join(BASE, "outputs", "charts")
CSV_OUT  = os.path.join(BASE, "outputs", "csv_exports")

BLUE   = "#2E86AB"; GREEN  = "#28A745"; ORANGE = "#FD7E14"
RED    = "#DC3545"; PURPLE = "#6F42C1"; TEAL   = "#20C997"
PINK   = "#E83E8C"; YELLOW = "#FFC107"

plt.rcParams.update({
    "figure.facecolor":"#F8F9FA","axes.facecolor":"#F8F9FA",
    "axes.grid":True,"grid.alpha":0.4,
    "axes.spines.top":False,"axes.spines.right":False,
})

print("=" * 60)
print("STEP 03 — Customer Segmentation & RFM Analysis")
print("=" * 60)

fact = pd.read_csv(f"{PROC_DIR}/orders_fact.csv",
                   parse_dates=["order_purchase_timestamp"])

# ─────────────────────────────────────────────
# 1. CUSTOMER BASE METRICS
# ─────────────────────────────────────────────
print("\n[1] Computing customer base metrics...")

cust = (
    fact.groupby("customer_id")
    .agg(
        total_orders     = ("order_id",                  "nunique"),
        total_revenue    = ("price",                     "sum"),
        total_freight    = ("freight_value",             "sum"),
        total_profit     = ("profit_estimate",           "sum"),
        avg_order_value  = ("price",                     "mean"),
        first_order      = ("order_purchase_timestamp",  "min"),
        last_order       = ("order_purchase_timestamp",  "max"),
    )
    .reset_index()
)
cust["first_order"] = pd.to_datetime(cust["first_order"])
cust["last_order"]  = pd.to_datetime(cust["last_order"])

# Reference date = last date in dataset
ref_date = cust["last_order"].max()
cust["recency_days"]       = (ref_date - cust["last_order"]).dt.days
cust["customer_lifetime_days"] = (cust["last_order"] - cust["first_order"]).dt.days

# Segment
def segment(n):
    if n == 1:   return "New"
    if n <= 4:   return "Active"
    return "Loyal"

cust["customer_segment"] = cust["total_orders"].apply(segment)

print(f"  Total customers: {len(cust):,}")
seg_counts = cust["customer_segment"].value_counts()
for seg, cnt in seg_counts.items():
    rev = cust[cust["customer_segment"]==seg]["total_revenue"].sum()
    print(f"  {seg:8s}: {cnt:6,} customers  |  Revenue R$ {rev:,.0f}")

cust.to_csv(f"{CSV_OUT}/customer_metrics.csv", index=False)

# ─────────────────────────────────────────────
# 2. RFM SCORING
# ─────────────────────────────────────────────
print("\n[2] RFM scoring...")

rfm = cust[["customer_id","recency_days","total_orders","total_revenue"]].copy()
rfm.columns = ["customer_id","recency","frequency","monetary"]

# Score 1-5
rfm["r_score"] = pd.qcut(rfm["recency"],   q=5, labels=[5,4,3,2,1]).astype(int)
rfm["f_score"] = pd.qcut(rfm["frequency"].rank(method="first"), q=5, labels=[1,2,3,4,5]).astype(int)
rfm["m_score"] = pd.qcut(rfm["monetary"],  q=5, labels=[1,2,3,4,5]).astype(int)
rfm["rfm_total"] = rfm["r_score"] + rfm["f_score"] + rfm["m_score"]

def rfm_label(row):
    if row["rfm_total"] >= 13:                          return "Champions"
    if row["rfm_total"] >= 10 and row["r_score"] >= 3:  return "Loyal Customers"
    if row["rfm_total"] >= 9  and row["r_score"] >= 4:  return "Potential Loyalists"
    if row["r_score"] == 5    and row["rfm_total"] < 8: return "New Customers"
    if row["r_score"] >= 3    and row["f_score"] <= 2:  return "Promising"
    if row["r_score"] <= 2    and row["rfm_total"] >= 9: return "At Risk"
    if row["r_score"] <= 2    and row["rfm_total"] >= 7: return "Cannot Lose Them"
    if row["r_score"] <= 2    and row["rfm_total"] < 7:  return "Lost"
    return "Need Attention"

rfm["rfm_segment"] = rfm.apply(rfm_label, axis=1)
rfm.to_csv(f"{CSV_OUT}/rfm_segments.csv", index=False)

rfm_summary = (
    rfm.groupby("rfm_segment")
    .agg(count=("customer_id","count"), avg_monetary=("monetary","mean"),
         avg_frequency=("frequency","mean"), avg_recency=("recency","mean"))
    .sort_values("avg_monetary", ascending=False)
    .reset_index()
)
print(rfm_summary.to_string(index=False))
rfm_summary.to_csv(f"{CSV_OUT}/rfm_segment_summary.csv", index=False)

# ─────────────────────────────────────────────
# 3. PARETO ANALYSIS (80/20 rule)
# ─────────────────────────────────────────────
print("\n[3] Pareto analysis...")

cust_sorted = cust.sort_values("total_revenue", ascending=False).reset_index(drop=True)
cust_sorted["cum_revenue_pct"] = cust_sorted["total_revenue"].cumsum() / cust_sorted["total_revenue"].sum() * 100
cust_sorted["cum_customer_pct"] = (cust_sorted.index + 1) / len(cust_sorted) * 100

# Find what % of customers generate 80% of revenue
p80 = cust_sorted[cust_sorted["cum_revenue_pct"] >= 80].iloc[0]["cum_customer_pct"]
print(f"  {p80:.1f}% of customers generate 80% of revenue  (Pareto principle)")

fig, ax = plt.subplots(figsize=(10, 6))
ax.plot(cust_sorted["cum_customer_pct"], cust_sorted["cum_revenue_pct"],
        color=BLUE, linewidth=2.5)
ax.axhline(80, color=RED, linestyle="--", linewidth=1.2, alpha=0.8, label="80% Revenue")
ax.axvline(p80, color=ORANGE, linestyle="--", linewidth=1.2, alpha=0.8,
           label=f"{p80:.0f}% Customers")
ax.fill_between(cust_sorted["cum_customer_pct"], cust_sorted["cum_revenue_pct"],
                alpha=0.15, color=BLUE)
ax.set_xlabel("Cumulative % of Customers")
ax.set_ylabel("Cumulative % of Revenue")
ax.set_title("Pareto Chart — Customer Revenue Concentration", fontsize=14, fontweight="bold")
ax.legend()
ax.set_xlim(0,100); ax.set_ylim(0,100)
plt.tight_layout()
plt.savefig(f"{CHARTS}/03a_pareto_analysis.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/03a_pareto_analysis.png")

# ─────────────────────────────────────────────
# 4. CUSTOMER SEGMENT CHART
# ─────────────────────────────────────────────
seg_rev = cust.groupby("customer_segment")["total_revenue"].sum()
seg_cnt = cust["customer_segment"].value_counts()

fig, axes = plt.subplots(1, 2, figsize=(12, 5))
fig.suptitle("Customer Segment Analysis", fontsize=14, fontweight="bold")

colors_seg = [GREEN, BLUE, ORANGE]
axes[0].pie(seg_cnt, labels=seg_cnt.index, autopct="%1.1f%%",
            colors=colors_seg, startangle=90)
axes[0].set_title("Customers by Segment")

axes[1].bar(seg_rev.index, seg_rev.values / 1e6, color=colors_seg, alpha=0.85, edgecolor="white")
axes[1].set_title("Revenue by Segment (R$ M)")
axes[1].set_xlabel("Segment")
axes[1].set_ylabel("Revenue (R$ Millions)")

plt.tight_layout()
plt.savefig(f"{CHARTS}/03b_customer_segments.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/03b_customer_segments.png")

# ─────────────────────────────────────────────
# 5. RFM SEGMENT CHART
# ─────────────────────────────────────────────
seg_colors = [PURPLE, BLUE, GREEN, TEAL, YELLOW, ORANGE, PINK, RED][:len(rfm_summary)]

fig, ax = plt.subplots(figsize=(12, 6))
bars = ax.barh(rfm_summary["rfm_segment"], rfm_summary["count"], color=seg_colors, alpha=0.85)
ax.set_xlabel("Number of Customers")
ax.set_title("RFM Customer Segments — Distribution", fontsize=14, fontweight="bold")
for bar, val in zip(bars, rfm_summary["count"]):
    ax.text(bar.get_width() + 10, bar.get_y() + bar.get_height()/2,
            f"{val:,}", va="center", fontsize=9)
plt.tight_layout()
plt.savefig(f"{CHARTS}/03c_rfm_segments.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/03c_rfm_segments.png")

# ─────────────────────────────────────────────
# 6. CLV ESTIMATION
# ─────────────────────────────────────────────
print("\n[4] Customer Lifetime Value estimation...")

# Simple CLV = AOV × Purchase Frequency × Average Customer Lifetime (years)
aov           = cust["total_revenue"].sum() / cust["total_orders"].sum()
avg_frequency = cust["total_orders"].mean()  # orders per customer
avg_lifetime  = cust[cust["customer_lifetime_days"] > 0]["customer_lifetime_days"].mean() / 365

clv_simple = aov * avg_frequency * avg_lifetime
print(f"  Avg Order Value          : R$ {aov:,.2f}")
print(f"  Avg Orders per Customer  : {avg_frequency:.2f}")
print(f"  Avg Customer Lifetime    : {avg_lifetime:.2f} years")
print(f"  Simple CLV Estimate      : R$ {clv_simple:,.2f}")

# CLV by segment
clv_seg = cust.groupby("customer_segment").agg(
    avg_revenue   = ("total_revenue","mean"),
    avg_orders    = ("total_orders","mean"),
    avg_lifetime  = ("customer_lifetime_days","mean")
).reset_index()
clv_seg["clv_estimate"] = clv_seg["avg_revenue"]
clv_seg.to_csv(f"{CSV_OUT}/clv_by_segment.csv", index=False)
print(clv_seg.to_string(index=False))

print("\n✅ STEP 03 COMPLETE\n")
