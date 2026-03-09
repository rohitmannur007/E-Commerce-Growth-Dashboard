"""
STEP 05 — Product & Category Profitability Analysis
=====================================================
Purpose:
  - Revenue vs profit comparison by category
  - Identify "revenue traps" (high revenue, low profit)
  - Top 10 and bottom 10 products by profit margin
  - Freight cost impact analysis
  - Review score correlation with revenue

Run: python3 python_analysis/05_product_profitability.py
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import warnings
warnings.filterwarnings("ignore")
import os

BASE     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROC_DIR = os.path.join(BASE, "data", "processed")
CHARTS   = os.path.join(BASE, "outputs", "charts")
CSV_OUT  = os.path.join(BASE, "outputs", "csv_exports")

BLUE   = "#2E86AB"; GREEN  = "#28A745"; ORANGE = "#FD7E14"
RED    = "#DC3545"; PURPLE = "#6F42C1"; GRAY   = "#6C757D"

plt.rcParams.update({
    "figure.facecolor":"#F8F9FA","axes.facecolor":"#F8F9FA",
    "axes.grid":True,"grid.alpha":0.4,
    "axes.spines.top":False,"axes.spines.right":False,
})

print("=" * 60)
print("STEP 05 — Product & Category Profitability")
print("=" * 60)

fact = pd.read_csv(f"{PROC_DIR}/orders_fact.csv",
                   parse_dates=["order_purchase_timestamp"])

# ─────────────────────────────────────────────
# 1. CATEGORY METRICS
# ─────────────────────────────────────────────
print("\n[1] Category profitability metrics...")

cat = (
    fact.groupby("product_category_name")
    .agg(
        total_orders     = ("order_id",          "nunique"),
        total_revenue    = ("price",             "sum"),
        total_freight    = ("freight_value",     "sum"),
        total_profit     = ("profit_estimate",   "sum"),
        avg_price        = ("price",             "mean"),
        avg_freight      = ("freight_value",     "mean"),
        avg_review       = ("review_score",      lambda x: x[x>0].mean()),
        unique_products  = ("product_id",        "nunique"),
    )
    .reset_index()
)

cat["profit_margin_pct"]    = (cat["total_profit"] / cat["total_revenue"] * 100).round(2)
cat["freight_pct_of_rev"]   = (cat["total_freight"] / cat["total_revenue"] * 100).round(2)
cat["revenue_share_pct"]    = (cat["total_revenue"] / cat["total_revenue"].sum() * 100).round(2)
cat = cat.sort_values("total_revenue", ascending=False).reset_index(drop=True)
cat.to_csv(f"{CSV_OUT}/category_metrics.csv", index=False)

print(cat[["product_category_name","total_revenue","total_profit",
           "profit_margin_pct","freight_pct_of_rev"]].to_string(index=False))

# ─────────────────────────────────────────────
# 2. REVENUE vs PROFIT BY CATEGORY (grouped bar)
# ─────────────────────────────────────────────
print("\n[2] Revenue vs Profit chart...")

top_cat = cat.head(12)
x       = np.arange(len(top_cat))
w       = 0.38

fig, ax = plt.subplots(figsize=(14, 7))
bars1 = ax.bar(x - w/2, top_cat["total_revenue"] / 1e6, w, label="Revenue",
               color=BLUE, alpha=0.85, edgecolor="white")
bars2 = ax.bar(x + w/2, top_cat["total_profit"]  / 1e6, w, label="Profit Est.",
               color=GREEN, alpha=0.85, edgecolor="white")

ax.set_xlabel("Category", fontsize=11)
ax.set_ylabel("R$ Millions", fontsize=11)
ax.set_title("Revenue vs Profit by Category", fontsize=14, fontweight="bold")
ax.set_xticks(x)
ax.set_xticklabels(top_cat["product_category_name"], rotation=35, ha="right", fontsize=9)
ax.legend()

# Annotate profit margin
for i, row in top_cat.iterrows():
    ax.text(i + w/2, row["total_profit"]/1e6 + 0.05,
            f"{row['profit_margin_pct']:.0f}%",
            ha="center", va="bottom", fontsize=7.5, color="#155724")

plt.tight_layout()
plt.savefig(f"{CHARTS}/05a_revenue_vs_profit_category.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/05a_revenue_vs_profit_category.png")

# ─────────────────────────────────────────────
# 3. FREIGHT AS % OF REVENUE (which categories
#    have high shipping burden?)
# ─────────────────────────────────────────────
print("\n[3] Freight burden by category...")

cat_sorted_freight = cat.sort_values("freight_pct_of_rev", ascending=True)

fig, ax = plt.subplots(figsize=(12, 7))
colors_bar = [RED if v > 15 else ORANGE if v > 10 else GREEN
              for v in cat_sorted_freight["freight_pct_of_rev"]]
bars = ax.barh(cat_sorted_freight["product_category_name"],
               cat_sorted_freight["freight_pct_of_rev"],
               color=colors_bar, alpha=0.85, edgecolor="white")
ax.axvline(10, color=ORANGE, linestyle="--", linewidth=1.2, alpha=0.8, label="10% threshold")
ax.axvline(15, color=RED,    linestyle="--", linewidth=1.2, alpha=0.8, label="15% threshold")
ax.set_xlabel("Freight Cost as % of Revenue")
ax.set_title("Freight Burden by Category\n(Higher % = More Shipping Eats into Profit)",
             fontsize=13, fontweight="bold")

green_p = mpatches.Patch(color=GREEN,  label="Efficient  (<10%)")
orange_p= mpatches.Patch(color=ORANGE, label="Warning   (10–15%)")
red_p   = mpatches.Patch(color=RED,    label="Critical   (>15%)")
ax.legend(handles=[green_p, orange_p, red_p], loc="lower right")

plt.tight_layout()
plt.savefig(f"{CHARTS}/05b_freight_burden.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/05b_freight_burden.png")

# ─────────────────────────────────────────────
# 4. TOP & BOTTOM PRODUCTS BY MARGIN
# ─────────────────────────────────────────────
print("\n[4] Top/Bottom products by profit margin...")

prod = (
    fact.groupby(["product_id","product_category_name"])
    .agg(
        total_revenue = ("price",           "sum"),
        total_profit  = ("profit_estimate", "sum"),
        total_orders  = ("order_id",        "nunique"),
    )
    .reset_index()
)
prod["profit_margin_pct"] = (prod["total_profit"] / prod["total_revenue"] * 100).round(2)
prod = prod[prod["total_revenue"] > 500]   # only meaningful products

top10    = prod.nlargest(10, "profit_margin_pct")
bottom10 = prod.nsmallest(10, "profit_margin_pct")

top10.to_csv(f"{CSV_OUT}/top10_products_margin.csv", index=False)
bottom10.to_csv(f"{CSV_OUT}/bottom10_products_margin.csv", index=False)
prod.to_csv(f"{CSV_OUT}/product_metrics.csv", index=False)

fig, axes = plt.subplots(1, 2, figsize=(14, 6))
fig.suptitle("Product Profit Margin Analysis", fontsize=14, fontweight="bold")

axes[0].barh(top10["product_id"].str[-6:], top10["profit_margin_pct"],
             color=GREEN, alpha=0.85)
axes[0].set_title("Top 10 — Highest Profit Margin %")
axes[0].set_xlabel("Profit Margin %")

axes[1].barh(bottom10["product_id"].str[-6:], bottom10["profit_margin_pct"],
             color=RED, alpha=0.85)
axes[1].set_title("Bottom 10 — Lowest Profit Margin %")
axes[1].set_xlabel("Profit Margin %")

plt.tight_layout()
plt.savefig(f"{CHARTS}/05c_top_bottom_products.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/05c_top_bottom_products.png")

# ─────────────────────────────────────────────
# 5. MONTHLY CATEGORY REVENUE TREND
# ─────────────────────────────────────────────
print("\n[5] Monthly category revenue trend...")

top5_cats = cat.head(5)["product_category_name"].tolist()
monthly_cat = (
    fact[fact["product_category_name"].isin(top5_cats)]
    .groupby(["order_year_month","product_category_name"])["price"]
    .sum()
    .unstack(fill_value=0)
)

fig, ax = plt.subplots(figsize=(14, 7))
palette = [BLUE, GREEN, ORANGE, RED, PURPLE]
for i, col in enumerate(monthly_cat.columns):
    ax.plot(range(len(monthly_cat)), monthly_cat[col] / 1e3,
            label=col, color=palette[i], linewidth=2, marker="o", markersize=3)

ax.set_xticks(range(0, len(monthly_cat), 2))
ax.set_xticklabels(list(monthly_cat.index)[::2], rotation=45, ha="right", fontsize=8)
ax.set_xlabel("Month"); ax.set_ylabel("Revenue (R$ Thousands)")
ax.set_title("Monthly Revenue Trend — Top 5 Categories", fontsize=14, fontweight="bold")
ax.legend(loc="upper left", fontsize=9)
plt.tight_layout()
plt.savefig(f"{CHARTS}/05d_category_monthly_trend.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/05d_category_monthly_trend.png")

# ─────────────────────────────────────────────
# 6. KEY INSIGHTS
# ─────────────────────────────────────────────
print("\n[6] Key insights:")
high_freight = cat[cat["freight_pct_of_rev"] > 15]
print(f"  Categories with freight >15% of revenue: {len(high_freight)}")
for _, row in high_freight.iterrows():
    print(f"    {row['product_category_name']:30s}  freight: {row['freight_pct_of_rev']:.1f}%")

best_margin = cat.nlargest(1,"profit_margin_pct").iloc[0]
worst_margin = cat.nsmallest(1,"profit_margin_pct").iloc[0]
print(f"  Best margin category  : {best_margin['product_category_name']} ({best_margin['profit_margin_pct']:.1f}%)")
print(f"  Worst margin category : {worst_margin['product_category_name']} ({worst_margin['profit_margin_pct']:.1f}%)")

print("\n✅ STEP 05 COMPLETE\n")
