"""
STEP 06 — Statistical Analysis
================================
Purpose:
  - Correlation analysis (price vs freight, review vs revenue)
  - Distribution analysis (order value, price by category)
  - Pareto verification with stats
  - 80/20 rule validation
  - Review score impact on repeat purchase

Run: python3 python_analysis/06_statistical_analysis.py
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import scipy.stats as stats
import warnings
warnings.filterwarnings("ignore")
import os

BASE     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROC_DIR = os.path.join(BASE, "data", "processed")
CHARTS   = os.path.join(BASE, "outputs", "charts")
CSV_OUT  = os.path.join(BASE, "outputs", "csv_exports")

BLUE  = "#2E86AB"; GREEN = "#28A745"; ORANGE = "#FD7E14"
RED   = "#DC3545"; GRAY  = "#6C757D"

plt.rcParams.update({
    "figure.facecolor":"#F8F9FA","axes.facecolor":"#F8F9FA",
    "axes.grid":True,"grid.alpha":0.4,
    "axes.spines.top":False,"axes.spines.right":False,
})

print("=" * 60)
print("STEP 06 — Statistical Analysis")
print("=" * 60)

fact = pd.read_csv(f"{PROC_DIR}/orders_fact.csv",
                   parse_dates=["order_purchase_timestamp"])

# ─────────────────────────────────────────────
# 1. CORRELATION MATRIX
# ─────────────────────────────────────────────
print("\n[1] Correlation analysis...")

corr_df = fact[["price","freight_value","profit_estimate",
                "review_score","delivery_days","payment_installments"]].copy()
corr_df = corr_df[corr_df["review_score"] > 0]   # exclude no-review orders

corr_matrix = corr_df.corr().round(3)
corr_matrix.to_csv(f"{CSV_OUT}/correlation_matrix.csv")

# Specific correlations of interest
c1 = corr_df["price"].corr(corr_df["freight_value"])
c2 = corr_df["review_score"].corr(corr_df["price"])
c3 = corr_df["delivery_days"].corr(corr_df["review_score"])

print(f"  Correlation: Price vs Freight            : {c1:.3f}")
print(f"  Correlation: Review Score vs Price       : {c2:.3f}")
print(f"  Correlation: Delivery Days vs Review     : {c3:.3f}")

print("\n  Interpretation:")
if abs(c1) > 0.5:
    print(f"  → Price and freight are {'positively' if c1 > 0 else 'negatively'} "
          f"correlated ({c1:.2f}). Expensive items carry higher shipping costs.")
if c3 < -0.1:
    print(f"  → Longer delivery = lower review scores ({c3:.2f}). "
          f"Fast delivery improves customer satisfaction.")

# Heatmap
fig, ax = plt.subplots(figsize=(8, 6))
im = ax.imshow(corr_matrix.values, cmap="RdYlGn", vmin=-1, vmax=1, aspect="auto")
ax.set_xticks(range(len(corr_matrix.columns)))
ax.set_yticks(range(len(corr_matrix.columns)))
ax.set_xticklabels(corr_matrix.columns, rotation=45, ha="right", fontsize=9)
ax.set_yticklabels(corr_matrix.columns, fontsize=9)
plt.colorbar(im, ax=ax, shrink=0.8)
for i in range(len(corr_matrix)):
    for j in range(len(corr_matrix)):
        ax.text(j, i, f"{corr_matrix.values[i,j]:.2f}",
                ha="center", va="center", fontsize=8,
                color="black" if abs(corr_matrix.values[i,j]) < 0.7 else "white")
ax.set_title("Correlation Matrix — Key Metrics", fontsize=13, fontweight="bold")
plt.tight_layout()
plt.savefig(f"{CHARTS}/06a_correlation_matrix.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/06a_correlation_matrix.png")

# ─────────────────────────────────────────────
# 2. ORDER VALUE DISTRIBUTION
# ─────────────────────────────────────────────
print("\n[2] Order value distribution...")

order_values = fact.groupby("order_id")["price"].sum()

p25, p50, p75, p90, p95 = np.percentile(order_values, [25,50,75,90,95])
print(f"  Order Value Distribution:")
print(f"    Min      : R$ {order_values.min():,.2f}")
print(f"    25th pct : R$ {p25:,.2f}")
print(f"    Median   : R$ {p50:,.2f}")
print(f"    Mean     : R$ {order_values.mean():,.2f}")
print(f"    75th pct : R$ {p75:,.2f}")
print(f"    90th pct : R$ {p90:,.2f}")
print(f"    95th pct : R$ {p95:,.2f}")
print(f"    Max      : R$ {order_values.max():,.2f}")

pd.DataFrame({
    "Percentile":["Min","25th","Median","Mean","75th","90th","95th","Max"],
    "Value (R$)": [order_values.min(),p25,p50,order_values.mean(),p75,p90,p95,order_values.max()]
}).to_csv(f"{CSV_OUT}/order_value_distribution.csv", index=False)

fig, axes = plt.subplots(1, 2, figsize=(14, 5))
fig.suptitle("Order Value Distribution", fontsize=14, fontweight="bold")

# Histogram (cap at 95th percentile for readability)
capped = order_values[order_values <= p95]
axes[0].hist(capped, bins=50, color=BLUE, alpha=0.75, edgecolor="white")
axes[0].axvline(p50, color=RED, linestyle="--", linewidth=1.5, label=f"Median R$ {p50:,.0f}")
axes[0].axvline(order_values.mean(), color=ORANGE, linestyle="--", linewidth=1.5,
                label=f"Mean R$ {order_values.mean():,.0f}")
axes[0].set_xlabel("Order Value (R$)")
axes[0].set_ylabel("Frequency")
axes[0].set_title("Order Value Histogram (capped at 95th pct)")
axes[0].legend()

# Box plot per category
cat_prices = fact.groupby("product_category_name")["price"].apply(list)
top8_cats = fact.groupby("product_category_name")["price"].median().nlargest(8).index
bp_data = [fact[fact["product_category_name"]==c]["price"].clip(upper=2000).values
           for c in top8_cats]
axes[1].boxplot(bp_data, labels=[c[:12] for c in top8_cats], patch_artist=True,
                boxprops=dict(facecolor=BLUE, alpha=0.6),
                medianprops=dict(color=RED, linewidth=2))
axes[1].set_xticklabels([c[:12] for c in top8_cats], rotation=35, ha="right", fontsize=8)
axes[1].set_ylabel("Price (R$, capped at 2000)")
axes[1].set_title("Price Distribution by Category")

plt.tight_layout()
plt.savefig(f"{CHARTS}/06b_order_value_distribution.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/06b_order_value_distribution.png")

# ─────────────────────────────────────────────
# 3. REVIEW SCORE ANALYSIS
# ─────────────────────────────────────────────
print("\n[3] Review score analysis...")

reviews = fact[fact["review_score"] > 0].drop_duplicates("order_id")
rev_dist = reviews["review_score"].value_counts().sort_index()
rev_pct  = rev_dist / rev_dist.sum() * 100

print("  Review Score Distribution:")
for score, pct in rev_pct.items():
    bar = "█" * int(pct / 2)
    print(f"    {score} stars : {pct:5.1f}%  {bar}")

# Review score vs category
cat_review = (
    fact[fact["review_score"] > 0]
    .groupby("product_category_name")["review_score"]
    .mean()
    .sort_values(ascending=False)
)
cat_review.to_csv(f"{CSV_OUT}/review_by_category.csv")

fig, axes = plt.subplots(1, 2, figsize=(14, 5))
fig.suptitle("Review Score Analysis", fontsize=14, fontweight="bold")

colors_bar = [RED if i < 2 else ORANGE if i == 2 else GREEN for i in rev_dist.index - 1]
axes[0].bar(rev_dist.index.astype(str), rev_dist.values,
            color=colors_bar, alpha=0.85, edgecolor="white")
axes[0].set_xlabel("Review Score"); axes[0].set_ylabel("Number of Reviews")
axes[0].set_title("Review Score Distribution")

colors_cat = [GREEN if v >= 4.0 else ORANGE if v >= 3.5 else RED for v in cat_review.values]
axes[1].barh(cat_review.index, cat_review.values, color=colors_cat, alpha=0.85, edgecolor="white")
axes[1].axvline(4.0, color=GREEN, linestyle="--", alpha=0.7, linewidth=1.2)
axes[1].set_xlabel("Avg Review Score")
axes[1].set_title("Avg Review Score by Category")
axes[1].set_xlim(0, 5)

plt.tight_layout()
plt.savefig(f"{CHARTS}/06c_review_analysis.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/06c_review_analysis.png")

# ─────────────────────────────────────────────
# 4. PRICE vs FREIGHT SCATTER
# ─────────────────────────────────────────────
print("\n[4] Price vs Freight scatter...")

sample = fact.sample(3000, random_state=42)
slope, intercept, r_value, p_value, std_err = stats.linregress(
    sample["price"], sample["freight_value"])

fig, ax = plt.subplots(figsize=(10, 6))
ax.scatter(sample["price"], sample["freight_value"],
           alpha=0.3, s=8, color=BLUE)
x_line = np.linspace(0, sample["price"].quantile(0.95), 100)
ax.plot(x_line, intercept + slope * x_line, color=RED, linewidth=2,
        label=f"Trend (R²={r_value**2:.3f})")
ax.set_xlim(0, sample["price"].quantile(0.95))
ax.set_ylim(0, sample["freight_value"].quantile(0.98))
ax.set_xlabel("Item Price (R$)"); ax.set_ylabel("Freight Value (R$)")
ax.set_title("Price vs Freight Cost — Correlation", fontsize=13, fontweight="bold")
ax.legend()
plt.tight_layout()
plt.savefig(f"{CHARTS}/06d_price_vs_freight.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/06d_price_vs_freight.png")
print(f"  R-squared: {r_value**2:.3f}  |  Slope: {slope:.4f}")
print(f"  Interpretation: Every R$ 1 increase in price → R$ {slope:.2f} increase in freight")

print("\n✅ STEP 06 COMPLETE\n")
