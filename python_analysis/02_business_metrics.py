"""
STEP 02 — Business Metrics & KPI Analysis
==========================================
Purpose:
  - Calculate core KPIs: Revenue, AOV, Repeat Rate
  - Monthly revenue trend analysis
  - Revenue by state (geographic)
  - Payment method analysis
  - Generate charts

Run: python3 python_analysis/02_business_metrics.py
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import warnings
warnings.filterwarnings("ignore")
import os

BASE     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROC_DIR = os.path.join(BASE, "data", "processed")
CHARTS   = os.path.join(BASE, "outputs", "charts")
CSV_OUT  = os.path.join(BASE, "outputs", "csv_exports")
os.makedirs(CHARTS, exist_ok=True)

# ── Style ────────────────────────────────────
plt.rcParams.update({
    "figure.facecolor": "#F8F9FA",
    "axes.facecolor":   "#F8F9FA",
    "axes.grid":        True,
    "grid.alpha":       0.4,
    "axes.spines.top":  False,
    "axes.spines.right":False,
    "font.family":      "DejaVu Sans",
})
BLUE   = "#2E86AB"
GREEN  = "#28A745"
ORANGE = "#FD7E14"
RED    = "#DC3545"
PURPLE = "#6F42C1"

print("=" * 60)
print("STEP 02 — Business Metrics & KPI Analysis")
print("=" * 60)

# ─────────────────────────────────────────────
# LOAD
# ─────────────────────────────────────────────
fact = pd.read_csv(f"{PROC_DIR}/orders_fact.csv", parse_dates=["order_purchase_timestamp"])

# ─────────────────────────────────────────────
# 1. CORE KPIs
# ─────────────────────────────────────────────
print("\n[1] Core KPIs")

total_revenue    = fact["price"].sum()
total_freight    = fact["freight_value"].sum()
total_profit_est = fact["profit_estimate"].sum()
total_orders     = fact["order_id"].nunique()
total_customers  = fact["customer_id"].nunique()
aov              = fact.groupby("order_id")["price"].sum().mean()
profit_margin    = total_profit_est / total_revenue * 100

# Repeat purchase rate
orders_per_cust  = fact.groupby("customer_id")["order_id"].nunique()
repeat_rate      = (orders_per_cust > 1).sum() / len(orders_per_cust) * 100

kpis = {
    "Total Revenue (R$)":     f"{total_revenue:,.2f}",
    "Total Freight (R$)":     f"{total_freight:,.2f}",
    "Total Profit Est (R$)":  f"{total_profit_est:,.2f}",
    "Profit Margin %":        f"{profit_margin:.1f}%",
    "Total Orders":           f"{total_orders:,}",
    "Total Customers":        f"{total_customers:,}",
    "Avg Order Value (R$)":   f"{aov:,.2f}",
    "Repeat Purchase Rate %": f"{repeat_rate:.1f}%",
}
for k, v in kpis.items():
    print(f"  {k:<30}: {v}")

pd.DataFrame(list(kpis.items()), columns=["KPI","Value"]).to_csv(
    f"{CSV_OUT}/core_kpis.csv", index=False)

# ─────────────────────────────────────────────
# 2. MONTHLY REVENUE TREND
# ─────────────────────────────────────────────
print("\n[2] Monthly revenue trend...")

monthly = (
    fact.groupby("order_year_month")
    .agg(revenue=("price","sum"), profit=("profit_estimate","sum"),
         orders=("order_id","nunique"))
    .reset_index()
    .sort_values("order_year_month")
)

# MoM growth
monthly["mom_growth_pct"] = monthly["revenue"].pct_change() * 100

monthly.to_csv(f"{CSV_OUT}/monthly_revenue.csv", index=False)

# Chart
fig, axes = plt.subplots(2, 1, figsize=(14, 9), sharex=True)
fig.suptitle("Monthly Revenue & Profit Trend", fontsize=16, fontweight="bold", y=0.98)

x = range(len(monthly))
axes[0].fill_between(x, monthly["revenue"] / 1e6, alpha=0.3, color=BLUE)
axes[0].plot(x, monthly["revenue"] / 1e6, color=BLUE, linewidth=2.5, marker="o", markersize=4)
axes[0].fill_between(x, monthly["profit"] / 1e6, alpha=0.3, color=GREEN)
axes[0].plot(x, monthly["profit"] / 1e6, color=GREEN, linewidth=2.5, marker="s", markersize=4)
axes[0].set_ylabel("R$ Millions")
axes[0].legend(["Revenue", "Profit Estimate"], loc="upper left")
axes[0].set_xticks(x[::2])
axes[0].set_xticklabels(list(monthly["order_year_month"])[::2], rotation=45, ha="right")

colors_bar = [GREEN if v >= 0 else RED for v in monthly["mom_growth_pct"].fillna(0)]
axes[1].bar(x, monthly["mom_growth_pct"].fillna(0), color=colors_bar, alpha=0.8)
axes[1].axhline(0, color="black", linewidth=0.8)
axes[1].set_ylabel("MoM Growth %")
axes[1].set_xlabel("Month")
axes[1].set_xticks(x[::2])
axes[1].set_xticklabels(list(monthly["order_year_month"])[::2], rotation=45, ha="right")

plt.tight_layout()
plt.savefig(f"{CHARTS}/02a_monthly_revenue_trend.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/02a_monthly_revenue_trend.png")

# ─────────────────────────────────────────────
# 3. REVENUE BY STATE
# ─────────────────────────────────────────────
print("\n[3] Revenue by state...")

state_rev = (
    fact.groupby("customer_state")
    .agg(revenue=("price","sum"), orders=("order_id","nunique"),
         customers=("customer_id","nunique"))
    .sort_values("revenue", ascending=False)
    .reset_index()
)
state_rev.to_csv(f"{CSV_OUT}/revenue_by_state.csv", index=False)

fig, ax = plt.subplots(figsize=(12, 6))
bars = ax.bar(state_rev["customer_state"], state_rev["revenue"] / 1e6,
              color=BLUE, alpha=0.8, edgecolor="white")
ax.set_title("Revenue by Customer State (R$ Millions)", fontsize=14, fontweight="bold")
ax.set_xlabel("State")
ax.set_ylabel("Revenue (R$ Millions)")
# Colour top 5 differently
for i, bar in enumerate(bars[:5]):
    bar.set_color(ORANGE)
plt.tight_layout()
plt.savefig(f"{CHARTS}/02b_revenue_by_state.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/02b_revenue_by_state.png")

# ─────────────────────────────────────────────
# 4. PAYMENT METHOD ANALYSIS
# ─────────────────────────────────────────────
print("\n[4] Payment method analysis...")

pay = fact.drop_duplicates("order_id")
pay_type = (
    pay.groupby("payment_type")
    .agg(orders=("order_id","nunique"), revenue=("payment_value","sum"))
    .sort_values("revenue", ascending=False)
    .reset_index()
)
pay_type.to_csv(f"{CSV_OUT}/payment_method.csv", index=False)

fig, axes = plt.subplots(1, 2, figsize=(12, 5))
fig.suptitle("Payment Method Analysis", fontsize=14, fontweight="bold")

colors_pie = [BLUE, GREEN, ORANGE, PURPLE, RED][:len(pay_type)]
axes[0].pie(pay_type["orders"], labels=pay_type["payment_type"],
            autopct="%1.1f%%", colors=colors_pie, startangle=90)
axes[0].set_title("Share of Orders")

axes[1].barh(pay_type["payment_type"], pay_type["revenue"] / 1e6,
             color=colors_pie, alpha=0.85)
axes[1].set_xlabel("Revenue (R$ Millions)")
axes[1].set_title("Revenue by Payment Method")

plt.tight_layout()
plt.savefig(f"{CHARTS}/02c_payment_analysis.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: charts/02c_payment_analysis.png")

# ─────────────────────────────────────────────
# 5. ORDERS BY DAY OF WEEK
# ─────────────────────────────────────────────
day_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
fact["day_name"] = fact["order_purchase_timestamp"].dt.day_name()
dow = (fact.groupby("day_name")["order_id"].nunique()
           .reindex(day_order).reset_index())
dow.columns = ["day","orders"]
dow.to_csv(f"{CSV_OUT}/orders_by_day.csv", index=False)

fig, ax = plt.subplots(figsize=(10, 5))
bars = ax.bar(dow["day"], dow["orders"], color=BLUE, alpha=0.8, edgecolor="white")
bars[dow["orders"].idxmax()].set_color(ORANGE)
ax.set_title("Orders by Day of Week", fontsize=14, fontweight="bold")
ax.set_xlabel("Day")
ax.set_ylabel("Number of Orders")
plt.tight_layout()
plt.savefig(f"{CHARTS}/02d_orders_by_dow.png", dpi=150, bbox_inches="tight")
plt.close()

print("\n  Summary:")
print(f"  Best revenue month  : {monthly.loc[monthly['revenue'].idxmax(), 'order_year_month']}")
print(f"  Best MoM growth     : {monthly['mom_growth_pct'].max():.1f}%")
print(f"  Top state           : {state_rev.iloc[0]['customer_state']}")
print(f"  Top payment type    : {pay_type.iloc[0]['payment_type']}")

print("\n✅ STEP 02 COMPLETE\n")
