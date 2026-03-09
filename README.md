# 📊 E-Commerce Growth & Profit Optimization Analytics

> **End-to-end analytics project** analyzing 100,000+ e-commerce orders to uncover drivers of revenue, customer retention, and category profitability — built using SQL, Python, and Excel.

---

## 🖼️ Dashboard Preview

### KPI Overview
![KPI Dashboard](https://raw.githubusercontent.com/rohitmannur007/E-Commerce-Growth-Dashboard/main/KPI.png)

### Month-over-Month Revenue Growth
![MoM Growth](https://raw.githubusercontent.com/rohitmannur007/E-Commerce-Growth-Dashboard/main/mom.png)

### Cohort Retention Analysis
![Cohort Retention](https://raw.githubusercontent.com/rohitmannur007/E-Commerce-Growth-Dashboard/main/Cohort%20.png)

---

## 🔑 Key Findings

| Metric | Result |
|--------|--------|
| 💰 Total Revenue | R$ 71.7M over 2 years |
| 📈 Profit Margin | **87.5%** avg (freight = 12.5% of revenue) |
| 🏆 Top Category | Electronics — 42% of total revenue |
| 👥 Pareto Rule | Top 61.6% of customers drive 80% of revenue |
| 🔄 Cohort Retention | 32.6% of customers return in Month 1 |
| ⭐ Top RFM Segment | Champions (1,655 customers) — avg R$ 11,646 spend |
| 📦 Price-Freight Correlation | R² = 0.824 (strong positive) |
| ⚠️ Worst Margin | Stationery — 83.2% margin due to 16.8% freight ratio |

---

## 🛠️ Tech Stack

| Tool | Usage |
|------|-------|
| **SQL** | Data extraction, joins, aggregations, CTEs, window functions |
| **Python** (pandas, numpy, matplotlib, seaborn, scipy) | Analysis & visualization |
| **Excel** (openpyxl) | Scenario simulation & financial modeling |
| **Power BI** | Executive dashboards (built separately) |

---

## 📁 Dataset

Synthetic Olist-style Brazilian e-commerce dataset:

- **100,000** orders &nbsp;|&nbsp; **10,000** customers &nbsp;|&nbsp; **3,000** products
- **15** categories &nbsp;|&nbsp; **20** states
- Date range: **Jan 2022 – Dec 2023**

---

## 🗂️ Project Structure

```
ecommerce-growth-analytics/
├── data/
│   ├── raw/                  ← Generated CSVs (customers, orders, products, etc.)
│   ├── processed/            ← Cleaned orders_fact.csv
│   └── generate_data.py      ← Synthetic data generator
├── sql_queries/
│   ├── 01_create_orders_fact.sql
│   ├── 02_customer_metrics.sql
│   ├── 03_product_metrics.sql
│   ├── 04_category_metrics.sql
│   ├── 05_cohort_analysis.sql
│   └── 06_rfm_segmentation.sql
├── python_analysis/
│   ├── 01_data_cleaning.py
│   ├── 02_business_metrics.py
│   ├── 03_customer_segmentation.py
│   ├── 04_cohort_analysis.py
│   ├── 05_product_profitability.py
│   └── 06_statistical_analysis.py
├── excel_simulation/
│   ├── build_excel_model.py
│   └── scenario_simulation.xlsx
├── outputs/
│   ├── charts/               ← 12 generated PNG charts
│   └── csv_exports/          ← 15 exported CSV analysis files
├── run_all.py                ← Master pipeline script
└── README.md
```

---

## 🚀 How to Run

```bash
# 1. Install dependencies
pip install pandas numpy matplotlib seaborn scipy openpyxl

# 2. Run full pipeline (recommended)
python3 run_all.py

# 3. Or run individual steps
python3 data/generate_data.py
python3 python_analysis/01_data_cleaning.py
python3 python_analysis/02_business_metrics.py
python3 python_analysis/03_customer_segmentation.py
python3 python_analysis/04_cohort_analysis.py
python3 python_analysis/05_product_profitability.py
python3 python_analysis/06_statistical_analysis.py
python3 excel_simulation/build_excel_model.py
```

---

## 📊 Power BI Dashboard Notes

Two dashboards built in Power BI:

1. **Executive Overview** — KPIs, monthly trend, revenue by category/region
2. **Customer & Profit Analytics** — RFM segments, cohort heatmap, CLV

---

## 📝 Resume Bullet

> Built end-to-end **E-Commerce Growth & Profit Analytics** system using SQL, Python, and Excel to analyze 100k+ orders, uncovering key revenue drivers, customer retention patterns, and category profitability across 15 product categories and 2 fiscal years.

---

## 👤 Author

**Rohit Mannur** &nbsp;·&nbsp; [GitHub](https://github.com/rohitmannur007/E-Commerce-Growth-Dashboard)
