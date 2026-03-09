"""
RUN ALL — Master Execution Script
===================================
Runs all project steps in order.
Usage: python3 run_all.py
"""

import subprocess
import sys
import os
import time

BASE = os.path.dirname(os.path.abspath(__file__))

steps = [
    ("Data Generation",       "data/generate_data.py"),
    ("Data Cleaning",         "python_analysis/01_data_cleaning.py"),
    ("Business Metrics",      "python_analysis/02_business_metrics.py"),
    ("Customer Segmentation", "python_analysis/03_customer_segmentation.py"),
    ("Cohort Analysis",       "python_analysis/04_cohort_analysis.py"),
    ("Product Profitability", "python_analysis/05_product_profitability.py"),
    ("Statistical Analysis",  "python_analysis/06_statistical_analysis.py"),
    ("Excel Simulation",      "excel_simulation/build_excel_model.py"),
]

print("=" * 65)
print(" E-COMMERCE GROWTH & PROFIT ANALYTICS — FULL PIPELINE")
print("=" * 65)

total_start = time.time()
errors = []

for step_name, script_path in steps:
    full_path = os.path.join(BASE, script_path)
    print(f"\n▶  {step_name}")
    print(f"   Running: {script_path}")
    t0 = time.time()
    result = subprocess.run(
        [sys.executable, full_path],
        capture_output=True, text=True
    )
    elapsed = time.time() - t0
    if result.returncode == 0:
        print(f"   ✅ Done in {elapsed:.1f}s")
    else:
        print(f"   ❌ ERROR in {elapsed:.1f}s")
        print(result.stderr[-500:])
        errors.append(step_name)

total_elapsed = time.time() - total_start
print("\n" + "=" * 65)
if not errors:
    print(f"✅ ALL STEPS COMPLETE in {total_elapsed:.1f}s")
else:
    print(f"⚠  Completed with errors in {total_elapsed:.1f}s: {errors}")
print("=" * 65)

print("\nOutputs generated:")
for root, dirs, files in os.walk(os.path.join(BASE, "outputs")):
    for f in sorted(files):
        rel = os.path.relpath(os.path.join(root, f), BASE)
        print(f"  {rel}")
for f in os.listdir(os.path.join(BASE, "excel_simulation")):
    if f.endswith(".xlsx"):
        print(f"  excel_simulation/{f}")
