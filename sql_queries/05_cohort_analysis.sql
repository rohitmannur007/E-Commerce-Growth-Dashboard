-- ============================================================
-- QUERY 05 — Cohort Analysis Base Query
-- Purpose : Assign each customer to their acquisition cohort
--           (month of first purchase). Then track how many
--           return in subsequent months → Retention table.
-- ============================================================

-- Step 1: Find each customer's first purchase month
WITH first_purchase AS (
    SELECT
        customer_id,
        MIN(order_year_month)           AS cohort_month
    FROM orders_fact
    GROUP BY customer_id
),

-- Step 2: Tag every order with cohort_month
orders_with_cohort AS (
    SELECT
        o.customer_id,
        o.order_year_month,
        fp.cohort_month,

        -- Month index: 0 = acquisition month, 1 = one month later, etc.
        -- Using a simplified numeric approach
        (CAST(SUBSTR(o.order_year_month, 1, 4) AS INTEGER) * 12 +
         CAST(SUBSTR(o.order_year_month, 6, 2) AS INTEGER))
        -
        (CAST(SUBSTR(fp.cohort_month, 1, 4) AS INTEGER) * 12 +
         CAST(SUBSTR(fp.cohort_month, 6, 2) AS INTEGER))  AS month_index

    FROM orders_fact         o
    JOIN first_purchase      fp ON o.customer_id = fp.customer_id
),

-- Step 3: Count distinct customers per cohort per month_index
cohort_data AS (
    SELECT
        cohort_month,
        month_index,
        COUNT(DISTINCT customer_id)     AS returning_customers
    FROM orders_with_cohort
    GROUP BY cohort_month, month_index
),

-- Step 4: Get cohort size (month 0 = total acquired)
cohort_sizes AS (
    SELECT
        cohort_month,
        returning_customers             AS cohort_size
    FROM cohort_data
    WHERE month_index = 0
)

-- Final: Retention rate = returning / cohort_size
SELECT
    cd.cohort_month,
    cd.month_index,
    cd.returning_customers,
    cs.cohort_size,
    ROUND(100.0 * cd.returning_customers / cs.cohort_size, 1) AS retention_rate_pct
FROM cohort_data     cd
JOIN cohort_sizes    cs ON cd.cohort_month = cs.cohort_month
ORDER BY cd.cohort_month, cd.month_index;


-- ─────────────────────────────────────────────
-- COHORT REVENUE (how much does each cohort generate over time)
-- ─────────────────────────────────────────────
WITH first_purchase AS (
    SELECT customer_id, MIN(order_year_month) AS cohort_month
    FROM orders_fact GROUP BY customer_id
),
orders_with_cohort AS (
    SELECT
        o.customer_id,
        o.order_year_month,
        o.price,
        fp.cohort_month,
        (CAST(SUBSTR(o.order_year_month,1,4) AS INT)*12 + CAST(SUBSTR(o.order_year_month,6,2) AS INT))
        - (CAST(SUBSTR(fp.cohort_month,1,4) AS INT)*12 + CAST(SUBSTR(fp.cohort_month,6,2) AS INT))
        AS month_index
    FROM orders_fact o
    JOIN first_purchase fp ON o.customer_id = fp.customer_id
)
SELECT
    cohort_month,
    month_index,
    ROUND(SUM(price), 2)        AS cohort_revenue
FROM orders_with_cohort
GROUP BY cohort_month, month_index
ORDER BY cohort_month, month_index;
