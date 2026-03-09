-- ============================================================
-- QUERY 02 — Customer Metrics Table
-- Purpose : Calculate per-customer KPIs such as total orders,
--           total spend, lifetime value, and segment label.
-- ============================================================

WITH customer_base AS (
    SELECT
        customer_id,
        COUNT(DISTINCT order_id)                      AS total_orders,
        SUM(price)                                    AS total_revenue,
        SUM(freight_value)                            AS total_freight_paid,
        SUM(price - freight_value)                    AS total_profit_generated,
        ROUND(AVG(price), 2)                          AS avg_order_value,
        MIN(order_purchase_timestamp)                 AS first_order_date,
        MAX(order_purchase_timestamp)                 AS last_order_date,
        COUNT(DISTINCT order_year_month)              AS active_months
    FROM orders_fact
    GROUP BY customer_id
),

customer_with_segment AS (
    SELECT
        *,
        -- Days since last purchase (recency)
        JULIANDAY('now') - JULIANDAY(last_order_date) AS days_since_last_order,

        -- Customer lifetime (days between first and last order)
        JULIANDAY(last_order_date) - JULIANDAY(first_order_date) AS customer_lifetime_days,

        -- Segment label
        CASE
            WHEN total_orders = 1            THEN 'New'
            WHEN total_orders BETWEEN 2 AND 4 THEN 'Active'
            WHEN total_orders >= 5           THEN 'Loyal'
        END AS customer_segment
    FROM customer_base
)

SELECT * FROM customer_with_segment
ORDER BY total_revenue DESC;


-- ─────────────────────────────────────────────
-- BONUS: Repeat purchase rate
-- ─────────────────────────────────────────────
SELECT
    ROUND(
        100.0 * COUNT(CASE WHEN total_orders > 1 THEN 1 END) / COUNT(*), 2
    ) AS repeat_purchase_rate_pct
FROM customer_base;
