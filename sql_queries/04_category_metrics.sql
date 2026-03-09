-- ============================================================
-- QUERY 04 — Category Metrics Table
-- Purpose : Compare revenue vs profit across product categories.
--           Helps identify which categories look good on revenue
--           but are actually poor on profit (freight eats margin).
-- ============================================================

SELECT
    product_category_name,

    COUNT(DISTINCT order_id)                           AS total_orders,
    COUNT(DISTINCT customer_id)                        AS unique_customers,
    COUNT(DISTINCT product_id)                         AS unique_products,

    ROUND(SUM(price), 2)                               AS total_revenue,
    ROUND(SUM(freight_value), 2)                       AS total_freight_cost,
    ROUND(SUM(price - freight_value), 2)               AS total_profit,

    ROUND(AVG(price), 2)                               AS avg_item_price,
    ROUND(AVG(freight_value), 2)                       AS avg_freight,

    -- Profit margin
    ROUND(100.0 * SUM(price - freight_value) / NULLIF(SUM(price), 0), 2)
                                                       AS profit_margin_pct,

    -- Revenue share of total
    ROUND(100.0 * SUM(price) /
          SUM(SUM(price)) OVER (), 2)                  AS revenue_share_pct,

    ROUND(AVG(review_score), 2)                        AS avg_review_score

FROM orders_fact
GROUP BY product_category_name
ORDER BY total_revenue DESC;


-- ─────────────────────────────────────────────
-- MONTHLY REVENUE BY CATEGORY (trend analysis)
-- ─────────────────────────────────────────────
SELECT
    order_year_month,
    product_category_name,
    ROUND(SUM(price), 2)                           AS monthly_revenue,
    ROUND(SUM(price - freight_value), 2)           AS monthly_profit
FROM orders_fact
GROUP BY order_year_month, product_category_name
ORDER BY order_year_month, monthly_revenue DESC;
