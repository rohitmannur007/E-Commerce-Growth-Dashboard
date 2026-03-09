-- ============================================================
-- QUERY 03 — Product Metrics Table
-- Purpose : Identify best-selling and most profitable products.
--           Also surfaces products with high revenue but low
--           profit (shipping-heavy products).
-- ============================================================

SELECT
    product_id,
    product_category_name,

    COUNT(DISTINCT order_id)                      AS total_orders,
    SUM(price)                                    AS total_revenue,
    SUM(freight_value)                            AS total_freight_cost,
    SUM(price - freight_value)                    AS total_profit,

    ROUND(AVG(price), 2)                          AS avg_selling_price,
    ROUND(AVG(freight_value), 2)                  AS avg_freight_cost,

    -- Profit margin %
    ROUND(100.0 * SUM(price - freight_value) / NULLIF(SUM(price), 0), 2)
                                                  AS profit_margin_pct,

    -- Freight as % of revenue (higher = less efficient)
    ROUND(100.0 * SUM(freight_value) / NULLIF(SUM(price), 0), 2)
                                                  AS freight_pct_of_revenue,

    ROUND(AVG(review_score), 2)                   AS avg_review_score

FROM orders_fact
GROUP BY product_id, product_category_name
ORDER BY total_revenue DESC;


-- ─────────────────────────────────────────────
-- TOP 10 PRODUCTS BY REVENUE
-- ─────────────────────────────────────────────
SELECT
    product_id,
    product_category_name,
    SUM(price)                 AS total_revenue,
    SUM(price - freight_value) AS total_profit
FROM orders_fact
GROUP BY product_id
ORDER BY total_revenue DESC
LIMIT 10;


-- ─────────────────────────────────────────────
-- WORST 10 PRODUCTS (high revenue, low profit)
-- ─────────────────────────────────────────────
SELECT
    product_id,
    product_category_name,
    SUM(price)                 AS total_revenue,
    SUM(price - freight_value) AS total_profit,
    ROUND(100.0 * SUM(price - freight_value) / NULLIF(SUM(price), 0), 2)
                               AS profit_margin_pct
FROM orders_fact
GROUP BY product_id
HAVING total_revenue > 1000   -- only products with meaningful revenue
ORDER BY profit_margin_pct ASC
LIMIT 10;
