-- ============================================================
-- QUERY 01 — Orders Fact Table
-- Purpose : Build the master analytics table by joining all
--           source tables. This is the foundation for every
--           downstream metric and analysis.
-- ============================================================

SELECT
    o.order_id,
    o.customer_id,
    o.order_status,
    o.order_purchase_timestamp,
    o.order_delivered_customer_date,

    oi.order_item_id,
    oi.product_id,
    oi.price,
    oi.freight_value,

    -- Derived profit estimate
    ROUND(oi.price - oi.freight_value, 2)           AS profit_estimate,

    -- Date parts (useful for time-series slicing)
    STRFTIME('%Y',  o.order_purchase_timestamp)      AS order_year,
    STRFTIME('%m',  o.order_purchase_timestamp)      AS order_month,
    STRFTIME('%Y-%m', o.order_purchase_timestamp)    AS order_year_month,

    p.product_category_name,
    p.product_weight_g,

    py.payment_type,
    py.payment_installments,
    py.payment_value,

    c.customer_city,
    c.customer_state,

    COALESCE(r.review_score, 0)                      AS review_score

FROM orders           o
JOIN order_items      oi ON o.order_id   = oi.order_id
JOIN products         p  ON oi.product_id = p.product_id
JOIN customers        c  ON o.customer_id = c.customer_id
LEFT JOIN payments    py ON o.order_id   = py.order_id
LEFT JOIN reviews     r  ON o.order_id   = r.order_id

WHERE o.order_status != 'cancelled';   -- Exclude cancelled orders from analytics
