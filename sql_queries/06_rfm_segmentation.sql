-- ============================================================
-- QUERY 06 — RFM Customer Segmentation
-- Purpose : Score each customer on Recency, Frequency, and
--           Monetary value. Combine to get an RFM segment.
--           This is standard in e-commerce analytics.
-- ============================================================

-- Using the analysis reference date = max order date in data
WITH reference_date AS (
    SELECT MAX(order_purchase_timestamp) AS ref_date
    FROM orders_fact
),

rfm_raw AS (
    SELECT
        customer_id,
        -- Recency: days since last order (lower = better)
        JULIANDAY((SELECT ref_date FROM reference_date))
            - JULIANDAY(MAX(order_purchase_timestamp))           AS recency_days,

        -- Frequency: number of orders
        COUNT(DISTINCT order_id)                                 AS frequency,

        -- Monetary: total revenue generated
        ROUND(SUM(price), 2)                                     AS monetary
    FROM orders_fact
    GROUP BY customer_id
),

rfm_scored AS (
    SELECT
        customer_id,
        recency_days,
        frequency,
        monetary,

        -- Score Recency: 5 = most recent, 1 = oldest
        NTILE(5) OVER (ORDER BY recency_days ASC)  AS r_score,

        -- Score Frequency: 5 = most frequent, 1 = least
        NTILE(5) OVER (ORDER BY frequency DESC)    AS f_score,

        -- Score Monetary: 5 = highest spender, 1 = lowest
        NTILE(5) OVER (ORDER BY monetary DESC)     AS m_score

    FROM rfm_raw
),

rfm_combined AS (
    SELECT
        *,
        CAST(r_score AS TEXT) || CAST(f_score AS TEXT) || CAST(m_score AS TEXT)
                                                        AS rfm_score,
        (r_score + f_score + m_score)                   AS rfm_total

    FROM rfm_scored
),

rfm_segment AS (
    SELECT
        *,
        CASE
            WHEN rfm_total >= 13                         THEN 'Champions'
            WHEN rfm_total >= 10 AND r_score >= 3        THEN 'Loyal Customers'
            WHEN rfm_total >= 9  AND r_score >= 4        THEN 'Potential Loyalists'
            WHEN r_score = 5     AND rfm_total < 8       THEN 'New Customers'
            WHEN r_score >= 3    AND f_score <= 2
                                 AND m_score <= 2        THEN 'Promising'
            WHEN r_score <= 2    AND rfm_total >= 9      THEN 'At Risk'
            WHEN r_score <= 2    AND rfm_total >= 7      THEN 'Cannot Lose Them'
            WHEN r_score <= 2    AND rfm_total < 7       THEN 'Lost'
            ELSE 'Need Attention'
        END AS rfm_segment
    FROM rfm_combined
)

SELECT * FROM rfm_segment
ORDER BY rfm_total DESC;


-- ─────────────────────────────────────────────
-- SEGMENT SUMMARY (for dashboards)
-- ─────────────────────────────────────────────
SELECT
    rfm_segment,
    COUNT(*)                    AS customer_count,
    ROUND(AVG(monetary), 2)     AS avg_monetary,
    ROUND(AVG(frequency), 2)    AS avg_frequency,
    ROUND(AVG(recency_days), 0) AS avg_recency_days
FROM (
    -- paste rfm_segment CTE result here
    SELECT rfm_segment, monetary, frequency, recency_days
    FROM rfm_segment
)
GROUP BY rfm_segment
ORDER BY avg_monetary DESC;
