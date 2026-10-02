-- ============================================================================
--  Customer Data Analysis  |  02 - Analysis queries
--  Project : customer-data-analysis-sql-excel
--
--  Every query below is written in a portable ANSI-ish dialect so that the
--  exact same file runs on MySQL 8 AND on SQLite (used by scripts/analysis.py
--  to build the live HTML report without requiring a MySQL server).
--
--  Each block starts with a machine-readable marker:
--      -- @query: <key> | <human readable title>
--  The report generator splits the file on those markers.
--
--  Demonstrated SQL features:
--      INNER / LEFT JOIN, GROUP BY + HAVING, subqueries (scalar, IN, correlated),
--      aggregate functions, CASE WHEN pivoting, COALESCE, window functions.
-- ============================================================================

-- @query: kpi_overview | Headline KPIs
SELECT
    (SELECT COUNT(*) FROM customers)                                  AS total_customers,
    COUNT(DISTINCT o.order_id)                                        AS total_orders,
    COUNT(DISTINCT o.customer_id)                                     AS buying_customers,
    SUM(oi.quantity)                                                  AS units_sold,
    ROUND(SUM(oi.line_total), 2)                                      AS total_revenue,
    ROUND(SUM(oi.line_total) / COUNT(DISTINCT o.order_id), 2)         AS avg_order_value
FROM orders o
JOIN order_items oi ON oi.order_id = o.order_id;


-- @query: revenue_by_segment | Revenue and AOV by customer segment
SELECT
    c.segment,
    COUNT(DISTINCT c.customer_id)                                AS customers,
    COUNT(DISTINCT o.order_id)                                   AS orders,
    ROUND(SUM(oi.line_total), 2)                                 AS revenue,
    ROUND(SUM(oi.line_total) / COUNT(DISTINCT o.order_id), 2)    AS avg_order_value,
    ROUND(AVG(oi.line_total), 2)                                 AS avg_line_value
FROM customers c
JOIN orders o       ON o.customer_id = c.customer_id
JOIN order_items oi ON oi.order_id = o.order_id
GROUP BY c.segment
ORDER BY revenue DESC;


-- @query: revenue_by_state | Revenue by state with share of total
SELECT
    c.state,
    COUNT(DISTINCT c.customer_id)                                AS customers,
    COUNT(DISTINCT o.order_id)                                   AS orders,
    ROUND(SUM(oi.line_total), 2)                                 AS revenue,
    ROUND(100.0 * SUM(oi.line_total) / (SELECT SUM(line_total) FROM order_items), 2) AS revenue_share_pct
FROM customers c
JOIN orders o       ON o.customer_id = c.customer_id
JOIN order_items oi ON oi.order_id = o.order_id
GROUP BY c.state
ORDER BY revenue DESC;


-- @query: revenue_by_category | Product category performance
SELECT
    p.category,
    COUNT(DISTINCT oi.order_id)                       AS orders,
    SUM(oi.quantity)                                  AS units_sold,
    ROUND(SUM(oi.line_total), 2)                      AS revenue,
    ROUND(AVG(oi.discount_pct) * 100, 2)              AS avg_discount_pct,
    ROUND(SUM(oi.line_total) / SUM(oi.quantity), 2)   AS revenue_per_unit
FROM order_items oi
JOIN products p ON p.product_id = oi.product_id
GROUP BY p.category
ORDER BY revenue DESC;


-- @query: monthly_revenue_trend | Monthly revenue and cumulative revenue
SELECT
    o.order_month                                            AS month,
    COUNT(DISTINCT o.order_id)                               AS orders,
    ROUND(SUM(oi.line_total), 2)                             AS revenue,
    ROUND(SUM(SUM(oi.line_total)) OVER (ORDER BY o.order_month), 2) AS running_total
FROM orders o
JOIN order_items oi ON oi.order_id = o.order_id
GROUP BY o.order_month
ORDER BY o.order_month;


-- @query: channel_year_crosstab | Revenue by channel and year (CASE WHEN pivot)
SELECT
    o.channel,
    ROUND(SUM(CASE WHEN o.order_year = 2023 THEN oi.line_total ELSE 0 END), 2) AS y2023,
    ROUND(SUM(CASE WHEN o.order_year = 2024 THEN oi.line_total ELSE 0 END), 2) AS y2024,
    ROUND(SUM(CASE WHEN o.order_year = 2025 THEN oi.line_total ELSE 0 END), 2) AS y2025,
    ROUND(SUM(oi.line_total), 2)                                              AS total_revenue
FROM orders o
JOIN order_items oi ON oi.order_id = o.order_id
GROUP BY o.channel
ORDER BY total_revenue DESC;


-- @query: top_10_customers | Top 10 customers by lifetime revenue
--   NOTE: first/last name are kept as separate columns so that the query stays
--   portable (MySQL's `||` operator is OR unless PIPES_AS_CONCAT is enabled,
--   so CONCAT()/|| string building is deliberately avoided here).
SELECT
    c.customer_id,
    c.first_name,
    c.last_name,
    c.city,
    c.state,
    c.segment,
    c.loyalty_tier,
    COUNT(DISTINCT o.order_id)                  AS orders,
    ROUND(SUM(oi.line_total), 2)                AS lifetime_revenue
FROM customers c
JOIN orders o       ON o.customer_id = c.customer_id
JOIN order_items oi ON oi.order_id = o.order_id
GROUP BY c.customer_id, c.first_name, c.last_name, c.city, c.state, c.segment, c.loyalty_tier
ORDER BY lifetime_revenue DESC
LIMIT 10;


-- @query: customer_value_tiers | Bucket customers into value tiers
SELECT
    CASE
        WHEN revenue >= 200000 THEN 'Platinum (2L+)'
        WHEN revenue >= 100000 THEN 'Gold (1L-2L)'
        WHEN revenue >= 50000  THEN 'Silver (50k-1L)'
        WHEN revenue >= 20000  THEN 'Bronze (20k-50k)'
        ELSE 'Low (<20k)'
    END                                          AS value_tier,
    COUNT(*)                                     AS customers,
    ROUND(SUM(revenue), 2)                       AS total_revenue,
    ROUND(AVG(revenue), 2)                       AS avg_customer_value
FROM (
    SELECT o.customer_id, SUM(oi.line_total) AS revenue
    FROM orders o
    JOIN order_items oi ON oi.order_id = o.order_id
    GROUP BY o.customer_id
) AS per_customer
GROUP BY value_tier
ORDER BY total_revenue DESC;


-- @query: high_value_customers | Customers above the 1,50,000 revenue threshold (HAVING)
SELECT
    c.customer_id,
    c.segment,
    c.loyalty_tier,
    COUNT(DISTINCT o.order_id)      AS orders,
    ROUND(SUM(oi.line_total), 2)    AS revenue
FROM customers c
JOIN orders o       ON o.customer_id = c.customer_id
JOIN order_items oi ON oi.order_id = o.order_id
GROUP BY c.customer_id, c.segment, c.loyalty_tier
HAVING SUM(oi.line_total) > 150000
ORDER BY revenue DESC;


-- @query: above_average_customers | Customers spending above the average (non-correlated subquery)
SELECT
    c.customer_id,
    c.segment,
    c.loyalty_tier,
    ROUND(SUM(oi.line_total), 2) AS revenue,
    ROUND((SELECT AVG(line_total) * 20 FROM order_items), 2) AS avg_customer_reference
FROM customers c
JOIN orders o       ON o.customer_id = c.customer_id
JOIN order_items oi ON oi.order_id = o.order_id
WHERE o.customer_id IN (
        SELECT o2.customer_id
        FROM orders o2
        JOIN order_items oi2 ON oi2.order_id = o2.order_id
        GROUP BY o2.customer_id
        HAVING SUM(oi2.line_total) > (
            SELECT AVG(customer_total)
            FROM (
                SELECT SUM(oi3.line_total) AS customer_total
                FROM orders o3
                JOIN order_items oi3 ON oi3.order_id = o3.order_id
                GROUP BY o3.customer_id
            ) AS totals
        )
)
GROUP BY c.customer_id, c.segment, c.loyalty_tier
ORDER BY revenue DESC
LIMIT 15;


-- @query: customers_without_orders | Registered but never purchased (LEFT JOIN ... IS NULL)
SELECT
    c.customer_id,
    c.city,
    c.state,
    c.segment,
    c.signup_date,
    c.loyalty_tier
FROM customers c
LEFT JOIN orders o ON o.customer_id = c.customer_id
WHERE o.order_id IS NULL
ORDER BY c.signup_date
LIMIT 20;


-- @query: repeat_vs_onetime | Repeat buyers vs one-time buyers
SELECT
    CASE WHEN order_count = 1 THEN 'One-time buyer'
         WHEN order_count BETWEEN 2 AND 4 THEN 'Occasional (2-4)'
         ELSE 'Loyal (5+)'
    END                             AS buyer_type,
    COUNT(*)                        AS customers,
    ROUND(AVG(revenue), 2)          AS avg_revenue,
    ROUND(SUM(revenue), 2)          AS total_revenue
FROM (
    SELECT o.customer_id,
           COUNT(DISTINCT o.order_id)     AS order_count,
           SUM(oi.line_total)             AS revenue
    FROM orders o
    JOIN order_items oi ON oi.order_id = o.order_id
    GROUP BY o.customer_id
) AS customer_orders
GROUP BY buyer_type
ORDER BY total_revenue DESC;


-- @query: status_by_channel | Order status mix per sales channel
SELECT
    o.channel,
    SUM(CASE WHEN o.status = 'Delivered' THEN 1 ELSE 0 END) AS delivered,
    SUM(CASE WHEN o.status = 'Shipped'   THEN 1 ELSE 0 END) AS shipped,
    SUM(CASE WHEN o.status = 'Pending'   THEN 1 ELSE 0 END) AS pending,
    SUM(CASE WHEN o.status = 'Cancelled' THEN 1 ELSE 0 END) AS cancelled,
    SUM(CASE WHEN o.status = 'Returned'  THEN 1 ELSE 0 END) AS returned,
    COUNT(*)                                                AS total_orders
FROM orders o
GROUP BY o.channel
ORDER BY total_orders DESC;


-- @query: top_customer_per_state | Highest spending customer in each state (window function)
SELECT state, customer_id, revenue, state_rank
FROM (
    SELECT
        c.state,
        c.customer_id,
        SUM(oi.line_total)                                              AS revenue,
        ROW_NUMBER() OVER (PARTITION BY c.state ORDER BY SUM(oi.line_total) DESC) AS state_rank
    FROM customers c
    JOIN orders o       ON o.customer_id = c.customer_id
    JOIN order_items oi ON oi.order_id = o.order_id
    GROUP BY c.state, c.customer_id
) AS ranked
WHERE state_rank <= 3
ORDER BY state, state_rank;


-- @query: yoy_by_channel | Year-over-year growth per channel
SELECT
    channel,
    ROUND(SUM(CASE WHEN order_year = 2023 THEN revenue ELSE 0 END), 2) AS y2023,
    ROUND(SUM(CASE WHEN order_year = 2024 THEN revenue ELSE 0 END), 2) AS y2024,
    ROUND(SUM(CASE WHEN order_year = 2025 THEN revenue ELSE 0 END), 2) AS y2025,
    ROUND(
        100.0 * (SUM(CASE WHEN order_year = 2025 THEN revenue ELSE 0 END)
                 - SUM(CASE WHEN order_year = 2024 THEN revenue ELSE 0 END))
        / NULLIF(SUM(CASE WHEN order_year = 2024 THEN revenue ELSE 0 END), 0), 2) AS yoy_growth_pct
FROM (
    SELECT o.channel, o.order_year, SUM(oi.line_total) AS revenue
    FROM orders o
    JOIN order_items oi ON oi.order_id = o.order_id
    GROUP BY o.channel, o.order_year
) AS channel_year
GROUP BY channel
ORDER BY y2025 DESC;
