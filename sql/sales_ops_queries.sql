-- =====================================================================
-- Sales operations KPI queries (SQLite). Run by src/report.py; each result
-- is exported to reports/tables/<name>.csv and feeds the Excel monthly report.
--
-- Revenue definitions used everywhere:
--   gross_sales   = SUM(line_revenue) of line_type 'sale'
--   cancellations = -SUM(line_revenue) of line_type 'cancellation'  (shown positive)
--   net_sales     = gross_sales - cancellations
-- Fees, postage, write-offs and bad-debt adjustments are excluded from product revenue.
-- Guest lines (no Customer ID) are INCLUDED in revenue, excluded from customer counts.
-- =====================================================================

-- name: monthly_kpis
WITH m AS (
    SELECT month,
           SUM(CASE WHEN line_type = 'sale' THEN line_revenue END)                     AS gross_sales,
           -SUM(CASE WHEN line_type = 'cancellation' THEN line_revenue ELSE 0 END)     AS cancellations,
           COUNT(DISTINCT CASE WHEN line_type = 'sale' THEN invoice END)               AS orders,
           SUM(CASE WHEN line_type = 'sale' THEN quantity END)                         AS units,
           SUM(CASE WHEN line_type = 'sale' AND is_guest = 1 THEN line_revenue END)    AS guest_sales,
           COUNT(DISTINCT CASE WHEN line_type = 'sale' THEN customer_id END)           AS active_customers,
           MAX(invoice_date)                                                           AS last_day_with_data
    FROM fact_sales_line
    GROUP BY month
),
newc AS (
    SELECT cohort_month AS month, COUNT(*) AS new_customers
    FROM dim_customer GROUP BY cohort_month
)
SELECT m.month,
       ROUND(m.gross_sales, 2)                                   AS gross_sales,
       ROUND(m.cancellations, 2)                                 AS cancellations,
       ROUND(m.gross_sales - m.cancellations, 2)                 AS net_sales,
       ROUND(m.cancellations / m.gross_sales, 4)                 AS cancellation_rate,
       m.orders,
       ROUND(m.gross_sales / m.orders, 2)                        AS avg_order_value,
       m.units,
       ROUND(m.guest_sales / m.gross_sales, 4)                   AS guest_share_of_sales,
       m.active_customers,
       COALESCE(n.new_customers, 0)                              AS new_customers,
       m.active_customers - COALESCE(n.new_customers, 0)         AS returning_customers,
       CASE WHEN m.month = (SELECT MAX(month) FROM m)
             AND CAST(strftime('%d', m.last_day_with_data) AS INT) < 25
            THEN 1 ELSE 0 END                                    AS is_partial_month,   -- data ends mid-month
       ROUND(1.0 * (m.gross_sales - m.cancellations)
             / LAG(m.gross_sales - m.cancellations) OVER (ORDER BY m.month) - 1, 4) AS net_sales_mom,
       ROUND(1.0 * (m.gross_sales - m.cancellations)
             / LAG(m.gross_sales - m.cancellations, 12) OVER (ORDER BY m.month) - 1, 4) AS net_sales_yoy
FROM m LEFT JOIN newc n ON n.month = m.month
ORDER BY m.month;

-- name: monthly_region_sales
SELECT f.month, c.region,
       ROUND(SUM(CASE WHEN f.line_type = 'sale' THEN f.line_revenue ELSE 0 END), 2)          AS gross_sales,
       ROUND(SUM(CASE WHEN f.line_type IN ('sale','cancellation') THEN f.line_revenue ELSE 0 END), 2) AS net_sales,
       COUNT(DISTINCT CASE WHEN f.line_type = 'sale' THEN f.invoice END)                     AS orders
FROM fact_sales_line f JOIN dim_country c ON c.country = f.country
GROUP BY f.month, c.region
ORDER BY f.month, c.region;

-- name: country_summary
SELECT f.country, c.region,
       ROUND(SUM(CASE WHEN f.line_type = 'sale' THEN f.line_revenue ELSE 0 END), 2)          AS gross_sales,
       ROUND(SUM(CASE WHEN f.line_type IN ('sale','cancellation') THEN f.line_revenue ELSE 0 END), 2) AS net_sales,
       ROUND(-SUM(CASE WHEN f.line_type = 'cancellation' THEN f.line_revenue ELSE 0 END)
             / SUM(CASE WHEN f.line_type = 'sale' THEN f.line_revenue ELSE 0 END), 4)        AS cancellation_rate,
       COUNT(DISTINCT CASE WHEN f.line_type = 'sale' THEN f.customer_id END)                 AS customers
FROM fact_sales_line f JOIN dim_country c ON c.country = f.country
GROUP BY f.country, c.region
ORDER BY net_sales DESC;

-- name: top_products
WITH p AS (
    SELECT stock_code,
           SUM(CASE WHEN line_type = 'sale' THEN line_revenue ELSE 0 END)          AS gross_sales,
           -SUM(CASE WHEN line_type = 'cancellation' THEN line_revenue ELSE 0 END) AS cancellations,
           SUM(CASE WHEN line_type = 'sale' THEN quantity ELSE 0 END)              AS units,
           COUNT(DISTINCT CASE WHEN line_type = 'sale' THEN invoice END)           AS orders
    FROM fact_sales_line
    WHERE line_type IN ('sale', 'cancellation')
    GROUP BY stock_code
)
SELECT p.stock_code, d.description,
       ROUND(p.gross_sales, 2)                            AS gross_sales,
       ROUND(p.gross_sales - p.cancellations, 2)          AS net_sales,
       ROUND(p.cancellations / NULLIF(p.gross_sales, 0), 4) AS cancellation_rate,
       p.units, p.orders,
       RANK() OVER (ORDER BY p.gross_sales - p.cancellations DESC) AS net_sales_rank
FROM p JOIN dim_product d ON d.stock_code = p.stock_code
ORDER BY net_sales DESC
LIMIT 50;

-- name: high_cancellation_products
-- Products with meaningful volume (>= 100 orders) and the highest share of value cancelled.
WITH p AS (
    SELECT stock_code,
           SUM(CASE WHEN line_type = 'sale' THEN line_revenue ELSE 0 END)          AS gross_sales,
           -SUM(CASE WHEN line_type = 'cancellation' THEN line_revenue ELSE 0 END) AS cancellations,
           COUNT(DISTINCT CASE WHEN line_type = 'sale' THEN invoice END)           AS orders
    FROM fact_sales_line
    WHERE line_type IN ('sale', 'cancellation')
    GROUP BY stock_code
)
SELECT p.stock_code, d.description, p.orders,
       ROUND(p.gross_sales, 2) AS gross_sales, ROUND(p.cancellations, 2) AS cancellations,
       ROUND(p.cancellations / p.gross_sales, 4) AS cancellation_rate
FROM p JOIN dim_product d ON d.stock_code = p.stock_code
WHERE p.orders >= 100 AND p.gross_sales > 0
ORDER BY cancellation_rate DESC
LIMIT 25;

-- name: customer_pareto
-- How concentrated is revenue among identified customers?
WITH c AS (
    SELECT customer_id,
           SUM(CASE WHEN line_type IN ('sale','cancellation') THEN line_revenue ELSE 0 END) AS net_sales
    FROM fact_sales_line WHERE is_guest = 0
    GROUP BY customer_id
),
r AS (
    SELECT net_sales,
           ROW_NUMBER() OVER (ORDER BY net_sales DESC)                            AS rnk,
           SUM(net_sales) OVER (ORDER BY net_sales DESC ROWS UNBOUNDED PRECEDING) AS cum_sales
    FROM c
),
t(pct) AS (VALUES (1), (5), (10), (20), (50)),
k AS (
    SELECT pct, CAST((SELECT COUNT(*) FROM c) * pct / 100.0 AS INT) AS cutoff_rank FROM t
)
SELECT k.pct                                                        AS top_pct_of_customers,
       k.cutoff_rank                                                AS customers_in_group,
       ROUND(r.cum_sales / (SELECT SUM(net_sales) FROM c), 4)       AS share_of_customer_net_sales
FROM k JOIN r ON r.rnk = k.cutoff_rank
ORDER BY k.pct;

-- name: largest_single_lines
-- Very large single lines distort monthly totals; most are cancelled again.
SELECT f.invoice, f.invoice_date, f.stock_code, d.description, f.quantity,
       ROUND(f.line_revenue, 2) AS line_revenue, f.line_type, f.customer_id, f.country
FROM fact_sales_line f LEFT JOIN dim_product d ON d.stock_code = f.stock_code
WHERE f.line_type IN ('sale', 'cancellation') AND ABS(f.quantity) >= 10000
ORDER BY ABS(f.quantity) DESC;

-- name: cohort_retention
-- % of each first-purchase cohort that bought again N months later.
WITH activity AS (
    SELECT DISTINCT f.customer_id, f.month
    FROM fact_sales_line f
    WHERE f.line_type = 'sale' AND f.is_guest = 0
),
x AS (
    SELECT c.cohort_month, a.customer_id,
           (CAST(strftime('%Y', a.month) AS INT) - CAST(strftime('%Y', c.cohort_month) AS INT)) * 12
         + (CAST(strftime('%m', a.month) AS INT) - CAST(strftime('%m', c.cohort_month) AS INT)) AS months_since
    FROM activity a JOIN dim_customer c ON c.customer_id = a.customer_id
)
SELECT cohort_month, months_since,
       COUNT(DISTINCT customer_id) AS customers,
       ROUND(1.0 * COUNT(DISTINCT customer_id)
             / FIRST_VALUE(COUNT(DISTINCT customer_id)) OVER (PARTITION BY cohort_month ORDER BY months_since), 4)
             AS retention_rate
FROM x
GROUP BY cohort_month, months_since
ORDER BY cohort_month, months_since;

-- name: line_type_summary
SELECT line_type,
       COUNT(*) AS lines,
       ROUND(SUM(line_revenue), 2) AS revenue_value,
       COUNT(DISTINCT invoice) AS invoices
FROM fact_sales_line
GROUP BY line_type
ORDER BY lines DESC;
