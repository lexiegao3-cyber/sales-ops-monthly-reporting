-- =====================================================================
-- Microsoft Access versions of the core queries (Access SQL dialect).
--
-- Setup in Access:
--   External Data > New Data Source > From File > Text File >
--   import data/processed/fact_sales_line.csv as table fact_sales_line
--   (first row = field names; set line_revenue, unit_price to Double,
--   quantity to Long Integer, invoice_date to Date/Time) and the dim_* CSVs.
--   Create each query below in SQL view and save it with the given name.
--
-- Differences from the SQLite versions in ../sales_ops_queries.sql:
--   IIf() instead of CASE, Format() for months, no window functions,
--   no COUNT(DISTINCT) -> distinct counts done with a sub-query.
-- =====================================================================


-- query name: qryMonthlySales
SELECT Format([invoice_date], "yyyy-mm")                                   AS month_label,
       Sum(IIf([line_type] = "sale", [line_revenue], 0))                   AS gross_sales,
       -Sum(IIf([line_type] = "cancellation", [line_revenue], 0))          AS cancellations,
       Sum(IIf([line_type] In ("sale", "cancellation"), [line_revenue], 0)) AS net_sales,
       Sum(IIf([line_type] = "sale", [quantity], 0))                       AS units
FROM fact_sales_line
GROUP BY Format([invoice_date], "yyyy-mm")
ORDER BY Format([invoice_date], "yyyy-mm");


-- query name: qryMonthlyOrders   (distinct invoices per month)
SELECT d.month_label, Count(*) AS orders
FROM (SELECT DISTINCT Format([invoice_date], "yyyy-mm") AS month_label, invoice
      FROM fact_sales_line
      WHERE line_type = "sale") AS d
GROUP BY d.month_label
ORDER BY d.month_label;


-- query name: qryMonthlyKPIs   (joins the two saved queries above)
SELECT s.month_label, s.gross_sales, s.cancellations, s.net_sales,
       IIf(s.gross_sales = 0, Null, s.cancellations / s.gross_sales) AS cancellation_rate,
       o.orders,
       IIf(o.orders = 0, Null, s.gross_sales / o.orders)             AS avg_order_value
FROM qryMonthlySales AS s
INNER JOIN qryMonthlyOrders AS o ON s.month_label = o.month_label
ORDER BY s.month_label;


-- query name: qryCountrySales
SELECT f.country, c.region,
       Sum(IIf(f.line_type In ("sale", "cancellation"), f.line_revenue, 0)) AS net_sales
FROM fact_sales_line AS f
INNER JOIN dim_country AS c ON f.country = c.country
GROUP BY f.country, c.region
ORDER BY 3 DESC;


-- query name: qryTopProducts   (parameter query: Access prompts for the month)
SELECT TOP 20 f.stock_code, p.description,
       Sum(IIf(f.line_type In ("sale", "cancellation"), f.line_revenue, 0)) AS net_sales,
       Sum(IIf(f.line_type = "sale", f.quantity, 0))                       AS units
FROM fact_sales_line AS f
INNER JOIN dim_product AS p ON f.stock_code = p.stock_code
WHERE Format(f.invoice_date, "yyyy-mm") = [Enter month (yyyy-mm)]
GROUP BY f.stock_code, p.description
ORDER BY 3 DESC;


-- query name: qryGuestShare
SELECT Format([invoice_date], "yyyy-mm") AS month_label,
       Sum(IIf([is_guest] = 1, [line_revenue], 0)) / Sum([line_revenue]) AS guest_share_of_sales
FROM fact_sales_line
WHERE line_type = "sale"
GROUP BY Format([invoice_date], "yyyy-mm");


-- query name: qryMakeTblMonthlySnapshot   (make-table query: run once per month-end)
SELECT k.*, Now() AS snapshot_at
INTO tblMonthlyKPISnapshot
FROM qryMonthlyKPIs AS k;
