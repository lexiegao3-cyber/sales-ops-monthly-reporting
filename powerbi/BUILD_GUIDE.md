# Power BI build guide

Run `python run_all.py` first (creates `data/processed/`). Then build the
`.pbix` in Power BI Desktop, about 3–4 hours the first time.

## 1. Load
1. **Transform data → Manage Parameters → New**: `ProjectFolder` (Text), the full path to this folder ending in `\`.
2. Paste each block of `power_query/queries.pq` into a Blank Query (Advanced Editor); rename to the comment name.
3. **Close & Apply.** Expected rows: `fact_sales_line` 1,033,036 · `dim_product` 4,907 ·
   `dim_customer` 5,852 · `dim_country` 43.

## 2. Model
1. **New table** → `Date` from `measures.dax`; mark as date table; sort `Month` by `Month Num`.
2. Relationships (one-to-many, single direction):
   `Date[Date]→fact_sales_line[invoice_date]`, `dim_product[stock_code]→fact[stock_code]`,
   `dim_customer[customer_id]→fact[customer_id]`, `dim_country[country]→fact[country]`.
3. Hide key and raw numeric columns from report view; create `_Measures` and add every measure.
4. **Check numbers before visuals** (no filters): Net Sales **£18,927,421**, Gross Sales **£19,643,847**,
   Orders **39,516**, Guest Share of Sales **13.1%**. If they differ, check `line_type` and relationships.

## 3. Pages
| Page | Content |
|---|---|
| **Monthly summary** | Slicer `Date[Year-Month]` (single select) · cards: Net Sales, MoM %, YoY %, Orders, AOV, Cancellation Rate, Active / New Customers · card bound to `[Partial Period Warning]` · column: Net Sales by Year-Month (last 13 months) · bar: Net Sales by region |
| **Products** | Table: description, Net Sales, Units, Cancellation Rate, Product Rank · top-N filter (20) · scatter: Units vs Net Sales by product |
| **Customers** | Matrix: `dim_customer[cohort_month]` × months since first purchase (add a calculated column) with % active · cards: New vs Returning · line: Guest Share of Sales by month |
| **Countries** | Filled map or bar by country · table with Net Sales, Cancellation Rate, Active Customers |
| **Method & data quality** | `data_quality_log` table · text box with definitions (gross / net / guest / partial month) and the data citation |

Report title pattern: `="Sales operations - " & [Selected Month Label]` (use a card or button with fx text).

## 4. Row-level security
1. **Modeling → Manage roles**: create `UK`, `EU`, `Rest of World` with the filters listed at the end of `measures.dax`.
2. **Modeling → View as → EU** and confirm Net Sales falls to the EU total from `reports/tables/country_summary.csv`.
3. After publishing: dataset → **Security** → add people or groups to each role.

## 5. Publish & administer
- Publish to a workspace (e.g. `Sales Ops – PROD`); give stakeholders Viewer, analysts Member.
- Distribute through a workspace **App** so viewers see only the finished pages.
- Scheduled refresh: local files need a gateway; storing the folder in OneDrive/SharePoint avoids it.
  Set refresh failure e-mails to yourself.
- Large file: the fact table is ~1M rows. Keep only the columns the queries keep, avoid
  calculated columns on the fact table, and check **Performance analyzer** for slow visuals.
- Write each problem you hit and its fix into `docs/troubleshooting.md` - it is good interview material.
