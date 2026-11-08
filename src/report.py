"""
Build the monthly sales-operations report from the SQLite warehouse:
  reports/tables/*.csv                 every KPI query result
  reports/figures/*.png                charts used in the README
  excel/sales_ops_monthly_report.xlsx  month-selectable Excel report (formulas, tables, validation)
  reports/findings.md                  narrative with every number computed here
"""
from __future__ import annotations

import re
import sqlite3
import time
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "warehouse" / "retail.sqlite"
SQL = ROOT / "sql" / "sales_ops_queries.sql"
TABLES = ROOT / "reports" / "tables"
FIGS = ROOT / "reports" / "figures"
XLSX = ROOT / "excel" / "sales_ops_monthly_report.xlsx"

BLUE, GREY, INK, INK2, GRID, SURFACE = "#2a78d6", "#b7b6b0", "#0b0b0b", "#52514e", "#e6e5e1", "#fcfcfb"


def run_sql(con) -> dict[str, pd.DataFrame]:
    out = {}
    for block in re.split(r"^-- name: ", SQL.read_text(), flags=re.M)[1:]:
        name, _, q = block.partition("\n")
        out[name.strip()] = pd.read_sql(q, con)
    return out


def same_day_reversals(con) -> pd.DataFrame:
    """Sale lines fully cancelled the same day by the same customer (order entered then voided)."""
    q = """
    SELECT s.month, s.invoice, s.stock_code, s.quantity, s.line_revenue
    FROM fact_sales_line s
    JOIN fact_sales_line c
      ON c.line_type = 'cancellation' AND c.customer_id = s.customer_id
     AND c.stock_code = s.stock_code AND c.quantity = -s.quantity AND c.invoice_date = s.invoice_date
    WHERE s.line_type = 'sale' AND s.quantity >= 1000
    """
    return pd.read_sql(q, con).drop_duplicates(["invoice", "stock_code"])


# ------------------------------------------------------------------ charts
def style(ax, title, ylabel=None):
    ax.set_facecolor(SURFACE); ax.figure.set_facecolor(SURFACE)
    ax.set_title(title, loc="left", fontsize=12, color=INK, pad=12)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color(GRID)
    ax.tick_params(colors=INK2, labelsize=8)
    ax.grid(axis="y", color=GRID, linewidth=0.8); ax.set_axisbelow(True)
    if ylabel:
        ax.set_ylabel(ylabel, color=INK2, fontsize=9)


def charts(t: dict[str, pd.DataFrame]) -> None:
    FIGS.mkdir(parents=True, exist_ok=True)
    m = t["monthly_kpis"].copy()
    m["label"] = pd.to_datetime(m["month"]).dt.strftime("%Y-%m")
    fig, ax = plt.subplots(figsize=(10, 4.2))
    colors = [GREY if p else BLUE for p in m["is_partial_month"]]
    ax.bar(m["label"], m["net_sales"] / 1e3, color=colors, width=0.7)
    style(ax, "Net sales by month (GBP thousands) - Q4 peak each year", "GBP k")
    ax.set_xticks(range(len(m))); ax.set_xticklabels(m["label"], rotation=60)
    last = len(m) - 1
    if m["is_partial_month"].iloc[-1]:
        ax.annotate("partial month\n(data ends 9 Dec)", (last, m["net_sales"].iloc[-1] / 1e3), xytext=(0, 8),
                    textcoords="offset points", ha="center", fontsize=8, color=INK2)
    fig.tight_layout(); fig.savefig(FIGS / "01_monthly_net_sales.png", dpi=160); plt.close(fig)

    c = t["cohort_retention"].copy()
    c = c[(c["months_since"] >= 1) & (c["months_since"] <= 12)]
    sizes = t["cohort_retention"].query("months_since == 0").set_index("cohort_month")["customers"]
    big = sizes[(sizes >= 100) & (sizes.index > sizes.index.min())].index   # exclude the pre-existing base
    first_cohort = c[c["cohort_month"] == sizes.index.min()]
    c = c[c["cohort_month"].isin(big)]
    curve = (c.assign(w=c["cohort_month"].map(sizes))
               .groupby("months_since").apply(lambda g: np.average(g["retention_rate"], weights=g["w"]),
                                              include_groups=False))
    first = first_cohort.set_index("months_since")["retention_rate"]
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.plot(curve.index, curve.values * 100, color=BLUE, linewidth=2, marker="o", markersize=5,
            label="New-customer cohorts (100+ customers, weighted avg)")
    ax.plot(first.index, first.values * 100, color=GREY, linewidth=2, marker="o", markersize=5,
            label="Dec-2009 cohort (existing customer base)")
    style(ax, "Share of a first-purchase cohort buying again N months later", "% of cohort active")
    ax.set_xlabel("Months since first purchase", color=INK2, fontsize=9)
    ax.set_xticks(range(1, 13)); ax.set_ylim(0, max(first.max(), curve.max()) * 100 * 1.2)
    ax.legend(frameon=False, fontsize=8)
    fig.tight_layout(); fig.savefig(FIGS / "02_cohort_retention.png", dpi=160); plt.close(fig)


# ------------------------------------------------------------------ Excel report
HEAD = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="2A78D6")
INPUT_FILL = PatternFill("solid", fgColor="FFF4CC")


def add_table(wb, sheet_name: str, table_name: str, df: pd.DataFrame, formats: dict[str, str] | None = None):
    ws = wb.create_sheet(sheet_name)
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
    ref = f"A1:{get_column_letter(df.shape[1])}{len(df) + 1}"
    tab = Table(displayName=table_name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    for j, col in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(11, min(40, len(col) + 3))
        fmt = (formats or {}).get(col)
        if fmt:
            for r in range(2, len(df) + 2):
                ws.cell(r, j).number_format = fmt
    ws.freeze_panes = "A2"
    return ws


def excel_report(t: dict[str, pd.DataFrame]) -> None:
    XLSX.parent.mkdir(parents=True, exist_ok=True)
    m = t["monthly_kpis"].copy()
    m.insert(0, "month_label", pd.to_datetime(m["month"]).dt.strftime("%Y-%m"))
    m = m.drop(columns="month")
    reg = t["monthly_region_sales"].copy()
    reg.insert(0, "month_label", pd.to_datetime(reg["month"]).dt.strftime("%Y-%m"))
    reg = reg.drop(columns="month")

    wb = Workbook()
    rep = wb.active
    rep.title = "Report"
    money, pct, intf = '£#,##0', '0.0%', '#,##0'

    add_table(wb, "Data_Monthly", "tblMonthly", m, {
        "gross_sales": money, "cancellations": money, "net_sales": money, "avg_order_value": '£#,##0.00',
        "cancellation_rate": pct, "guest_share_of_sales": pct, "net_sales_mom": pct, "net_sales_yoy": pct,
        "orders": intf, "units": intf})
    add_table(wb, "Data_RegionMonth", "tblRegion", reg, {"gross_sales": money, "net_sales": money})
    add_table(wb, "Data_Country", "tblCountry", t["country_summary"],
              {"gross_sales": money, "net_sales": money, "cancellation_rate": pct})
    add_table(wb, "Data_TopProducts", "tblProducts", t["top_products"],
              {"gross_sales": money, "net_sales": money, "cancellation_rate": pct})
    add_table(wb, "Data_Quality", "tblQuality", pd.read_csv(ROOT / "reports" / "data_quality_log.csv"))

    # ---- Report sheet: everything below is a formula driven by the month in C3
    rep["B1"] = "Monthly Sales Operations Report"
    rep["B1"].font = Font(bold=True, size=16)
    rep["B3"], rep["C3"] = "Report month", m["month_label"].iloc[-2]      # default: last complete month
    rep["C3"].fill = INPUT_FILL
    rep["C3"].font = Font(bold=True)
    rep["D3"] = '=IF(INDEX(tblMonthly[is_partial_month],MATCH($C$3,tblMonthly[month_label],0))=1,' \
                '"Partial month - data ends mid-month; do not compare to full months","")'
    rep["D3"].font = Font(color="C00000", italic=True)
    dv = DataValidation(type="list", formula1=f"=Data_Monthly!$A$2:$A${len(m) + 1}", allow_blank=False)
    rep.add_data_validation(dv)
    dv.add("C3")
    rep["B4"] = "Pick a month in the yellow cell; every figure below recalculates."
    rep["B4"].font = Font(italic=True, color="52514E")

    for c, h in zip("BCDEF", ["KPI", "This month", "Prior month", "Same month last year", "MoM change"]):
        rep[f"{c}6"] = h
        rep[f"{c}6"].font, rep[f"{c}6"].fill = HEAD, HEAD_FILL

    def look(col: str, offset: int = 0) -> str:
        # INDEX into the column at (row of selected month + offset); blank if it does not exist
        pos = f"MATCH($C$3,tblMonthly[month_label],0)+{offset}" if offset else "MATCH($C$3,tblMonthly[month_label],0)"
        return f'=IFERROR(IF({pos}<1,"",INDEX(tblMonthly[{col}],{pos})),"")'

    kpis = [("Net sales", "net_sales", money), ("Gross sales", "gross_sales", money),
            ("Cancellations", "cancellations", money), ("Cancellation rate", "cancellation_rate", pct),
            ("Orders", "orders", intf), ("Average order value", "avg_order_value", '£#,##0.00'),
            ("Units sold", "units", intf), ("Active customers", "active_customers", intf),
            ("New customers", "new_customers", intf), ("Guest share of sales", "guest_share_of_sales", pct)]
    for i, (label, col, fmt) in enumerate(kpis, start=7):
        rep[f"B{i}"] = label
        rep[f"C{i}"] = look(col)
        rep[f"D{i}"] = look(col, -1)
        rep[f"E{i}"] = look(col, -12)
        rep[f"F{i}"] = f'=IF(OR(C{i}="",D{i}="",D{i}=0),"",C{i}/D{i}-1)'
        for c in "CDE":
            rep[f"{c}{i}"].number_format = fmt
        rep[f"F{i}"].number_format = '+0.0%;-0.0%;0.0%'
    last_kpi = 6 + len(kpis)
    red, green = PatternFill("solid", fgColor="F8D7D7"), PatternFill("solid", fgColor="D9F2D9")
    rep.conditional_formatting.add(f"F7:F{last_kpi}", CellIsRule(operator="lessThan", formula=["-0.1"], fill=red))
    rep.conditional_formatting.add(f"F7:F{last_kpi}", CellIsRule(operator="greaterThan", formula=["0.1"], fill=green))

    r0 = last_kpi + 2
    for c, h in zip("BCDE", ["Region", "Net sales", "Share of month", "Orders"]):
        rep[f"{c}{r0}"] = h
        rep[f"{c}{r0}"].font, rep[f"{c}{r0}"].fill = HEAD, HEAD_FILL
    for k, region in enumerate(["UK", "EU", "Rest of World"], start=1):
        r = r0 + k
        rep[f"B{r}"] = region
        rep[f"C{r}"] = f'=SUMIFS(tblRegion[net_sales],tblRegion[month_label],$C$3,tblRegion[region],B{r})'
        rep[f"D{r}"] = f'=IFERROR(C{r}/SUM($C${r0 + 1}:$C${r0 + 3}),"")'
        rep[f"E{r}"] = f'=SUMIFS(tblRegion[orders],tblRegion[month_label],$C$3,tblRegion[region],B{r})'
        rep[f"C{r}"].number_format, rep[f"D{r}"].number_format, rep[f"E{r}"].number_format = money, pct, intf
    rt = r0 + 4
    rep[f"B{rt}"] = "Total"
    rep[f"C{rt}"] = f"=SUM(C{r0 + 1}:C{r0 + 3})"
    rep[f"E{rt}"] = f"=SUM(E{r0 + 1}:E{r0 + 3})"
    rep[f"C{rt}"].number_format, rep[f"E{rt}"].number_format = money, intf
    rep[f"B{rt}"].font = Font(bold=True)
    rep[f"F{rt}"] = f'=IF(ABS(C{rt}-C7)<1,"ties to KPI table","CHECK: does not tie to net sales")'
    rep[f"F{rt}"].font = Font(italic=True, color="52514E")

    rr = rt + 2
    rep[f"B{rr}"] = "Year to date (calendar year of report month)"
    rep[f"B{rr}"].font = Font(bold=True)
    rep[f"B{rr + 1}"] = "Net sales YTD"
    rep[f"C{rr + 1}"] = ('=SUMPRODUCT((LEFT(tblMonthly[month_label],4)=LEFT($C$3,4))'
                         '*(tblMonthly[month_label]<=$C$3)*tblMonthly[net_sales])')
    rep[f"B{rr + 2}"] = "Best month YTD"
    rep[f"C{rr + 2}"] = ('=INDEX(tblMonthly[month_label],MATCH(_xlfn.MAXIFS(tblMonthly[net_sales],tblMonthly[month_label],'
                         '"<="&$C$3,tblMonthly[month_label],LEFT($C$3,4)&"*"),tblMonthly[net_sales],0))')
    rep[f"C{rr + 1}"].number_format = money

    rep.column_dimensions["A"].width = 2
    rep.column_dimensions["B"].width = 30
    for c in "CDEF":
        rep.column_dimensions[c].width = 20
    rep.sheet_view.showGridLines = False
    wb.save(XLSX)


# ------------------------------------------------------------------ findings
def findings(t, rev: pd.DataFrame, line_types: pd.DataFrame, dq: pd.DataFrame, seconds: float) -> str:
    m = t["monthly_kpis"].copy()
    m["label"] = pd.to_datetime(m["month"]).dt.strftime("%Y-%m")
    lt = line_types.set_index("line_type")
    gross_total = lt.loc["sale", "revenue_value"]
    net_total = gross_total + lt.loc["cancellation", "revenue_value"]
    guest_gross = m.eval("gross_sales * guest_share_of_sales").sum()
    full = m[m["is_partial_month"] == 0]
    y2010 = m[m["label"].str.startswith("2010")]["net_sales"].sum()
    p = t["customer_pareto"].set_index("top_pct_of_customers")
    cs = t["country_summary"]
    uk_share = cs.loc[cs["country"] == "United Kingdom", "net_sales"].iat[0] / cs["net_sales"].sum()
    dec11 = m[m["label"] == "2011-12"].iloc[0]
    rev_dec11 = rev[pd.to_datetime(rev["month"]).dt.strftime("%Y-%m") == "2011-12"]["line_revenue"].sum()
    adj_rate = (dec11.cancellations - rev_dec11) / (dec11.gross_sales - rev_dec11)
    nov = m[m["label"] == "2011-11"].iloc[0]
    hc = t["high_cancellation_products"].iloc[0]
    dq = dq.set_index("step")
    ov_row = [s for s in dq.index if s.startswith("invoices present in both")][0]
    ret = t["cohort_retention"]
    sizes = ret.query("months_since == 0").set_index("cohort_month")["customers"]
    later = ret[(ret["months_since"] == 3) & ret["cohort_month"].isin(sizes[sizes >= 100].index)
                & (ret["cohort_month"] > ret["cohort_month"].min())]
    m3 = np.average(later["retention_rate"], weights=later["cohort_month"].map(sizes))
    first3 = ret[(ret["cohort_month"] == ret["cohort_month"].min()) & (ret["months_since"] == 3)]["retention_rate"].iat[0]

    return f"""# Findings

_Generated by `src/report.py` from UCI Online Retail II (1,067,371 raw lines, Dec 2009 - 9 Dec 2011).
Currency: GBP. Report build time: {seconds:.0f} s from warehouse to finished Excel on the build machine._

## Data preparation that changes the numbers
| Issue | Lines | Effect if ignored |
|---|---:|---|
| Same invoices in both Excel sheets ({ov_row.split('(')[1].rstrip(')')}) | {int(dq.loc[ov_row, 'rows']):,} | Dec 2010 double-counted |
| Exact duplicate lines | {int(dq.loc['exact duplicate lines', 'rows']):,} | revenue overstated |
| Lines with no Customer ID | {int(dq.loc['lines with no Customer ID (guest)', 'rows']):,} | if dropped, **GBP {guest_gross:,.0f} ({guest_gross / gross_total:.1%}) of gross sales disappears** |
| Fees, postage, manual & bad-debt lines | {int(lt.loc['fee_or_service', 'lines'] + lt.loc['adjustment', 'lines']):,} | mixed into product revenue |
| Zero-price / negative-qty stock write-offs | {int(lt.loc['stock_writeoff', 'lines']):,} | distort units and price metrics |

## Headline
- Gross product sales GBP {gross_total:,.0f}; cancellations GBP {-lt.loc['cancellation', 'revenue_value']:,.0f};
  **net sales GBP {net_total:,.0f}**.
- Strong Q4 seasonality: November 2011 was the peak month (net GBP {nov.net_sales:,.0f}, {int(nov.orders):,} orders).
- United Kingdom is {uk_share:.0%} of net sales.

## Customers
- Revenue is concentrated: the top 1% of identified customers ({int(p.loc[1, 'customers_in_group'])}) bring
  {p.loc[1, 'share_of_customer_net_sales']:.0%} of customer net sales; the top 20% bring {p.loc[20, 'share_of_customer_net_sales']:.0%}.
- Retention: {first3:.0%} of the Dec-2009 group bought again 3 months later, but that group is the
  *existing* customer base (data starts that month). For genuinely new cohorts the 3-month figure is {m3:.0%}.
  Treating the first month as "new customers" would overstate acquisition and retention.

## Cancellations - one order can move the month
- {hc.description.title()} shows a {hc.cancellation_rate:.0%} cancellation rate: almost all of it is one
  74,215-unit order entered and cancelled the same day (Jan 2011).
- December 2011 reports a {dec11.cancellation_rate:.1%} cancellation rate. Excluding same-day
  order-and-cancel pairs of 1,000+ units, it is **{adj_rate:.1%}**. The monthly report should show both.
- December 2011 is also a **partial month** (data ends 9 Dec). Its {dec11.net_sales_yoy:.0%} YoY change is an
  artefact; the Excel report flags it automatically.

## Process
- The whole monthly pack (KPI tables, regional split, charts, Excel workbook) regenerates from one command.
  The Excel report is formula-driven: choosing a month in one cell refreshes every KPI, MoM and YoY comparison,
  regional split and YTD figure, with a built-in check that regional totals tie to net sales.
"""


def main() -> None:
    TABLES.mkdir(parents=True, exist_ok=True)
    t0 = time.perf_counter()
    with sqlite3.connect(DB) as con:
        t = run_sql(con)
        rev = same_day_reversals(con)
    for name, df in t.items():
        df.to_csv(TABLES / f"{name}.csv", index=False)
    rev.to_csv(TABLES / "same_day_reversals.csv", index=False)
    charts(t)
    excel_report(t)
    elapsed = time.perf_counter() - t0
    dq = pd.read_csv(ROOT / "reports" / "data_quality_log.csv")
    (ROOT / "reports" / "findings.md").write_text(findings(t, rev, t["line_type_summary"], dq, elapsed))
    print(f"tables={len(t)}  report={XLSX.name}  {elapsed:.1f}s")


if __name__ == "__main__":
    main()
