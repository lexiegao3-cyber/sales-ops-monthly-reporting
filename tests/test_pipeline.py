"""Consistency checks on the warehouse and reports. Run: python -m pytest -q"""
import sqlite3
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "warehouse" / "retail.sqlite"


@pytest.fixture(scope="module")
def con():
    if not DB.exists():
        pytest.skip("run `python run_all.py` first (needs data/raw/online_retail_II.xlsx)")
    with sqlite3.connect(DB) as c:
        yield c


def q(con, sql):
    return pd.read_sql(sql, con)


def test_no_lines_lost(con):
    log = pd.read_csv(ROOT / "reports" / "data_quality_log.csv").set_index("step")
    removed = log.filter(like="present in both sheets", axis=0)["rows"].sum() + log.loc["exact duplicate lines", "rows"]
    n = q(con, "SELECT COUNT(*) n FROM fact_sales_line").n[0]
    assert n == 1_067_371 - removed


def test_no_duplicate_invoice_across_sheets(con):
    d = q(con, """SELECT invoice FROM fact_sales_line GROUP BY invoice
                  HAVING COUNT(DISTINCT source_sheet) > 1""")
    assert d.empty


def test_cancellations_are_negative(con):
    d = q(con, "SELECT MAX(quantity) mx FROM fact_sales_line WHERE line_type = 'cancellation'")
    assert d.mx[0] < 0


def test_sales_lines_positive(con):
    d = q(con, "SELECT MIN(quantity) a, MIN(unit_price) b FROM fact_sales_line WHERE line_type = 'sale'")
    assert d.a[0] > 0 and d.b[0] > 0


def test_guests_kept_in_revenue(con):
    d = q(con, "SELECT SUM(line_revenue) s FROM fact_sales_line WHERE line_type='sale' AND is_guest=1")
    assert d.s[0] > 0


def test_monthly_ties_to_total(con):
    m = pd.read_csv(ROOT / "reports" / "tables" / "monthly_kpis.csv")
    total = q(con, """SELECT SUM(line_revenue) s FROM fact_sales_line
                      WHERE line_type IN ('sale','cancellation')""").s[0]
    assert abs(m["net_sales"].sum() - total) < 1


def test_region_ties_to_monthly():
    m = pd.read_csv(ROOT / "reports" / "tables" / "monthly_kpis.csv")
    r = pd.read_csv(ROOT / "reports" / "tables" / "monthly_region_sales.csv")
    assert abs(m["net_sales"].sum() - r["net_sales"].sum()) < 1


def test_only_last_month_partial():
    m = pd.read_csv(ROOT / "reports" / "tables" / "monthly_kpis.csv")
    assert m["is_partial_month"].tolist() == [0] * (len(m) - 1) + [1]
