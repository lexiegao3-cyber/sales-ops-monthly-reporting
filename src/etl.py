"""
ETL: UCI Online Retail II (online_retail_II.xlsx) -> classified line-level fact,
star schema in SQLite, and a data-quality log.

Source: Chen, D. (2012). Online Retail II [Dataset]. UCI Machine Learning Repository.
https://doi.org/10.24432/C5CG6D  (CC BY 4.0). Not committed - see data/README.md.

Decisions (all counted in reports/data_quality_log.csv):
  1. The two sheets overlap (2010-12-01..2010-12-09 appears in both). The
     overlapping invoices are verified identical, then kept once.
  2. Every line gets a line_type instead of being deleted:
        sale            normal product line
        cancellation    invoice starts with "C" (negative quantity)
        adjustment      invoice starts with "A" (bad-debt write-off)
        fee_or_service  postage, carriage, manual, discount, bank/marketplace fees,
                        samples, vouchers, tests (non-merchandise stock codes)
        stock_writeoff  non-cancelled line with quantity <= 0 or price <= 0
                        (warehouse notes like "damaged", "?", "check")
  3. Lines without a Customer ID are KEPT and flagged as guest. They are real
     sales; dropping them understates revenue. They are excluded only from
     customer-level metrics (retention, RFM).
  4. Exact duplicate lines within an invoice are removed (scanner double-posts).
"""
from __future__ import annotations

import hashlib
import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw" / "online_retail_II.xlsx"
CACHE = ROOT / "data" / "interim" / "online_retail_II_raw.pkl"
PROCESSED = ROOT / "data" / "processed"
WAREHOUSE = ROOT / "warehouse" / "retail.sqlite"
REPORTS = ROOT / "reports"

EXPECTED_SHA256 = "bcbe73b35f5b7babf197fb0cb983a11f5d9ff929078d4aa53d171b1f2df2e980"
EXPECTED_ROWS = 1_067_371
PRODUCT_CODE = r"^(\d{5}[A-Za-z]{0,2}|DCGS\w+)$"   # 5-digit codes (+variant letters) and dotcom gift sets
COUNTRY_REGION = {
    "United Kingdom": "UK",
    **{c: "EU" for c in ["Austria", "Belgium", "Cyprus", "Denmark", "EIRE", "Finland", "France", "Germany",
                         "Greece", "Italy", "Lithuania", "Malta", "Netherlands", "Poland", "Portugal",
                         "Spain", "Sweden", "Czech Republic", "Channel Islands", "European Community"]},
}


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load_raw() -> pd.DataFrame:
    if not RAW.exists():
        raise SystemExit(f"Missing {RAW}. Download it first - see data/README.md")
    if CACHE.exists():
        return pd.read_pickle(CACHE)
    digest = sha256(RAW)
    if digest != EXPECTED_SHA256:
        print(f"WARNING: file hash {digest[:12]}... differs from the version this project was built on")
    sheets = pd.read_excel(RAW, sheet_name=None, dtype={"Invoice": str, "StockCode": str})
    parts = [df.assign(source_sheet=name) for name, df in sheets.items()]
    df = pd.concat(parts, ignore_index=True)
    assert len(df) == EXPECTED_ROWS, f"expected {EXPECTED_ROWS} rows, got {len(df)}"
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    df.to_pickle(CACHE)
    return df


def clean(df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    log: list[dict] = []
    df = df.rename(columns={"Invoice": "invoice", "StockCode": "stock_code", "Description": "description",
                            "Quantity": "quantity", "InvoiceDate": "invoice_ts", "Price": "unit_price",
                            "Customer ID": "customer_id", "Country": "country"})
    df["invoice"] = df["invoice"].astype(str).str.strip()
    df["stock_code"] = df["stock_code"].astype(str).str.strip()
    df["description"] = df["description"].astype("string").str.strip()
    log.append({"step": "raw rows (both sheets)", "rows": len(df), "action": "loaded"})

    # 1. sheet overlap
    s1, s2 = sorted(df["source_sheet"].unique())
    ov = set(df.loc[df.source_sheet == s1, "invoice"]) & set(df.loc[df.source_sheet == s2, "invoice"])
    cols = ["invoice", "stock_code", "quantity", "invoice_ts", "unit_price", "customer_id"]
    a = df.loc[(df.source_sheet == s1) & df.invoice.isin(ov), cols].sort_values(cols).reset_index(drop=True)
    b = df.loc[(df.source_sheet == s2) & df.invoice.isin(ov), cols].sort_values(cols).reset_index(drop=True)
    assert a.equals(b), "overlapping invoices differ between sheets - inspect before de-duplicating"
    drop = (df.source_sheet == s2) & df.invoice.isin(ov)
    log.append({"step": f"invoices present in both sheets ({len(ov):,} invoices)", "rows": int(drop.sum()),
                "action": "verified identical; second copy removed"})
    df = df.loc[~drop].copy()

    # 4. exact duplicate lines
    dup = df.duplicated(subset=cols + ["description"])
    log.append({"step": "exact duplicate lines", "rows": int(dup.sum()), "action": "removed"})
    df = df.loc[~dup].copy()

    # 2. line types
    is_cancel = df["invoice"].str.startswith("C")
    is_adjust = df["invoice"].str.startswith("A")
    is_product = df["stock_code"].str.match(PRODUCT_CODE)
    bad_qty_price = (df["quantity"] <= 0) | (df["unit_price"] <= 0)
    df["line_type"] = np.select(
        [is_adjust, is_cancel & is_product, is_cancel & ~is_product, ~is_product, bad_qty_price],
        ["adjustment", "cancellation", "fee_or_service", "fee_or_service", "stock_writeoff"],
        default="sale")
    for lt, n in df["line_type"].value_counts().items():
        log.append({"step": f"line_type = {lt}", "rows": int(n),
                    "action": "kept, used in revenue" if lt in ("sale", "cancellation") else "kept, excluded from product revenue"})

    # 3. guests
    df["is_guest"] = df["customer_id"].isna().astype(int)
    log.append({"step": "lines with no Customer ID (guest)", "rows": int(df["is_guest"].sum()),
                "action": "kept in revenue; excluded from customer metrics"})
    df["customer_id"] = df["customer_id"].astype("Int64")

    df["line_revenue"] = (df["quantity"] * df["unit_price"]).round(2)
    df["invoice_date"] = df["invoice_ts"].dt.normalize()
    df["month"] = df["invoice_ts"].dt.to_period("M").dt.to_timestamp()
    df["region"] = df["country"].map(COUNTRY_REGION).fillna("Rest of World")
    log.append({"step": "rows out", "rows": len(df), "action": "line-level fact"})
    return df.reset_index(drop=True), log


def star_schema(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    prod = df[df.line_type.isin(["sale", "cancellation"])]
    # most frequent description per code (descriptions drift over time)
    desc = (prod.dropna(subset=["description"])
                .groupby(["stock_code", "description"]).size().reset_index(name="n")
                .sort_values(["stock_code", "n"], ascending=[True, False])
                .drop_duplicates("stock_code")[["stock_code", "description"]])
    dim_product = (prod.groupby("stock_code", as_index=False)
                       .agg(median_unit_price=("unit_price", "median"), first_sold=("invoice_date", "min"))
                       .merge(desc, on="stock_code", how="left"))
    dim_product["description"] = dim_product["description"].fillna("(no description)")

    cust = df[(df.is_guest == 0) & (df.line_type == "sale")]
    dim_customer = (cust.groupby("customer_id", as_index=False)
                        .agg(first_purchase=("invoice_date", "min"), last_purchase=("invoice_date", "max"),
                             home_country=("country", lambda s: s.mode().iat[0])))
    dim_customer["cohort_month"] = dim_customer["first_purchase"].dt.to_period("M").dt.to_timestamp()
    dim_customer["region"] = dim_customer["home_country"].map(COUNTRY_REGION).fillna("Rest of World")

    dim_country = (df[["country", "region"]].drop_duplicates().sort_values("country").reset_index(drop=True))

    fact = df[["invoice", "invoice_ts", "invoice_date", "month", "stock_code", "quantity", "unit_price",
               "line_revenue", "customer_id", "is_guest", "country", "line_type", "source_sheet"]].copy()
    return {"fact_sales_line": fact, "dim_product": dim_product, "dim_customer": dim_customer,
            "dim_country": dim_country}


def main() -> None:
    for p in (PROCESSED, WAREHOUSE.parent, REPORTS):
        p.mkdir(parents=True, exist_ok=True)
    df, log = clean(load_raw())
    pd.DataFrame(log).to_csv(REPORTS / "data_quality_log.csv", index=False)
    tables = star_schema(df)

    WAREHOUSE.unlink(missing_ok=True)
    with sqlite3.connect(WAREHOUSE) as con:
        for name, t in tables.items():
            out = t.copy()
            for c in out.select_dtypes("datetime").columns:
                out[c] = out[c].dt.strftime("%Y-%m-%d %H:%M:%S" if c == "invoice_ts" else "%Y-%m-%d")
            out.to_sql(name, con, index=False)
            out.to_csv(PROCESSED / f"{name}.csv", index=False)
        con.executescript("""
            CREATE INDEX ix_f_month   ON fact_sales_line(month);
            CREATE INDEX ix_f_cust    ON fact_sales_line(customer_id);
            CREATE INDEX ix_f_stock   ON fact_sales_line(stock_code);
            CREATE INDEX ix_f_invoice ON fact_sales_line(invoice);
        """)
    print(pd.DataFrame(log).to_string(index=False))
    print({k: len(v) for k, v in tables.items()})


if __name__ == "__main__":
    main()
