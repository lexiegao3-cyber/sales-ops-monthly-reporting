# Data

This project uses **Online Retail II** from the UCI Machine Learning Repository.

> Chen, D. (2012). *Online Retail II* [Dataset]. UCI Machine Learning Repository.
> https://doi.org/10.24432/C5CG6D — licence CC BY 4.0

All transactions of a UK-based online gift-ware retailer (many customers are
wholesalers), 1 Dec 2009 – 9 Dec 2011, 1,067,371 lines in two Excel sheets.

## Get the file
1. Open https://archive.ics.uci.edu/dataset/502/online+retail+ii and click **Download**.
2. Unzip and place `online_retail_II.xlsx` in `data/raw/`.
3. `python run_all.py` — the ETL checks the row count and the SHA-256
   (`bcbe73b3…2df2e980`) against the version this project was built on.

The raw file (43.5 MB) and the line-level output (~108 MB) are not committed;
dimension tables and all report outputs are.
