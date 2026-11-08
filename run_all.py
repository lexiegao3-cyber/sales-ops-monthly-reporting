"""Raw Excel -> cleaned warehouse -> monthly report pack. Needs data/raw/online_retail_II.xlsx (see data/README.md)."""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
for step in ["src/etl.py", "src/report.py"]:
    print(f"\n=== {step}")
    subprocess.run([sys.executable, str(ROOT / step)], check=True)
print("\nDone: excel/sales_ops_monthly_report.xlsx and reports/findings.md")
