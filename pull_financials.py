# pull_financials.py
"""
Pull 3 years of SEC EDGAR financial statements into Excel.

Line items are in the exact same order and with the exact same labels
as the actual 10-K/10-Q filing.

Usage:
    python pull_financials.py --cik 0001364479
    python pull_financials.py --cik 0001364479 --years 3 --output hertz_financials.xlsx
"""
import argparse
import os
import time
import pandas as pd
from financials_builder import build_all_statements


SHEET_NAMES = {
    "income": "Income Statement",
    "balance": "Balance Sheet",
    "cashflow": "Cash Flow",
}


def write_excel(statements: dict, output_path: str, cik: str, years: int):
    def _write(path):
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for key, sheet_name in SHEET_NAMES.items():
                df = statements.get(key)
                if df is not None and not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            meta = pd.DataFrame([
                {"field": "CIK", "value": cik},
                {"field": "Years", "value": years},
                {"field": "Source", "value": "SEC EDGAR (HTML statement viewer)"},
            ])
            meta.to_excel(writer, sheet_name="Meta", index=False)

    try:
        _write(output_path)
        print(f"\nSaved {output_path}")
    except PermissionError:
        alt = f"{os.path.splitext(output_path)[0]}_{int(time.time())}.xlsx"
        print(f"Permission denied on {output_path}. Saving as {alt}")
        _write(alt)
        print(f"Saved {alt}")


def main():
    parser = argparse.ArgumentParser(
        description="Pull SEC financial statements into Excel with exact filing labels."
    )
    parser.add_argument("--cik", required=True, help="Company CIK, e.g. 0001364479")
    parser.add_argument("--years", type=int, default=3, help="Fiscal years to include (default 3)")
    parser.add_argument("--output", default=None, help="Output .xlsx path (default: <cik>_financials.xlsx)")
    parser.add_argument("--user-agent", default="Harneet Kaur harneet@example.com")
    args = parser.parse_args()

    cik = str(args.cik).zfill(10)
    output = args.output or f"{cik}_financials.xlsx"

    statements = build_all_statements(cik, years=args.years, user_agent=args.user_agent)

    for key, df in statements.items():
        print(f"\n{SHEET_NAMES[key]}: {len(df)} line items, {len(df.columns) - 1} periods")
        if not df.empty:
            print(df["Line Item"].head(5).to_string())

    write_excel(statements, output, cik, args.years)


if __name__ == "__main__":
    main()
