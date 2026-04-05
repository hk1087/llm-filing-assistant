# SEC EDGAR Financial Statement Puller

Pull Income Statement, Balance Sheet, and Cash Flow from SEC EDGAR into a single Excel workbook — with line items in the **exact same order and exact same labels** as the actual 10-K/10-Q filing.

## What it does

Given a company's CIK number, the tool:

- Fetches up to 3 years of 10-K and 10-Q filings from EDGAR
- Extracts all three financial statements from each filing
- Aligns rows across filings so each line item stays in one row, even when companies change their label wording between years
- Writes one Excel workbook with four sheets: Income Statement, Balance Sheet, Cash Flow, and Meta

Columns run oldest-to-newest left-to-right. Each column is one filing period (e.g. `2024-09-30 (10-Q)`).

## Usage

```bash
python pull_financials.py --cik 0001364479
python pull_financials.py --cik 0001364479 --years 3 --output hertz.xlsx
```

| Argument | Default | Description |
|---|---|---|
| `--cik` | required | Company CIK (e.g. `0001364479`) |
| `--years` | `3` | Number of fiscal years to include |
| `--output` | `<cik>_financials.xlsx` | Output file path |
| `--user-agent` | `Harneet Kaur harneet@example.com` | SEC requires a contact User-Agent header |

Find a company's CIK at [sec.gov/cgi-bin/browse-edgar](https://www.sec.gov/cgi-bin/browse-edgar).

## Requirements

```bash
pip install requests pandas openpyxl
```

Python 3.10+.

## How it works

1. **`edgar_client.py`** — fetches the EDGAR submissions JSON for the CIK, filters to 10-K/10-Q filings within the date range, finds each filing's `FilingSummary.xml`, and fetches the HTML statement viewer page (the `R*.htm` files SEC uses for interactive viewer)
2. **`financials_builder.py`** — parses each HTML table with `pandas.read_html`, normalizes values (handles `(1,234)`, `$ (23)`, dashes, etc.), then aligns rows across filings using:
   - Exact label matching (normalized, case-insensitive)
   - Fuzzy string matching via `difflib.SequenceMatcher`
   - A hand-curated equivalences table for labels that are too different to fuzzy-match but refer to the same concept (e.g. `"Loss before income taxes"` → `"(Loss) income from continuing operations before income taxes"`)
3. **`pull_financials.py`** — orchestrates the above and writes the Excel file

## Caveats

- **Row order follows the most recent filing.** The latest 10-K sets the canonical row order; older filings are aligned to it. If a company adds or removes a line item over time, older periods will show blank for that row.
- **Units are as reported.** Some companies report 10-Ks in thousands and 10-Qs in millions. The tool preserves the numbers as printed in the filing.
- **Section headers show as blank.** Rows like "Revenues" or "Operating costs and expenses" that are bold headers with no value in the filing will be blank across all periods — this is correct.
- **SEC rate limits.** EDGAR asks that automated tools not exceed 10 requests/second. The tool makes one request per filing per statement type, so for 13 filings you'll see ~40 requests total.

## Running tests

```bash
pytest
```

32 tests covering value normalization, label normalization, HTML table parsing, row alignment (exact, fuzzy, and equivalence matching), and the wide-DataFrame builder.
