# SEC Financial Statements → Excel Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Pull 3 years of Income Statement, Balance Sheet, and Cash Flow data from SEC EDGAR into Excel, with line items in the exact same order and with the exact same labels as the actual 10-K/10-Q filing.

**Architecture:** Fetch FilingSummary.xml for each filing to locate the three statement HTML files, then read those HTML tables directly (preserving exact labels and order). The most recent 10-K supplies the canonical row order; 10-Qs contribute quarterly columns aligned to that order. Multiple filings are merged into a single wide DataFrame per statement type.

**Tech Stack:** Python 3, requests, pandas, BeautifulSoup4, openpyxl, difflib (stdlib), pytest

---

## File Structure

| File | Action | Responsibility |
|------|--------|----------------|
| `edgar_client.py` | Create | All EDGAR HTTP calls: fetch submissions JSON, fetch FilingSummary.xml, fetch statement HTML |
| `financials_builder.py` | Create | Parse HTML tables; normalize values; align labels across filings; build wide DataFrames |
| `pull_financials.py` | Create | CLI entry point (argparse → orchestrate → Excel output) |
| `tests/__init__.py` | Create | Package marker (empty file) |
| `tests/test_financials_builder.py` | Create | Unit tests for parsing and merging logic |
| `tests/test_edgar_client.py` | Create | Unit tests for EDGAR client (mocked HTTP) |

Existing files (`edgar_to_excel.py`, `filing_statements_to_excel.py`, etc.) are left untouched.

---

## Task 1: Scaffolding + Value/Label Normalization

**Files:**
- Create: `tests/__init__.py`
- Create: `financials_builder.py`
- Create: `tests/test_financials_builder.py`

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_financials_builder.py
import pytest
from financials_builder import normalize_value, normalize_label

def test_normalize_value_plain():
    assert normalize_value("1,234") == 1234.0

def test_normalize_value_negative_parens():
    assert normalize_value("(1,234)") == -1234.0

def test_normalize_value_dollar():
    assert normalize_value("$1,234") == 1234.0

def test_normalize_value_dash():
    assert normalize_value("—") is None

def test_normalize_value_empty():
    assert normalize_value("") is None

def test_normalize_value_already_numeric():
    assert normalize_value(1234) == 1234.0

def test_normalize_label_strips_whitespace():
    assert normalize_label("  Revenues  ") == "Revenues"

def test_normalize_label_collapses_internal_spaces():
    assert normalize_label("Cost  of  Revenue") == "Cost of Revenue"

def test_normalize_label_strips_trailing_colon():
    assert normalize_label("Total assets:") == "Total assets"

def test_normalize_label_unicode_dash():
    # Non-breaking spaces and unicode dashes should normalize
    assert normalize_label("Net\u00a0income") == "Net income"
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_financials_builder.py -v
```
Expected: `ModuleNotFoundError` or `ImportError` (module doesn't exist yet)

- [ ] **Step 3: Create `tests/__init__.py` (empty)**

```python
# tests/__init__.py
```

- [ ] **Step 4: Implement `financials_builder.py` with normalize functions**

```python
# financials_builder.py
import re
import difflib
import pandas as pd
from collections import OrderedDict


def normalize_value(val) -> float | None:
    """Convert a string cell value to float. Returns None for blank/dash."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s in ("—", "–", "-", "N/A", "nan", "None"):
        return None
    negative = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace("$", "").replace(",", "").strip()
    try:
        result = float(s)
        return -result if negative else result
    except ValueError:
        return None


def normalize_label(label: str) -> str:
    """Normalize a line-item label for display and matching."""
    if not isinstance(label, str):
        label = str(label)
    # Replace non-breaking spaces and other unicode whitespace with regular space
    label = label.replace("\u00a0", " ").replace("\u2009", " ")
    # Collapse multiple spaces
    label = re.sub(r"\s+", " ", label).strip()
    # Strip trailing colon
    label = label.rstrip(":")
    return label
```

- [ ] **Step 5: Run tests to verify they pass**

```
pytest tests/test_financials_builder.py -v
```
Expected: all 10 tests PASS

- [ ] **Step 6: Commit**

```bash
git add financials_builder.py tests/__init__.py tests/test_financials_builder.py
git commit -m "feat: add financials_builder scaffold with value/label normalization"
```

---

## Task 2: Parse a Single Statement HTML Table

**Files:**
- Modify: `financials_builder.py`
- Modify: `tests/test_financials_builder.py`

The goal: given the raw HTML of an EDGAR statement viewer page (e.g., R2.htm), extract the line-item labels (in order) and the data values for each column.

- [ ] **Step 1: Write the failing test**

```python
# Add to tests/test_financials_builder.py
from financials_builder import parse_statement_table

SAMPLE_HTML = """
<html><body>
<table>
  <tr><th></th><th>Year Ended Dec 31, 2024</th><th>Year Ended Dec 31, 2023</th></tr>
  <tr><td>Revenues</td><td>1,000</td><td>900</td></tr>
  <tr><td>Cost of revenues</td><td>600</td><td>550</td></tr>
  <tr><td>Gross profit</td><td>400</td><td>350</td></tr>
  <tr><td></td><td></td><td></td></tr>
  <tr><td>Net income (loss)</td><td>(50)</td><td>80</td></tr>
</table>
</body></html>
"""

def test_parse_statement_table_labels():
    result = parse_statement_table(SAMPLE_HTML)
    assert result["labels"] == ["Revenues", "Cost of revenues", "Gross profit", "Net income (loss)"]

def test_parse_statement_table_column_count():
    result = parse_statement_table(SAMPLE_HTML)
    assert len(result["columns"]) == 2

def test_parse_statement_table_column_headers():
    result = parse_statement_table(SAMPLE_HTML)
    headers = [c["header"] for c in result["columns"]]
    assert "Year Ended Dec 31, 2024" in headers[0]
    assert "Year Ended Dec 31, 2023" in headers[1]

def test_parse_statement_table_values():
    result = parse_statement_table(SAMPLE_HTML)
    col = result["columns"][0]  # 2024
    assert col["values"][0] == 1000.0   # Revenues
    assert col["values"][1] == 600.0    # Cost of revenues
    assert col["values"][4 - 1] == -50.0  # Net income (loss) — index 3

def test_parse_statement_table_skips_blank_rows():
    result = parse_statement_table(SAMPLE_HTML)
    # blank row should be excluded
    assert len(result["labels"]) == 4
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_financials_builder.py::test_parse_statement_table_labels -v
```
Expected: FAIL with `ImportError` (function not defined yet)

- [ ] **Step 3: Implement `parse_statement_table`**

```python
# Add to financials_builder.py
from io import StringIO
from bs4 import BeautifulSoup


def parse_statement_table(html: str) -> dict:
    """
    Parse an EDGAR statement HTML page.

    Returns:
        {
            "labels": ["Revenues", "Cost of revenues", ...],   # in filing order
            "columns": [
                {"header": "Year Ended Dec 31, 2024", "values": [1000.0, 600.0, ...]},
                ...
            ]
        }
    """
    try:
        tables = pd.read_html(StringIO(html), thousands=",")
    except Exception:
        return {"labels": [], "columns": []}

    if not tables:
        return {"labels": [], "columns": []}

    # Pick the largest table (by cell count)
    df = max(tables, key=lambda t: t.shape[0] * t.shape[1])

    if df.shape[1] < 2:
        return {"labels": [], "columns": []}

    # First column = labels; remaining = data columns
    label_col = df.iloc[:, 0]
    data_cols = df.iloc[:, 1:]

    # Normalize column headers
    col_headers = [normalize_label(str(c)) for c in data_cols.columns]

    # Build label list and value arrays, skipping blank rows
    labels = []
    row_indices = []
    for i, raw_label in enumerate(label_col):
        lab = normalize_label(str(raw_label)) if not pd.isna(raw_label) else ""
        if not lab or lab.lower() in ("nan", "none", ""):
            continue
        labels.append(lab)
        row_indices.append(i)

    columns = []
    for col_idx, header in enumerate(col_headers):
        values = []
        for row_idx in row_indices:
            raw = df.iloc[row_idx, col_idx + 1]
            values.append(normalize_value(raw))
        columns.append({"header": header, "values": values})

    return {"labels": labels, "columns": columns}
```

- [ ] **Step 4: Run tests to verify they pass**

```
pytest tests/test_financials_builder.py -v
```
Expected: all tests PASS

- [ ] **Step 5: Commit**

```bash
git add financials_builder.py tests/test_financials_builder.py
git commit -m "feat: add parse_statement_table to read EDGAR HTML statement into structured data"
```

---

## Task 3: Label Alignment Across Filings

**Files:**
- Modify: `financials_builder.py`
- Modify: `tests/test_financials_builder.py`

When pulling multiple quarters/years, labels shift slightly (e.g., "Total revenues" vs "Revenues, net"). This function aligns an older filing's values to the canonical label order from the most recent filing.

- [ ] **Step 1: Write the failing tests**

```python
# Add to tests/test_financials_builder.py
from financials_builder import align_to_canonical

def test_align_exact_match():
    canonical = ["Revenues", "Cost of revenues", "Net income"]
    filing_labels = ["Revenues", "Cost of revenues", "Net income"]
    filing_values = [100.0, 60.0, 40.0]
    result = align_to_canonical(canonical, filing_labels, filing_values)
    assert result == [100.0, 60.0, 40.0]

def test_align_case_insensitive():
    canonical = ["Revenues", "Net income"]
    filing_labels = ["revenues", "net income"]
    filing_values = [100.0, 40.0]
    result = align_to_canonical(canonical, filing_labels, filing_values)
    assert result == [100.0, 40.0]

def test_align_missing_label_returns_none():
    canonical = ["Revenues", "Cost of revenues", "Net income"]
    filing_labels = ["Revenues", "Net income"]
    filing_values = [100.0, 40.0]
    result = align_to_canonical(canonical, filing_labels, filing_values)
    assert result[0] == 100.0
    assert result[1] is None  # Cost of revenues not in filing
    assert result[2] == 40.0

def test_align_fuzzy_match():
    canonical = ["Total revenues"]
    filing_labels = ["Revenues, net"]
    filing_values = [500.0]
    result = align_to_canonical(canonical, filing_labels, filing_values)
    # fuzzy match should find it
    assert result[0] == 500.0
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_financials_builder.py::test_align_exact_match -v
```
Expected: FAIL with `ImportError`

- [ ] **Step 3: Implement `align_to_canonical`**

```python
# Add to financials_builder.py

def _label_key(label: str) -> str:
    """Normalized key for label comparison: lowercase, no punctuation."""
    s = normalize_label(label).lower()
    s = re.sub(r"[^a-z0-9 ]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def align_to_canonical(
    canonical_labels: list,
    filing_labels: list,
    filing_values: list,
    fuzzy_threshold: float = 0.6,
) -> list:
    """
    Align filing_values to canonical_labels order.

    For each canonical label, find the best matching filing label using:
    1. Exact normalized match
    2. Fuzzy match (difflib SequenceMatcher)

    Returns list of values (same length as canonical_labels), with None
    where no match was found.
    """
    filing_keys = [_label_key(lab) for lab in filing_labels]
    canonical_keys = [_label_key(lab) for lab in canonical_labels]

    # Build a lookup: normalized key → index in filing
    key_to_idx = {k: i for i, k in enumerate(filing_keys)}

    result = []
    used = set()

    for c_key in canonical_keys:
        # 1. Exact match
        if c_key in key_to_idx and key_to_idx[c_key] not in used:
            idx = key_to_idx[c_key]
            result.append(filing_values[idx])
            used.add(idx)
            continue

        # 2. Fuzzy match
        best_idx = None
        best_score = fuzzy_threshold
        for i, f_key in enumerate(filing_keys):
            if i in used:
                continue
            score = difflib.SequenceMatcher(None, c_key, f_key).ratio()
            if score > best_score:
                best_score = score
                best_idx = i

        if best_idx is not None:
            result.append(filing_values[best_idx])
            used.add(best_idx)
        else:
            result.append(None)

    return result
```

- [ ] **Step 4: Run tests to verify they pass**

```
pytest tests/test_financials_builder.py -v
```
Expected: all tests PASS

- [ ] **Step 5: Commit**

```bash
git add financials_builder.py tests/test_financials_builder.py
git commit -m "feat: add align_to_canonical for merging label rows across filings"
```

---

## Task 4: EDGAR Client — Fetch Submissions and List Filings

**Files:**
- Create: `edgar_client.py`
- Create: `tests/test_edgar_client.py`

- [ ] **Step 1: Write the failing tests**

```python
# tests/test_edgar_client.py
from unittest.mock import patch, MagicMock
import json
import pytest
from edgar_client import list_filings

MOCK_SUBMISSIONS = {
    "filings": {
        "recent": {
            "form": ["10-K", "10-Q", "10-Q", "10-Q", "10-Q", "8-K"],
            "accessionNumber": [
                "0001364479-25-000001",
                "0001364479-24-000010",
                "0001364479-24-000009",
                "0001364479-24-000008",
                "0001364479-23-000007",
                "0001364479-24-000099",
            ],
            "primaryDocument": [
                "htz-20241231.htm",
                "htz-20240930.htm",
                "htz-20240630.htm",
                "htz-20240331.htm",
                "htz-20231231.htm",
                "8k.htm",
            ],
            "periodOfReport": [
                "2024-12-31",
                "2024-09-30",
                "2024-06-30",
                "2024-03-31",
                "2023-12-31",
                "2024-11-15",
            ],
            "filingDate": [
                "2025-02-15",
                "2024-11-10",
                "2024-08-10",
                "2024-05-10",
                "2024-02-15",
                "2024-11-15",
            ],
        }
    }
}

@patch("edgar_client.requests.get")
def test_list_filings_returns_10k_and_10q(mock_get):
    mock_resp = MagicMock()
    mock_resp.json.return_value = MOCK_SUBMISSIONS
    mock_resp.raise_for_status = MagicMock()
    mock_get.return_value = mock_resp

    filings = list_filings("0001364479", form_types=["10-K", "10-Q"], years=3)

    forms = [f["form"] for f in filings]
    assert "10-K" in forms
    assert "10-Q" in forms
    assert "8-K" not in forms

@patch("edgar_client.requests.get")
def test_list_filings_excludes_old(mock_get):
    mock_resp = MagicMock()
    mock_resp.json.return_value = MOCK_SUBMISSIONS
    mock_resp.raise_for_status = MagicMock()
    mock_get.return_value = mock_resp

    filings = list_filings("0001364479", form_types=["10-K", "10-Q"], years=1)
    periods = [f["period"] for f in filings]
    # With years=1, should only include filings within 1 fiscal year of most recent
    for p in periods:
        assert p >= "2024-01-01"

@patch("edgar_client.requests.get")
def test_list_filings_sorted_newest_first(mock_get):
    mock_resp = MagicMock()
    mock_resp.json.return_value = MOCK_SUBMISSIONS
    mock_resp.raise_for_status = MagicMock()
    mock_get.return_value = mock_resp

    filings = list_filings("0001364479", form_types=["10-K", "10-Q"], years=3)
    periods = [f["period"] for f in filings]
    assert periods == sorted(periods, reverse=True)

@patch("edgar_client.requests.get")
def test_list_filings_includes_required_fields(mock_get):
    mock_resp = MagicMock()
    mock_resp.json.return_value = MOCK_SUBMISSIONS
    mock_resp.raise_for_status = MagicMock()
    mock_get.return_value = mock_resp

    filings = list_filings("0001364479")
    f = filings[0]
    assert "form" in f
    assert "accession" in f
    assert "period" in f
    assert "accession_nodashes" in f
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_edgar_client.py -v
```
Expected: FAIL with `ModuleNotFoundError`

- [ ] **Step 3: Implement `edgar_client.py` with `list_filings`**

```python
# edgar_client.py
import requests
from datetime import datetime, timedelta

USER_AGENT = "Harneet Kaur harneet@example.com"

def _session():
    s = requests.Session()
    s.headers.update({"User-Agent": USER_AGENT})
    return s


def list_filings(
    cik: str,
    form_types: list = None,
    years: int = 3,
    user_agent: str = None,
) -> list:
    """
    Return a list of filings for the given CIK, filtered by form_types and
    within the last `years` fiscal years, sorted newest-first.

    Each entry is a dict:
        {
            "form": "10-K",
            "accession": "0001364479-25-000001",
            "accession_nodashes": "0001364479250000001",
            "period": "2024-12-31",
            "primary_doc": "htz-20241231.htm",
        }
    """
    if form_types is None:
        form_types = ["10-K", "10-Q"]

    cik_padded = str(cik).zfill(10)
    s = _session()
    if user_agent:
        s.headers.update({"User-Agent": user_agent})

    url = f"https://data.sec.gov/submissions/CIK{cik_padded}.json"
    r = s.get(url, timeout=30)
    r.raise_for_status()
    data = r.json()

    recent = data["filings"]["recent"]
    forms = recent["form"]
    accessions = recent["accessionNumber"]
    primary_docs = recent["primaryDocument"]
    periods = recent["periodOfReport"]

    results = []
    for i, form in enumerate(forms):
        if form not in form_types:
            continue
        results.append({
            "form": form,
            "accession": accessions[i],
            "accession_nodashes": accessions[i].replace("-", ""),
            "period": periods[i],
            "primary_doc": primary_docs[i],
        })

    if not results:
        return []

    # Sort newest first
    results.sort(key=lambda f: f["period"], reverse=True)

    # Filter to last `years` fiscal years from most recent period
    most_recent = results[0]["period"]
    cutoff_dt = datetime.strptime(most_recent, "%Y-%m-%d") - timedelta(days=years * 366)
    cutoff = cutoff_dt.strftime("%Y-%m-%d")

    return [f for f in results if f["period"] >= cutoff]
```

- [ ] **Step 4: Run tests to verify they pass**

```
pytest tests/test_edgar_client.py -v
```
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add edgar_client.py tests/test_edgar_client.py
git commit -m "feat: add edgar_client with list_filings"
```

---

## Task 5: EDGAR Client — Fetch FilingSummary.xml and Find Statement HTML Files

**Files:**
- Modify: `edgar_client.py`
- Modify: `tests/test_edgar_client.py`

- [ ] **Step 1: Write the failing tests**

```python
# Add to tests/test_edgar_client.py
from edgar_client import find_statement_report

SAMPLE_FILING_SUMMARY_XML = b"""<?xml version="1.0" encoding="utf-8"?>
<FilingSummary>
  <MyReports>
    <Report instance="R1.htm">
      <ShortName>Document and Entity Information</ShortName>
      <LongName>Document And Entity Information</LongName>
      <HtmlFileName>R1.htm</HtmlFileName>
    </Report>
    <Report instance="R2.htm">
      <ShortName>CONSOLIDATED STATEMENTS OF OPERATIONS</ShortName>
      <LongName>Consolidated Statements of Operations</LongName>
      <HtmlFileName>R2.htm</HtmlFileName>
    </Report>
    <Report instance="R3.htm">
      <ShortName>CONSOLIDATED BALANCE SHEETS</ShortName>
      <LongName>Consolidated Balance Sheets</LongName>
      <HtmlFileName>R3.htm</HtmlFileName>
    </Report>
    <Report instance="R4.htm">
      <ShortName>CONSOLIDATED STATEMENTS OF CASH FLOWS</ShortName>
      <LongName>Consolidated Statements of Cash Flows</LongName>
      <HtmlFileName>R4.htm</HtmlFileName>
    </Report>
  </MyReports>
</FilingSummary>"""


@patch("edgar_client.requests.get")
def test_find_income_statement(mock_get):
    mock_resp = MagicMock()
    mock_resp.content = SAMPLE_FILING_SUMMARY_XML
    mock_resp.status_code = 200
    mock_resp.raise_for_status = MagicMock()
    mock_get.return_value = mock_resp

    reports = find_statement_report(
        "0001364479", "00013644792500001", "income"
    )
    assert reports is not None
    assert reports["html_file"] == "R2.htm"


@patch("edgar_client.requests.get")
def test_find_balance_sheet(mock_get):
    mock_resp = MagicMock()
    mock_resp.content = SAMPLE_FILING_SUMMARY_XML
    mock_resp.status_code = 200
    mock_resp.raise_for_status = MagicMock()
    mock_get.return_value = mock_resp

    reports = find_statement_report(
        "0001364479", "00013644792500001", "balance"
    )
    assert reports is not None
    assert reports["html_file"] == "R3.htm"


@patch("edgar_client.requests.get")
def test_find_cashflow(mock_get):
    mock_resp = MagicMock()
    mock_resp.content = SAMPLE_FILING_SUMMARY_XML
    mock_resp.status_code = 200
    mock_resp.raise_for_status = MagicMock()
    mock_get.return_value = mock_resp

    reports = find_statement_report(
        "0001364479", "00013644792500001", "cashflow"
    )
    assert reports is not None
    assert reports["html_file"] == "R4.htm"
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_edgar_client.py::test_find_income_statement -v
```
Expected: FAIL with `ImportError`

- [ ] **Step 3: Implement `find_statement_report` in `edgar_client.py`**

```python
# Add to edgar_client.py
import xml.etree.ElementTree as ET

STATEMENT_CANDIDATES = {
    "income": [
        "statement of operations",
        "statements of operations",
        "statement of income",
        "statements of income",
        "statement of earnings",
        "statements of earnings",
    ],
    "balance": [
        "balance sheet",
        "balance sheets",
        "statement of financial position",
    ],
    "cashflow": [
        "statement of cash flows",
        "statements of cash flows",
        "cash flows",
    ],
}


def find_statement_report(
    cik: str,
    accession_nodashes: str,
    statement_type: str,
    user_agent: str = None,
) -> dict | None:
    """
    Fetch FilingSummary.xml for a filing and find the HTML file for the
    given statement_type ('income', 'balance', 'cashflow').

    Returns dict with keys: short_name, long_name, html_file
    Returns None if not found.
    """
    cik_int = str(int(cik))
    base = f"https://www.sec.gov/Archives/edgar/data/{cik_int}/{accession_nodashes}/"
    url = base + "FilingSummary.xml"

    s = _session()
    if user_agent:
        s.headers.update({"User-Agent": user_agent})

    r = s.get(url, timeout=30)
    if r.status_code != 200:
        return None
    r.raise_for_status()

    root = ET.fromstring(r.content)
    reports = []
    for rep in root.findall(".//Report"):
        short = (rep.findtext("ShortName") or "").strip()
        long = (rep.findtext("LongName") or "").strip()
        html_file = (rep.findtext("HtmlFileName") or "").strip()
        reports.append({"short_name": short, "long_name": long, "html_file": html_file})

    candidates = STATEMENT_CANDIDATES.get(statement_type, [])
    for cand in candidates:
        cand_lower = cand.lower()
        for rep in reports:
            haystack = f"{rep['short_name']} {rep['long_name']}".lower()
            if cand_lower in haystack and rep["html_file"]:
                return rep

    return None
```

- [ ] **Step 4: Run tests to verify they pass**

```
pytest tests/test_edgar_client.py -v
```
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add edgar_client.py tests/test_edgar_client.py
git commit -m "feat: add find_statement_report to locate statement HTML via FilingSummary.xml"
```

---

## Task 6: EDGAR Client — Fetch Statement HTML

**Files:**
- Modify: `edgar_client.py`
- Modify: `tests/test_edgar_client.py`

- [ ] **Step 1: Write the failing test**

```python
# Add to tests/test_edgar_client.py
from edgar_client import fetch_statement_html

@patch("edgar_client.requests.get")
def test_fetch_statement_html_returns_string(mock_get):
    mock_resp = MagicMock()
    mock_resp.text = "<html><body><table><tr><td>Revenues</td><td>100</td></tr></table></body></html>"
    mock_resp.raise_for_status = MagicMock()
    mock_get.return_value = mock_resp

    html = fetch_statement_html("0001364479", "00013644792500001", "R2.htm")
    assert isinstance(html, str)
    assert "Revenues" in html
```

- [ ] **Step 2: Run test to verify it fails**

```
pytest tests/test_edgar_client.py::test_fetch_statement_html_returns_string -v
```
Expected: FAIL with `ImportError`

- [ ] **Step 3: Implement `fetch_statement_html` in `edgar_client.py`**

```python
# Add to edgar_client.py

def fetch_statement_html(
    cik: str,
    accession_nodashes: str,
    html_file: str,
    user_agent: str = None,
) -> str:
    """Fetch the HTML for a specific statement viewer file."""
    cik_int = str(int(cik))
    url = f"https://www.sec.gov/Archives/edgar/data/{cik_int}/{accession_nodashes}/{html_file}"
    s = _session()
    if user_agent:
        s.headers.update({"User-Agent": user_agent})
    r = s.get(url, timeout=60)
    r.raise_for_status()
    return r.text
```

- [ ] **Step 4: Run all tests**

```
pytest tests/ -v
```
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add edgar_client.py tests/test_edgar_client.py
git commit -m "feat: add fetch_statement_html to edgar_client"
```

---

## Task 7: Build Wide DataFrames from Multiple Filings

**Files:**
- Modify: `financials_builder.py`
- Modify: `tests/test_financials_builder.py`

This is the core merge step: given parsed data from multiple filings, build a single wide DataFrame where rows = line items (canonical order from most recent filing) and columns = one per filing period.

- [ ] **Step 1: Write the failing tests**

```python
# Add to tests/test_financials_builder.py
from financials_builder import build_wide_df

def _make_filing_data(labels, col_header, values):
    return {
        "labels": labels,
        "columns": [{"header": col_header, "values": values}],
        "period": col_header,
        "form": "10-K",
    }

def test_build_wide_df_single_filing():
    filing = _make_filing_data(
        ["Revenues", "Net income"],
        "2024-12-31 (10-K)",
        [1000.0, 100.0],
    )
    df = build_wide_df([filing])
    assert list(df["Line Item"]) == ["Revenues", "Net income"]
    assert "2024-12-31 (10-K)" in df.columns
    assert df["2024-12-31 (10-K)"].iloc[0] == 1000.0

def test_build_wide_df_two_filings_aligned():
    f1 = _make_filing_data(
        ["Revenues", "Cost of revenues", "Net income"],
        "2024-12-31 (10-K)",
        [1000.0, 600.0, 100.0],
    )
    f2 = _make_filing_data(
        ["Revenues", "Cost of revenues", "Net income"],
        "2023-12-31 (10-K)",
        [900.0, 550.0, 80.0],
    )
    df = build_wide_df([f1, f2])
    assert list(df["Line Item"]) == ["Revenues", "Cost of revenues", "Net income"]
    assert df["2023-12-31 (10-K)"].iloc[0] == 900.0

def test_build_wide_df_missing_label_is_none():
    f1 = _make_filing_data(
        ["Revenues", "Cost of revenues", "Net income"],
        "2024-12-31 (10-K)",
        [1000.0, 600.0, 100.0],
    )
    f2 = _make_filing_data(
        ["Revenues", "Net income"],  # missing "Cost of revenues"
        "2023-12-31 (10-K)",
        [900.0, 80.0],
    )
    df = build_wide_df([f1, f2])
    cost_row = df[df["Line Item"] == "Cost of revenues"]
    assert cost_row["2023-12-31 (10-K)"].iloc[0] is None or \
           pd.isna(cost_row["2023-12-31 (10-K)"].iloc[0])

def test_build_wide_df_columns_newest_first():
    f1 = _make_filing_data(["Revenues"], "2024-12-31 (10-K)", [1000.0])
    f2 = _make_filing_data(["Revenues"], "2023-12-31 (10-K)", [900.0])
    df = build_wide_df([f1, f2])
    data_cols = [c for c in df.columns if c != "Line Item"]
    # newest first
    assert data_cols[0] == "2024-12-31 (10-K)"
```

- [ ] **Step 2: Run tests to verify they fail**

```
pytest tests/test_financials_builder.py::test_build_wide_df_single_filing -v
```
Expected: FAIL with `ImportError`

- [ ] **Step 3: Implement `build_wide_df` in `financials_builder.py`**

```python
# Add to financials_builder.py

def build_wide_df(filings_data: list) -> pd.DataFrame:
    """
    Merge a list of parsed filing data into a wide DataFrame.

    filings_data: list of dicts, each with keys:
        - "labels": list of line-item label strings (in filing order)
        - "columns": list of {"header": str, "values": list}
        - "period": str  (used as column header, e.g. "2024-12-31 (10-K)")
        - "form": str

    The most recent filing (first in list, caller must sort newest-first)
    sets the canonical label order. Older filings are aligned to it.

    Returns a DataFrame with:
        Column 0: "Line Item"
        Remaining columns: one per filing period, newest first.
    """
    if not filings_data:
        return pd.DataFrame(columns=["Line Item"])

    # Canonical labels from the first filing (most recent)
    canonical = filings_data[0]["labels"]

    # Collect all column headers (period labels) across all filings
    # Each filing contributes the columns from its own "columns" list
    all_col_headers = []
    for fd in filings_data:
        for col in fd["columns"]:
            h = col["header"]
            if h not in all_col_headers:
                all_col_headers.append(h)

    rows = []
    for i, canon_label in enumerate(canonical):
        row = OrderedDict()
        row["Line Item"] = canon_label

        for fd in filings_data:
            filing_labels = fd["labels"]
            for col in fd["columns"]:
                header = col["header"]
                filing_values = col["values"]
                aligned = align_to_canonical(
                    [canon_label],
                    filing_labels,
                    filing_values,
                )
                row[header] = aligned[0]

        rows.append(row)

    df = pd.DataFrame(rows)
    # Ensure column order: Line Item first, then period columns
    period_cols = [c for c in df.columns if c != "Line Item"]
    return df[["Line Item"] + period_cols]
```

- [ ] **Step 4: Run tests to verify they pass**

```
pytest tests/test_financials_builder.py -v
```
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add financials_builder.py tests/test_financials_builder.py
git commit -m "feat: add build_wide_df to merge multiple filings into a wide DataFrame"
```

---

## Task 8: Orchestration — Fetch + Parse + Merge for All Three Statements

**Files:**
- Modify: `financials_builder.py`

This function ties together `edgar_client` and `financials_builder` to produce three DataFrames (income, balance, cashflow) ready for Excel output.

No new unit tests here — this is integration logic tested in Task 10.

- [ ] **Step 1: Implement `build_all_statements` in `financials_builder.py`**

```python
# Add to financials_builder.py
import edgar_client


def build_all_statements(
    cik: str,
    years: int = 3,
    user_agent: str = None,
) -> dict:
    """
    Fetch and merge financial statements for the given CIK.

    Returns:
        {
            "income": DataFrame,
            "balance": DataFrame,
            "cashflow": DataFrame,
        }
    """
    cik_padded = str(cik).zfill(10)

    print(f"Listing filings for CIK {cik_padded}...")
    filings = edgar_client.list_filings(cik_padded, years=years, user_agent=user_agent)
    print(f"Found {len(filings)} filings: {[f['period'] + ' ' + f['form'] for f in filings]}")

    statement_types = ["income", "balance", "cashflow"]
    # filing_data_by_type[type] = list of parsed dicts, newest first
    filing_data_by_type = {t: [] for t in statement_types}

    for filing in filings:
        acc = filing["accession_nodashes"]
        period = filing["period"]
        form = filing["form"]
        col_label = f"{period} ({form})"
        print(f"  Processing {col_label}...")

        for stmt_type in statement_types:
            report = edgar_client.find_statement_report(
                cik_padded, acc, stmt_type, user_agent=user_agent
            )
            if report is None:
                print(f"    No {stmt_type} report found for {col_label}")
                continue

            try:
                html = edgar_client.fetch_statement_html(
                    cik_padded, acc, report["html_file"], user_agent=user_agent
                )
            except Exception as e:
                print(f"    Error fetching {stmt_type} HTML: {e}")
                continue

            parsed = parse_statement_table(html)
            if not parsed["labels"]:
                print(f"    Empty {stmt_type} table for {col_label}")
                continue

            # Keep only the first data column (current period) from every filing.
            # The first column is always the most recent period being reported.
            # This gives each filing exactly one column in the merged output,
            # avoiding duplicate headers when a 10-K table already has 2–3 comparative years.
            if parsed["columns"]:
                current_col = parsed["columns"][0]
                current_col["header"] = col_label
                parsed["columns"] = [current_col]

            filing_data_by_type[stmt_type].append({
                "labels": parsed["labels"],
                "columns": parsed["columns"],
                "period": period,
                "form": form,
            })

    results = {}
    for stmt_type in statement_types:
        data = filing_data_by_type[stmt_type]
        results[stmt_type] = build_wide_df(data) if data else pd.DataFrame(columns=["Line Item"])

    return results
```

- [ ] **Step 2: Run existing tests to make sure nothing broke**

```
pytest tests/ -v
```
Expected: all PASS

- [ ] **Step 3: Commit**

```bash
git add financials_builder.py
git commit -m "feat: add build_all_statements orchestration function"
```

---

## Task 9: Excel Writer + CLI Entry Point

**Files:**
- Create: `pull_financials.py`

- [ ] **Step 1: Write `pull_financials.py`**

```python
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
```

- [ ] **Step 2: Run tests to make sure nothing broke**

```
pytest tests/ -v
```
Expected: all PASS

- [ ] **Step 3: Commit**

```bash
git add pull_financials.py
git commit -m "feat: add pull_financials.py CLI entry point with Excel output"
```

---

## Task 10: Smoke Test Against Real CIK

**Files:** None (manual verification)

- [ ] **Step 1: Install dependencies (if not already installed)**

```bash
pip install requests pandas openpyxl beautifulsoup4 lxml pytest
```

- [ ] **Step 2: Run against Hertz (CIK 0001364479) — 1 year to keep it fast**

```bash
python pull_financials.py --cik 0001364479 --years 1
```

Expected output:
```
Listing filings for CIK 0001364479...
Found N filings: [...]
  Processing 2024-12-31 (10-K)...
  Processing 2024-09-30 (10-Q)...
  ...

Income Statement: N line items, M periods
  [first 5 line items from the actual filing]

Balance Sheet: N line items, M periods
Cash Flow: N line items, M periods

Saved 0001364479_financials.xlsx
```

- [ ] **Step 3: Open Excel and verify**

Open `0001364479_financials.xlsx`.

Verify:
- [ ] Income Statement labels match what you see in the Hertz 10-K (open the actual filing at SEC.gov to compare)
- [ ] Row order matches the actual filing
- [ ] Values look reasonable (not wildly off, right sign convention)
- [ ] No `NaN` column headers
- [ ] Column headers show the period clearly (e.g., `2024-12-31 (10-K)`)

- [ ] **Step 4: Run against a second company to verify it generalizes**

```bash
python pull_financials.py --cik 0001657853 --years 2
```

Expected: succeeds without errors, produces Excel with recognizable financial statement labels.

- [ ] **Step 5: Run full test suite one final time**

```
pytest tests/ -v
```
Expected: all PASS

- [ ] **Step 6: Final commit**

```bash
git add .
git commit -m "feat: complete SEC financials-to-excel implementation with exact label/order preservation"
```

---

## Quick Reference

**Run everything:**
```bash
python pull_financials.py --cik 0001364479 --years 3
```

**Run tests:**
```bash
pytest tests/ -v
```

**Key design decisions:**
- HTML table parsing (not XBRL API) preserves exact labels and row order
- Most recent 10-K's labels are the canonical row template
- Older filings aligned via exact → fuzzy label matching
- 10-K tables contribute all their columns (already multi-year); 10-Q tables contribute only their first (current quarter) column
