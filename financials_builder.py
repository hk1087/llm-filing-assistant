import re
import difflib
from io import StringIO
import math
from collections import defaultdict, OrderedDict
import pandas as pd
import edgar_client


def normalize_value(val) -> float | None:
    """Convert a string cell value to float. Returns None for blank/dash."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return None if math.isnan(val) else float(val)
    s = str(val).strip()
    if not s or s in ("—", "–", "-", "N/A", "nan", "None"):
        return None
    # Strip dollar sign before checking for parenthetical negatives (handles "$ (23)" format)
    s = s.replace("$", "").strip()
    negative = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace(",", "").strip()
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
    except ValueError:
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
            raw = data_cols.iloc[row_idx, col_idx]
            values.append(normalize_value(raw))
        columns.append({"header": header, "values": values})

    return {"labels": labels, "columns": columns}


def _label_key(label: str) -> str:
    """Normalized key for label comparison: lowercase, no punctuation."""
    s = normalize_label(label).lower()
    s = re.sub(r"[^a-z0-9 ]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# Groups of label keys that refer to the same financial concept but appear with
# different wording across EDGAR filings. Any key in a group matches any other
# key in the same group, regardless of fuzzy score.
EQUIVALENT_LABEL_GROUPS: list[set] = [
    {
        "loss income from continuing operations before income taxes",
        "loss income before income taxes",
        "loss before income taxes",
        "income loss from continuing operations before income taxes",
    },
    # Total basic EPS — various label forms across EDGAR filing years
    {
        "net loss per common share basic in dollars per share",
        "net loss per common share basic in usd per share",
        "net income per common share basic in dollars per share",
        "net income per common share basic in usd per share",
        "basic in dollars per share",
        "basic in usd per share",
        "net income in usd per share",
        "net income in dollars per share",
        "net loss in usd per share",
        "net loss in dollars per share",
    },
    # Total diluted EPS
    {
        "net loss per common share diluted in dollars per share",
        "net loss per common share diluted in usd per share",
        "net income per common share diluted in dollars per share",
        "net income per common share diluted in usd per share",
        "diluted in dollars per share",
        "diluted in usd per share",
    },
]

# Pre-computed lookup: label key → group index
_EQUIV_MAP: dict[str, int] = {
    key: idx
    for idx, group in enumerate(EQUIVALENT_LABEL_GROUPS)
    for key in group
}


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
    # Validate that filing_labels and filing_values have the same length
    if len(filing_labels) != len(filing_values):
        raise ValueError(
            f"filing_labels ({len(filing_labels)}) and filing_values ({len(filing_values)}) must be the same length"
        )

    filing_keys = [_label_key(lab) for lab in filing_labels]
    canonical_keys = [_label_key(lab) for lab in canonical_labels]

    # Build lookup: normalized key → list of indices (in order of appearance)
    key_to_indices = defaultdict(list)
    for i, k in enumerate(filing_keys):
        key_to_indices[k].append(i)

    result = []
    used = set()

    for c_key in canonical_keys:
        # 1. Exact match — take first unused index for this key
        matched = False
        for idx in key_to_indices.get(c_key, []):
            if idx not in used:
                result.append(filing_values[idx])
                used.add(idx)
                matched = True
                break

        if matched:
            continue

        # 1.5 Known-equivalence match — catches labels too far apart for fuzzy
        c_group = _EQUIV_MAP.get(c_key)
        if c_group is not None:
            for i, f_key in enumerate(filing_keys):
                if i in used:
                    continue
                if _EQUIV_MAP.get(f_key) == c_group:
                    result.append(filing_values[i])
                    used.add(i)
                    matched = True
                    break

        if matched:
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


def build_wide_df(filings_data: list) -> pd.DataFrame:
    """
    Merge a list of parsed filing data into a wide DataFrame.

    filings_data: list of dicts, each with keys:
        - "labels": list of line-item label strings (in filing order)
        - "columns": list of {"header": str, "values": list}
        - "period": str
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

    rows = []
    for canon_label in canonical:
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
                    fuzzy_threshold=0.68,
                )
                row[header] = aligned[0]

        rows.append(row)

    df = pd.DataFrame(rows)
    period_cols = [c for c in df.columns if c != "Line Item"]
    return df[["Line Item"] + list(reversed(period_cols))]


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
            # avoiding duplicate headers when a 10-K table already has 2-3 comparative years.
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
