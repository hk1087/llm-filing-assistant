import pytest
import pandas as pd
from financials_builder import normalize_value, normalize_label, parse_statement_table


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


def test_normalize_value_dollar_space_parens():
    """Handle EDGAR 10-Q format: '$ (23)' should be -23.0"""
    assert normalize_value("$ (23)") == -23.0


def test_normalize_value_dollar_space_parens_with_commas():
    """Handle EDGAR 10-Q format: '$ (1,234)' should be -1234.0"""
    assert normalize_value("$ (1,234)") == -1234.0


def test_normalize_label_strips_whitespace():
    assert normalize_label("  Revenues  ") == "Revenues"


def test_normalize_label_collapses_internal_spaces():
    assert normalize_label("Cost  of  Revenue") == "Cost of Revenue"


def test_normalize_label_strips_trailing_colon():
    assert normalize_label("Total assets:") == "Total assets"


def test_normalize_label_unicode_dash():
    # Non-breaking spaces and unicode dashes should normalize
    assert normalize_label("Net\u00a0income") == "Net income"


# Tests for parse_statement_table
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
    assert col["values"][3] == -50.0    # Net income (loss) — index 3


def test_parse_statement_table_skips_blank_rows():
    result = parse_statement_table(SAMPLE_HTML)
    # blank row should be excluded
    assert len(result["labels"]) == 4


# Tests for align_to_canonical
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


def test_align_mismatched_lengths_raises():
    with pytest.raises(ValueError):
        align_to_canonical(["Revenues"], ["Revenues", "Net income"], [100.0])


def test_align_known_equivalence_basic_eps_dollars():
    """'Net loss per common share Basic (in dollars per share)' should match
    bare 'Basic (in dollars per share)' at the build_wide_df threshold (0.68)."""
    canonical = ["Net loss", "Net loss per common share Basic (in dollars per share)"]
    filing_labels = ["Net loss", "Basic (in dollars per share)"]
    filing_values = [-100.0, -1.32]
    result = align_to_canonical(canonical, filing_labels, filing_values, fuzzy_threshold=0.68)
    assert result[1] == -1.32


def test_align_known_equivalence_basic_eps_usd():
    """'Net loss per common share Basic (in dollars per share)' should match
    'Basic (in usd per share)' — different currency unit label."""
    canonical = ["Net loss", "Net loss per common share Basic (in dollars per share)"]
    filing_labels = ["Net loss", "Basic (in usd per share)"]
    filing_values = [-100.0, -1.32]
    result = align_to_canonical(canonical, filing_labels, filing_values, fuzzy_threshold=0.68)
    assert result[1] == -1.32


def test_align_known_equivalence_diluted_eps_usd():
    """'Net loss per common share Diluted (in dollars per share)' should match
    bare 'Diluted (in usd per share)' from older filings."""
    canonical = ["Net loss", "Net loss per common share Diluted (in dollars per share)"]
    filing_labels = ["Net loss", "Diluted (in usd per share)"]
    filing_values = [-100.0, -1.33]
    result = align_to_canonical(canonical, filing_labels, filing_values, fuzzy_threshold=0.68)
    assert result[1] == -1.33


def test_align_known_equivalence_loss_before_taxes():
    """'(Loss) income from continuing operations before income taxes' should match
    'Loss before income taxes' via the known-equivalences table, even though
    the fuzzy score (0.585) is below threshold."""
    canonical = ["Net revenue", "(Loss) income from continuing operations before income taxes"]
    filing_labels = ["Net revenue", "Loss before income taxes"]
    filing_values = [1000.0, -50.0]
    result = align_to_canonical(canonical, filing_labels, filing_values)
    assert result[1] == -50.0


# Tests for build_wide_df
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


def test_build_wide_df_columns_oldest_first():
    """Columns should be ordered oldest date first, newest last."""
    f1 = _make_filing_data(["Revenues"], "2024-12-31 (10-K)", [1000.0])
    f2 = _make_filing_data(["Revenues"], "2023-12-31 (10-K)", [900.0])
    df = build_wide_df([f1, f2])
    data_cols = [c for c in df.columns if c != "Line Item"]
    assert data_cols[0] == "2023-12-31 (10-K)"
    assert data_cols[-1] == "2024-12-31 (10-K)"


def test_build_wide_df_matches_net_loss_income_variant():
    """'Net loss' canonical should match 'Net (loss) income' from older filings."""
    f1 = _make_filing_data(
        ["Revenues", "Net loss"],
        "2025-12-31 (10-K)",
        [1000.0, -100.0],
    )
    f2 = _make_filing_data(
        ["Revenues", "Net (loss) income"],   # older filing variant
        "2024-12-31 (10-K)",
        [900.0, -40.0],
    )
    df = build_wide_df([f1, f2])
    nl_row = df[df["Line Item"] == "Net loss"]
    assert nl_row["2024-12-31 (10-K)"].iloc[0] == -40.0


def test_build_wide_df_matches_loss_variant():
    """'Net income' (10-K canonical) should match 'Net income (loss)' (10-Q label)."""
    f1 = _make_filing_data(
        ["Revenues", "Net income"],
        "2024-12-31 (10-K)",
        [1000.0, 100.0],
    )
    f2 = _make_filing_data(
        ["Revenues", "Net income (loss)"],   # 10-Q variant
        "2024-09-30 (10-Q)",
        [250.0, -18.0],
    )
    df = build_wide_df([f1, f2])
    ni_row = df[df["Line Item"] == "Net income"]
    assert ni_row["2024-09-30 (10-Q)"].iloc[0] == -18.0
