from unittest.mock import patch, MagicMock
import json
import pytest
from edgar_client import list_filings, find_statement_report, fetch_statement_html

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
            "reportDate": [
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

@patch("edgar_client.requests.Session")
def test_list_filings_returns_10k_and_10q(mock_session_class):
    mock_session = MagicMock()
    mock_session_class.return_value = mock_session

    mock_resp = MagicMock()
    mock_resp.json.return_value = MOCK_SUBMISSIONS
    mock_resp.raise_for_status = MagicMock()
    mock_session.get.return_value = mock_resp

    filings = list_filings("0001364479", form_types=["10-K", "10-Q"], years=3)

    forms = [f["form"] for f in filings]
    assert "10-K" in forms
    assert "10-Q" in forms
    assert "8-K" not in forms

@patch("edgar_client.requests.Session")
def test_list_filings_excludes_old(mock_session_class):
    mock_session = MagicMock()
    mock_session_class.return_value = mock_session

    mock_resp = MagicMock()
    mock_resp.json.return_value = MOCK_SUBMISSIONS
    mock_resp.raise_for_status = MagicMock()
    mock_session.get.return_value = mock_resp

    filings = list_filings("0001364479", form_types=["10-K", "10-Q"], years=1)
    periods = [f["period"] for f in filings]
    # With years=1, should only include filings within 1 fiscal year of most recent
    for p in periods:
        assert p >= "2024-01-01"

@patch("edgar_client.requests.Session")
def test_list_filings_sorted_newest_first(mock_session_class):
    mock_session = MagicMock()
    mock_session_class.return_value = mock_session

    mock_resp = MagicMock()
    mock_resp.json.return_value = MOCK_SUBMISSIONS
    mock_resp.raise_for_status = MagicMock()
    mock_session.get.return_value = mock_resp

    filings = list_filings("0001364479", form_types=["10-K", "10-Q"], years=3)
    periods = [f["period"] for f in filings]
    assert periods == sorted(periods, reverse=True)

@patch("edgar_client.requests.Session")
def test_list_filings_includes_required_fields(mock_session_class):
    mock_session = MagicMock()
    mock_session_class.return_value = mock_session

    mock_resp = MagicMock()
    mock_resp.json.return_value = MOCK_SUBMISSIONS
    mock_resp.raise_for_status = MagicMock()
    mock_session.get.return_value = mock_resp

    filings = list_filings("0001364479")
    f = filings[0]
    assert "form" in f
    assert "accession" in f
    assert "period" in f
    assert "accession_nodashes" in f
    assert "primary_doc" in f


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


@patch("edgar_client.requests.Session")
def test_find_income_statement(mock_session_class):
    mock_session = MagicMock()
    mock_session_class.return_value = mock_session

    mock_resp = MagicMock()
    mock_resp.content = SAMPLE_FILING_SUMMARY_XML
    mock_resp.status_code = 200
    mock_resp.raise_for_status = MagicMock()
    mock_session.get.return_value = mock_resp

    reports = find_statement_report("0001364479", "00013644792500001", "income")
    assert reports is not None
    assert reports["html_file"] == "R2.htm"
    assert "short_name" in reports
    assert "long_name" in reports


@patch("edgar_client.requests.Session")
def test_find_balance_sheet(mock_session_class):
    mock_session = MagicMock()
    mock_session_class.return_value = mock_session

    mock_resp = MagicMock()
    mock_resp.content = SAMPLE_FILING_SUMMARY_XML
    mock_resp.status_code = 200
    mock_resp.raise_for_status = MagicMock()
    mock_session.get.return_value = mock_resp

    reports = find_statement_report("0001364479", "00013644792500001", "balance")
    assert reports is not None
    assert reports["html_file"] == "R3.htm"


@patch("edgar_client.requests.Session")
def test_find_cashflow(mock_session_class):
    mock_session = MagicMock()
    mock_session_class.return_value = mock_session

    mock_resp = MagicMock()
    mock_resp.content = SAMPLE_FILING_SUMMARY_XML
    mock_resp.status_code = 200
    mock_resp.raise_for_status = MagicMock()
    mock_session.get.return_value = mock_resp

    reports = find_statement_report("0001364479", "00013644792500001", "cashflow")
    assert reports is not None
    assert reports["html_file"] == "R4.htm"


@patch("edgar_client.requests.Session")
def test_fetch_statement_html_returns_string(mock_session_class):
    mock_session = MagicMock()
    mock_session_class.return_value = mock_session

    mock_resp = MagicMock()
    mock_resp.text = "<html><body><table><tr><td>Revenues</td><td>100</td></tr></table></body></html>"
    mock_resp.raise_for_status = MagicMock()
    mock_session.get.return_value = mock_resp

    html = fetch_statement_html("0001364479", "00013644792500001", "R2.htm")
    assert isinstance(html, str)
    assert "Revenues" in html
