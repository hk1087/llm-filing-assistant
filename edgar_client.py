import re
import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

USER_AGENT = "Harneet Kaur harneet@example.com"

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
    periods = recent["reportDate"]

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

    return [f for f in results if f["period"] > cutoff]


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

    # Strip XML namespace declarations so ElementTree tag search works regardless
    # of whether SEC FilingSummary.xml includes an xmlns attribute
    xml_content = re.sub(rb' xmlns[^"]*"[^"]*"', b"", r.content)
    root = ET.fromstring(xml_content)
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
