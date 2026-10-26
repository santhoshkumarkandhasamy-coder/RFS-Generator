"""Regulatory reference checks at the time of preparing the RFS:

* latest amendments / circulars under the SEBI (ICDR) Regulations, 2018 (from sebi.gov.in listings);
* latest Draft Red Herring Prospectus (DRHP) / Red Herring Prospectus (RHP) filings, to use as format references;
* comparison of the restated financial information in a DRHP / RHP PDF with the annexures this app produces.

Listing pages are read live (internet needed). If a page cannot be read, the app shows the source links to open
manually; URLs can be changed in the app (Settings) if SEBI moves a page.
"""
import datetime as dt
import re

FORMAT_REVIEWED_ON = dt.date(2026, 10, 26)   # date the RFS layout in builder.py was last reviewed against ICDR / recent DRHPs

SOURCES = {
    "ICDR regulations (SEBI)": "https://www.sebi.gov.in/sebiweb/home/HomeAction.do?doListing=yes&sid=1&ssid=3&smid=0",
    "SEBI circulars": "https://www.sebi.gov.in/sebiweb/home/HomeAction.do?doListing=yes&sid=1&ssid=7&smid=0",
    "Draft offer documents filed with SEBI (DRHP)": "https://www.sebi.gov.in/sebiweb/home/HomeAction.do?doListing=yes&sid=3&ssid=15&smid=10",
    "Red herring documents filed with RoC (RHP)": "https://www.sebi.gov.in/sebiweb/home/HomeAction.do?doListing=yes&sid=3&ssid=16&smid=11",
}
MANUAL_LINKS = {
    "SEBI ICDR Regulations, 2018 (latest consolidated)": "https://www.sebi.gov.in/legal/regulations.html",
    "SEBI - Filings: public issues": "https://www.sebi.gov.in/filings/public-issues.html",
    "NSE Emerge - offer documents (SME)": "https://www.nseindia.com/companies-listing/public-issues-emerge-draft-offer-documents",
    "BSE SME - offer documents": "https://www.bsesme.com/PublicIssues/DraftOfferDocument.aspx",
    "ICAI Guidance Note on Reports in Company Prospectuses (Revised 2019)": "https://www.icai.org/",
}
ICDR_KEYWORDS = re.compile(r"issue of capital and disclosure requirements|\bICDR\b|offer document|public issue|SME", re.I)

# requirements of Schedule VI / ICDR the RFS layout follows - shown as a checklist in the app
CHECKLIST = [
    ("Restated financial information for the last three financial years (plus stub period if the latest is older than six months at filing)", "Schedule VI Part A, item 11(I)"),
    ("Restated statements of assets and liabilities, profit and loss, cash flows; notes; statement of adjustments", "Schedule VI Part A, item 11(I)(A)"),
    ("Basis of preparation and significant accounting policies", "Guidance Note on Reports in Company Prospectuses"),
    ("Reconciliation of profit and equity with audited financial statements; material regrouping; non-adjusting items", "Guidance Note / Schedule VI"),
    ("Accounting ratios - EPS (basic / diluted), RONW, NAV per share, EBITDA", "Schedule VI Part A, item 11(II)"),
    ("Capitalisation statement", "Schedule VI Part A, item 11(II)"),
    ("Related party transactions", "AS 18 / Schedule VI"),
    ("Statement of tax shelter (customary in SME offer documents)", "Market practice"),
    ("Contingent liabilities and commitments", "AS 29 / Schedule III"),
    ("Schedule III additional regulatory information (ageing, analytical ratios, other disclosures)", "Schedule III Division I"),
]

# what the app produces -> keywords used to spot the same item in a DRHP / RHP
APP_ITEMS = {
    "Restated statement of assets and liabilities": ["assets and liabilities", "balance sheet"],
    "Restated statement of profit and loss": ["profit and loss", "profit & loss"],
    "Restated statement of cash flows": ["cash flow"],
    "Significant accounting policies / basis of preparation": ["accounting policies", "basis of preparation", "corporate information"],
    "Statement of adjustments / reconciliation to audited": ["adjustments", "reconciliation", "restatement adjustment"],
    "Share capital": ["share capital"],
    "Reserves and surplus / other equity": ["reserves and surplus", "other equity"],
    "Borrowings (long-term and short-term)": ["borrowing"],
    "Trade payables (MSME, ageing)": ["trade payables"],
    "Other current liabilities / provisions": ["other current liabilities", "provisions"],
    "Property, plant and equipment / intangibles": ["property, plant", "fixed assets", "intangible"],
    "Deferred tax": ["deferred tax"],
    "Loans and advances / other assets": ["loans and advances", "other current assets", "other non-current assets"],
    "Inventories": ["inventor"],
    "Trade receivables (ageing)": ["trade receivables"],
    "Cash and bank balances": ["cash and cash equivalents", "cash and bank", "bank balances"],
    "Revenue from operations / other income": ["revenue from operations", "other income"],
    "Expenses (purchases, employee, finance, depreciation, other)": ["employee benefit", "finance cost", "depreciation", "other expenses"],
    "Tax expense": ["tax expense"],
    "Earnings per share": ["earnings per share"],
    "Related party transactions": ["related party"],
    "Contingent liabilities and commitments": ["contingent liabilit"],
    "Employee benefits / gratuity": ["gratuity", "employee benefit"],
    "Tax shelter": ["tax shelter"],
    "Accounting ratios": ["accounting ratios", "return on net worth", "net asset value"],
    "Analytical ratios (Schedule III)": ["current ratio", "debt-equity", "debt equity", "analytical ratio"],
    "Capitalisation statement": ["capitalisation"],
    "Additional regulatory information (Schedule III)": ["additional regulatory", "additional disclosures", "benami", "wilful defaulter"],
}


# ------------------------------------------------------------------ live listings
def _get(url, timeout=25):
    import requests
    r = requests.get(url, timeout=timeout, headers={"User-Agent": "Mozilla/5.0 (RFS Generator)", "Accept-Language": "en"})
    r.raise_for_status()
    return r.text


_DATE = re.compile(r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.? \d{1,2}, \d{4}\b")


def _parse_date(t):
    m = _DATE.search(t or "")
    if not m:
        return None
    s = m.group(0).replace(".", "")
    for f in ("%b %d, %Y", "%B %d, %Y"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def parse_listing(html_text, base="https://www.sebi.gov.in"):
    """Returns [{date, title, url}] from a SEBI listing page (table rows with a date and a link)."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html_text, "html.parser")
    items, seen = [], set()
    for tr in soup.find_all("tr"):
        a = tr.find("a", href=True)
        if not a:
            continue
        d = _parse_date(tr.get_text(" ", strip=True))
        title = (a.get("title") or a.get_text(" ", strip=True)).strip()
        href = a["href"]
        if href.startswith("/"):
            href = base + href
        if not title or href in seen:
            continue
        seen.add(href)
        items.append({"date": d, "title": title, "url": href})
    if not items:   # fall back to any link that looks like a SEBI legal / filing page
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if re.search(r"/(legal|filings)/", href):
                if href.startswith("/"):
                    href = base + href
                if href in seen:
                    continue
                seen.add(href)
                items.append({"date": _parse_date(a.parent.get_text(" ", strip=True)), "title": a.get_text(" ", strip=True), "url": href})
    return items


def fetch(name, url=None):
    url = url or SOURCES[name]
    return parse_listing(_get(url))


def icdr_updates(urls=None, limit=15):
    """Recent ICDR-related regulations / amendments / circulars, newest first. Returns (items, errors)."""
    urls = urls or SOURCES
    items, errors = [], []
    for name in ("ICDR regulations (SEBI)", "SEBI circulars"):
        try:
            for it in fetch(name, urls.get(name)):
                if ICDR_KEYWORDS.search(it["title"]):
                    it["source"] = name
                    items.append(it)
        except Exception as e:  # network / layout change - reported in the app
            errors.append(f"{name}: {e}")
    items.sort(key=lambda x: x["date"] or dt.date.min, reverse=True)
    return items[:limit], errors


def offer_documents(urls=None, limit=20):
    """Latest DRHP / RHP filings. Returns (items, errors)."""
    urls = urls or SOURCES
    items, errors = [], []
    for name in ("Draft offer documents filed with SEBI (DRHP)", "Red herring documents filed with RoC (RHP)"):
        try:
            for it in fetch(name, urls.get(name))[:limit]:
                it["source"] = "DRHP" if "DRHP" in name else "RHP"
                items.append(it)
        except Exception as e:
            errors.append(f"{name}: {e}")
    items.sort(key=lambda x: x["date"] or dt.date.min, reverse=True)
    return items, errors


def pdf_links(page_url):
    """PDF links on a SEBI filing page (the DRHP / RHP document itself)."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(_get(page_url), "html.parser")
    out = []
    for tag in soup.find_all(["a", "iframe", "embed"]):
        src = tag.get("href") or tag.get("src") or ""
        m = re.search(r"(https?://[^\"'\s]+\.pdf)", src) or re.search(r"file=(https?://[^\"'&\s]+\.pdf)", src)
        if m:
            out.append(m.group(1))
        elif src.lower().endswith(".pdf"):
            out.append(src if src.startswith("http") else "https://www.sebi.gov.in" + src)
    return list(dict.fromkeys(out))


def download(url, path):
    import requests
    with requests.get(url, timeout=120, stream=True, headers={"User-Agent": "Mozilla/5.0 (RFS Generator)"}) as r:
        r.raise_for_status()
        with open(path, "wb") as fh:
            for chunk in r.iter_content(1 << 16):
                fh.write(chunk)
    return path


# ------------------------------------------------------------------ DRHP / RHP format comparison
_HEAD = re.compile(r"^\s*((annexure|note|schedule)\s*[-–:.]?\s*([ivxlc]+|\d+)(\.\d+)?\b.*|.*restated.*(statement|summary|information|financial).*)$", re.I)


def restated_section(pages):
    """Indexes of pages that belong to the restated financial information."""
    hits = [i for i, t in enumerate(pages) if re.search(r"restated", t, re.I) and re.search(r"annexure|note|statement of assets", t, re.I)]
    if not hits:
        return []
    # keep the densest run of pages (the section itself, not the table of contents / risk factors)
    runs, cur = [], [hits[0]]
    for h in hits[1:]:
        if h - cur[-1] <= 3:
            cur.append(h)
        else:
            runs.append(cur); cur = [h]
    runs.append(cur)
    best = max(runs, key=len)
    return list(range(best[0], best[-1] + 1))


def analyse_offer_document(pdf_path):
    """Headings found in the restated financial information of a DRHP / RHP and how they map to this app's annexures."""
    from .docreader import page_texts
    pages = page_texts(pdf_path)
    sec = restated_section(pages)
    text = "\n".join(pages[i] for i in sec)
    heads = []
    for line in text.splitlines():
        line = re.sub(r"\s+", " ", line).strip()
        if 6 <= len(line) <= 160 and _HEAD.match(line) and line not in heads:
            heads.append(line)
    low = text.lower()
    coverage = []
    for item, keys in APP_ITEMS.items():
        found = next((k for k in keys if k in low), None)
        coverage.append({"Item produced by RFS Generator": item, "Seen in DRHP / RHP": "Yes" if found else "Not found", "Matched on": found or ""})
    known = [k for keys in APP_ITEMS.values() for k in keys]
    extra = [h for h in heads if not any(k in h.lower() for k in known)]
    return {"pages": (sec[0] + 1, sec[-1] + 1) if sec else None, "total_pages": len(pages), "headings": heads,
            "coverage": coverage, "not_in_app": extra[:80]}
