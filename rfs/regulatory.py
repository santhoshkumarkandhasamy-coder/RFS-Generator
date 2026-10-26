"""Regulatory reference checks at the time of preparing the RFS:

* latest amendments / circulars under the SEBI (ICDR) Regulations, 2018 (from sebi.gov.in listings);
* latest Draft Red Herring Prospectus (DRHP) / Red Herring Prospectus (RHP) filings, to use as format references;
* comparison of the restated financial information in a DRHP / RHP PDF with the annexures this app produces.

Listing pages are read live (internet needed). If a page cannot be read, the app shows the source links to open
manually; URLs can be changed in the app (Settings) if SEBI moves a page.
"""
import datetime as dt
import json
import os
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


# SEBI listing pages load older rows through this endpoint (same sid / ssid / smid as the listing URL)
LISTING_AJAX = "https://www.sebi.gov.in/sebiweb/ajax/home/getnewslistinfo.jsp"


def _listing_ids(url):
    q = dict(re.findall(r"[?&](sid|ssid|smid)=(\d+)", url or ""))
    return q.get("sid", ""), q.get("ssid", ""), q.get("smid", "")


def fetch_period(name, start, end, url=None, max_pages=25):
    """Rows of a SEBI listing dated between start and end (inclusive), paging back through older rows."""
    import requests
    url = url or SOURCES[name]
    sid, ssid, smid = _listing_ids(url)
    rows = [it for it in parse_listing(_get(url))]
    seen = {it["url"] for it in rows}
    page = 1
    while page < max_pages and (not rows or min((it["date"] for it in rows if it["date"]), default=end) >= start):
        data = {"nextValue": str(page), "next": "n", "search": "", "fromDate": start.strftime("%d-%m-%Y"),
                "toDate": end.strftime("%d-%m-%Y"), "fromYear": "", "toYear": "", "deptId": "-1", "sid": sid, "ssid": ssid,
                "smid": smid, "ssidhidden": ssid, "intmid": "-1", "sText": "", "ssText": "", "smText": "", "doDirect": str(page)}
        r = requests.post(LISTING_AJAX, data=data, timeout=25, headers={"User-Agent": "Mozilla/5.0 (RFS Generator)", "Referer": url})
        r.raise_for_status()
        new = [it for it in parse_listing(r.text) if it["url"] not in seen]
        if not new:
            break
        rows += new
        seen.update(it["url"] for it in new)
        page += 1
    return [it for it in rows if it["date"] and start <= it["date"] <= end]


def offer_documents_month(year, month, urls=None):
    """DRHP / RHP filings of one calendar month. Returns (items, errors)."""
    import calendar
    urls = urls or SOURCES
    start = dt.date(year, month, 1)
    end = dt.date(year, month, calendar.monthrange(year, month)[1])
    items, errors = [], []
    for name in ("Draft offer documents filed with SEBI (DRHP)", "Red herring documents filed with RoC (RHP)"):
        try:
            for it in fetch_period(name, start, end, urls.get(name)):
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


# ------------------------------------------------------------------ company location (registered office on the cover page)
STATES = ["Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar", "Chhattisgarh", "Goa", "Gujarat", "Haryana", "Himachal Pradesh",
          "Jharkhand", "Karnataka", "Kerala", "Madhya Pradesh", "Maharashtra", "Manipur", "Meghalaya", "Mizoram", "Nagaland", "Odisha",
          "Punjab", "Rajasthan", "Sikkim", "Tamil Nadu", "Telangana", "Tripura", "Uttar Pradesh", "Uttarakhand", "West Bengal",
          "Andaman and Nicobar Islands", "Chandigarh", "Dadra and Nagar Haveli and Daman and Diu", "Delhi", "Jammu and Kashmir",
          "Ladakh", "Lakshadweep", "Puducherry"]
_STATE_ALIASES = {"orissa": "Odisha", "pondicherry": "Puducherry", "new delhi": "Delhi", "nct of delhi": "Delhi", "uttaranchal": "Uttarakhand",
                  "tamilnadu": "Tamil Nadu", "t.n.": "Tamil Nadu", "j&k": "Jammu and Kashmir", "daman": "Dadra and Nagar Haveli and Daman and Diu",
                  "silvassa": "Dadra and Nagar Haveli and Daman and Diu"}
CITIES = {"mumbai": "Maharashtra", "thane": "Maharashtra", "navi mumbai": "Maharashtra", "pune": "Maharashtra", "nagpur": "Maharashtra",
          "nashik": "Maharashtra", "aurangabad": "Maharashtra", "chennai": "Tamil Nadu", "coimbatore": "Tamil Nadu", "madurai": "Tamil Nadu",
          "tiruppur": "Tamil Nadu", "salem": "Tamil Nadu", "trichy": "Tamil Nadu", "tiruchirappalli": "Tamil Nadu", "erode": "Tamil Nadu",
          "hosur": "Tamil Nadu", "bengaluru": "Karnataka", "bangalore": "Karnataka", "mysuru": "Karnataka", "mysore": "Karnataka",
          "mangaluru": "Karnataka", "hubli": "Karnataka", "hyderabad": "Telangana", "secunderabad": "Telangana", "visakhapatnam": "Andhra Pradesh",
          "vijayawada": "Andhra Pradesh", "kochi": "Kerala", "cochin": "Kerala", "ernakulam": "Kerala", "thiruvananthapuram": "Kerala",
          "kozhikode": "Kerala", "thrissur": "Kerala", "ahmedabad": "Gujarat", "surat": "Gujarat", "vadodara": "Gujarat", "rajkot": "Gujarat",
          "gandhinagar": "Gujarat", "jaipur": "Rajasthan", "jodhpur": "Rajasthan", "udaipur": "Rajasthan", "kolkata": "West Bengal",
          "howrah": "West Bengal", "noida": "Uttar Pradesh", "ghaziabad": "Uttar Pradesh", "lucknow": "Uttar Pradesh", "kanpur": "Uttar Pradesh",
          "gurugram": "Haryana", "gurgaon": "Haryana", "faridabad": "Haryana", "indore": "Madhya Pradesh", "bhopal": "Madhya Pradesh",
          "ludhiana": "Punjab", "mohali": "Punjab", "bhubaneswar": "Odisha", "patna": "Bihar", "ranchi": "Jharkhand", "raipur": "Chhattisgarh",
          "guwahati": "Assam", "dehradun": "Uttarakhand", "new delhi": "Delhi", "delhi": "Delhi", "panaji": "Goa", "chandigarh": "Chandigarh"}
# first two digits of the PIN code -> state (used only when the address names no state)
_PIN = {"11": "Delhi", "12": "Haryana", "13": "Haryana", "14": "Punjab", "15": "Punjab", "16": "Chandigarh", "17": "Himachal Pradesh",
        "18": "Jammu and Kashmir", "19": "Jammu and Kashmir", "20": "Uttar Pradesh", "21": "Uttar Pradesh", "22": "Uttar Pradesh",
        "23": "Uttar Pradesh", "24": "Uttarakhand", "25": "Uttar Pradesh", "26": "Uttar Pradesh", "27": "Uttar Pradesh", "28": "Uttar Pradesh",
        "30": "Rajasthan", "31": "Rajasthan", "32": "Rajasthan", "33": "Rajasthan", "34": "Rajasthan", "36": "Gujarat", "37": "Gujarat",
        "38": "Gujarat", "39": "Gujarat", "40": "Maharashtra", "41": "Maharashtra", "42": "Maharashtra", "43": "Maharashtra", "44": "Maharashtra",
        "45": "Madhya Pradesh", "46": "Madhya Pradesh", "47": "Madhya Pradesh", "48": "Madhya Pradesh", "49": "Chhattisgarh",
        "50": "Telangana", "51": "Andhra Pradesh", "52": "Andhra Pradesh", "53": "Andhra Pradesh", "56": "Karnataka", "57": "Karnataka",
        "58": "Karnataka", "59": "Karnataka", "60": "Tamil Nadu", "61": "Tamil Nadu", "62": "Tamil Nadu", "63": "Tamil Nadu", "64": "Tamil Nadu",
        "67": "Kerala", "68": "Kerala", "69": "Kerala", "70": "West Bengal", "71": "West Bengal", "72": "West Bengal", "73": "West Bengal",
        "74": "West Bengal", "75": "Odisha", "76": "Odisha", "77": "Odisha", "78": "Assam", "80": "Bihar", "81": "Bihar", "82": "Jharkhand",
        "83": "Jharkhand", "84": "Bihar", "85": "Bihar"}


def extract_location(text):
    """Registered office address -> {address, city, state, pin}. Reads the 'Registered Office' block of a DRHP / RHP cover page."""
    t = re.sub(r"[ \t]+", " ", text or "")
    m = re.search(r"registered (and corporate )?office\s*[:\-]?\s*(.{20,400}?)(?:\n\s*\n|corporate office|contact person|telephone|tel\s*[:.]|e-?mail|website|$)",
                  t, re.I | re.S)
    block = re.sub(r"\s+", " ", m.group(2)).strip(" ,;") if m else ""
    hay = (block or t[:6000]).lower()
    state = next((s_ for s_ in sorted(STATES, key=len, reverse=True) if s_.lower() in hay), None)
    if not state:
        state = next((v for k, v in _STATE_ALIASES.items() if k in hay), None)
    city = next((c.title() for c in sorted(CITIES, key=len, reverse=True) if re.search(r"\b" + re.escape(c) + r"\b", hay)), "")
    pin_m = re.search(r"\b([1-8]\d{2})\s?(\d{3})\b", block) or (None if block else re.search(r"\b([1-8]\d{2})\s?(\d{3})\b", hay))
    pin = (pin_m.group(1) + pin_m.group(2)) if pin_m else ""
    if not state and city:
        state = CITIES[city.lower()]
    if not state and pin:
        state = _PIN.get(pin[:2])
    if not city and block:   # the word just before the PIN / state is usually the city
        mm = re.search(r"([A-Za-z][A-Za-z .]{2,30}?)\s*[-,]?\s*(?:\d{3}\s?\d{3})", block)
        if mm:
            city = mm.group(1).strip(" ,.-").split(",")[-1].strip().title()
    dm = re.search(r"\b(?:district|dist\.?)\s*[:\-]?\s*([A-Za-z][A-Za-z .]{2,30}?)(?=\s*[-,]|\s*\d|$)", block, re.I)
    district = dm.group(1).strip(" .").title() if dm else ""
    return {"address": block[:300], "city": city, "district": district, "state": state or "", "pin": pin}


PIN_API = "https://api.postalpincode.in/pincode/{pin}"   # India Post PIN directory (free, no key)
_pin_cache = {}


def pin_lookup(pin):
    """Official district and state of a PIN code from India Post. Returns {} if it cannot be reached."""
    if not pin or len(pin) != 6:
        return {}
    if pin in _pin_cache:
        return _pin_cache[pin]
    out = {}
    try:
        import requests
        r = requests.get(PIN_API.format(pin=pin), timeout=15, headers={"User-Agent": "Mozilla/5.0 (RFS Generator)"})
        data = r.json()
        po = (data[0].get("PostOffice") or []) if data else []
        if po:
            out = {"district": po[0].get("District", ""), "state": po[0].get("State", "")}
    except Exception:
        pass
    _pin_cache[pin] = out
    return out


def _norm_state(s_):
    low = (s_ or "").lower()
    return next((x for x in STATES if x.lower() == low), None) or _STATE_ALIASES.get(low) or (s_ or "").title()


def add_district(loc):
    """Fills district / state from the PIN (India Post), falling back to the address text, then the city."""
    official = pin_lookup(loc.get("pin"))
    if official.get("district"):
        loc["district"] = official["district"].title()
    if official.get("state"):
        loc["state"] = _norm_state(official["state"])
    if not loc.get("district"):
        loc["district"] = loc.get("city", "")
    return loc


def _cache_path():
    d = os.path.join(os.path.expanduser("~"), ".rfs_generator")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, "offer_locations.json")


def load_location_cache():
    try:
        with open(_cache_path(), encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def save_location_cache(cache):
    try:
        with open(_cache_path(), "w", encoding="utf-8") as fh:
            json.dump(cache, fh, indent=1)
    except Exception:
        pass


def company_location(page_url, work_dir):
    """Downloads the offer document behind a SEBI filing page and reads the registered office from the first pages."""
    links = pdf_links(page_url)
    if not links:
        raise ValueError("no PDF on the filing page")
    path = os.path.join(work_dir, "loc_%08x.pdf" % (abs(hash(page_url)) % 0xFFFFFFFF))
    try:
        download(links[0], path)
        from pypdf import PdfReader
        rd = PdfReader(path, strict=False)
        text = "\n".join((rd.pages[i].extract_text() or "") for i in range(min(4, len(rd.pages))))
    finally:
        try:
            os.remove(path)
        except OSError:
            pass
    loc = add_district(extract_location(text))
    loc["pdf"] = links[0]
    return loc


def locate_all(items, work_dir, cache=None, workers=4, progress=None):
    """Adds city / state to listing items (cached by filing URL). progress(done, total) is called as filings finish."""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    cache = cache if cache is not None else load_location_cache()
    todo = [it for it in items if it["url"] not in cache]
    done = len(items) - len(todo)
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(company_location, it["url"], work_dir): it for it in todo}
        for f in as_completed(futs):
            it = futs[f]
            try:
                cache[it["url"]] = f.result()
            except Exception as e:
                cache[it["url"]] = {"address": "", "city": "", "district": "", "state": "", "pin": "", "error": str(e)[:200]}
            done += 1
            if progress:
                progress(done, len(items))
    save_location_cache(cache)
    for it in items:
        it.update({k: cache[it["url"]].get(k, "") for k in ("city", "district", "state", "pin", "address")})
    return items


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
