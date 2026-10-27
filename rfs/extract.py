"""Company details from the uploaded documents (audited financial statements, auditor's report, MCA master data,
board reports, PAN / incorporation certificates ...). Reads plain text; every value comes with the document it was found in
so the user can check it in the Company tab."""
import collections
import datetime as dt
import re

CIN_RX = re.compile(r"\b([LU]\d{5}[A-Z]{2}\d{4}(?:PLC|PTC|OPC|NPL|GOI|SGC|FLC|GAP|GAT|ULL|ULT)\d{6})\b")
DIN_RX = re.compile(r"\bDIN\s*(?:No\.?|Number)?\s*[:\-.]?\s*(\d{8})\b", re.I)
UDIN_RX = re.compile(r"\bUDIN\s*[:\-.]?\s*(\d{8}[A-Z0-9]{10})\b", re.I)
FRN_RX = re.compile(r"\b(?:FRN|Firm\s+Reg(?:istration|n)?\.?\s*(?:No\.?|Number)?)\s*[:\-.]?\s*([0-9]{5,6}[A-Z]{0,2}(?:/[A-Z0-9]+)?)\b", re.I)
MNO_RX = re.compile(r"\b(?:M\.?\s*No\.?|Membership\s+(?:No\.?|Number))\s*[:\-.]?\s*([0-9]{5,6})\b", re.I)
PEER_RX = re.compile(r"peer\s+review\s+(?:certificate\s+)?(?:no\.?|number)?\s*[:\-.]?\s*([0-9]{4,6})\b", re.I)
_DATE = re.compile(r"\b(\d{1,2})(?:st|nd|rd|th)?[\s./-]+(January|February|March|April|May|June|July|August|September|October|November|December|"
                   r"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|\d{1,2})[\s.,/-]+(\d{4})\b", re.I)
_MONTH_FIRST = re.compile(r"\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})\b", re.I)
_DESIG = ["Managing Director", "Whole[- ]?time Director", "Wholetime Director", "Executive Director", "Non[- ]Executive Director",
          "Independent Director", "Chairman", "Chairperson", "Chief Financial Officer", "CFO", "Company Secretary", "Director"]
from .regulatory import STATES as _ST
_STATES_RX = "|".join(re.escape(x) for x in sorted(_ST, key=len, reverse=True))
_NOT_COMPANY = re.compile(r"chartered accountants|\bLLP\b|& co\b|& associates|bank\b|insurance|auditor|ministry|registrar", re.I)


def _date(t):
    m = _DATE.search(t or "")
    if m:
        d, mo, y = m.groups()
        mo = int(mo) if mo.isdigit() else dt.datetime.strptime(mo[:3].title(), "%b").month
        try:
            return dt.date(int(y), mo, int(d))
        except ValueError:
            return None
    m = _MONTH_FIRST.search(t or "")
    if m:
        try:
            return dt.datetime.strptime(" ".join(m.groups()), "%B %d %Y").date()
        except ValueError:
            return None
    return None


def _amount(s):
    s = (s or "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def _company_names(text):
    """'XYZ Private Limited' / 'XYZ Limited' style names (upper or title case)."""
    out = []
    for m in re.finditer(r"([A-Z][A-Za-z0-9&.,'()\- ]{2,80}?\s(?:PRIVATE\s+LIMITED|Private\s+Limited|LIMITED|Limited|LTD\.?|Ltd\.?))\b", text):
        n = re.sub(r"\s+", " ", m.group(1)).strip(" ,.-")
        n = re.sub(r"^(?:M/s\.?|of|the|for|to|in|and|by)\s+", "", n, flags=re.I)
        if 4 <= len(n) <= 90 and not _NOT_COMPANY.search(n) and not re.match(r"(?i)(the company|our company|a company|private limited|limited)$", n):
            out.append(n)
    return out


# label | value layouts (Excel forms, MCA master data): first non-empty cell after the label
_KV = {
    "name": r"(?:company|entity)\s+name", "cin": r"cin|corporate identity number", "aud_partner": r"partner'?s?\s+name|name of (?:the )?partner",
    "aud_frn": r"frn(?:\s+number)?|firm reg(?:istration|n)?\.?\s*(?:no\.?|number)?", "aud_firm": r"(?:audit\s+)?firm\s+name|name of (?:the )?(?:audit )?firm",
    "aud_mno": r"membership\s+(?:no\.?|number)|m\.?\s*no\.?", "address": r"registered\s+(?:office\s+)?address|registered office",
}


def _kv(text):
    out = {}
    for line in text.splitlines():
        cells = [c.strip() for c in line.split("|")]
        for i, c in enumerate(cells):
            for field, rx in _KV.items():
                if field not in out and re.fullmatch(rx, c, re.I):
                    val = next((x for x in cells[i + 1:] if x), "")
                    if val and not re.fullmatch(r"[\d.]+|0", val) or field in ("aud_mno", "aud_frn", "cin"):
                        if val and not any(re.fullmatch(r2, val, re.I) for r2 in _KV.values()):
                            out[field] = val
    return out


def _amount_rupees(v):
    """Statement amounts may be in lakhs (e.g. 110.15077): small values with decimals are read as lakhs."""
    return v * 100000 if v < 100000 and v != int(v) else v


def company_details(docs):
    """docs: {name: text}. Returns {field: (value, source document)} for what could be found."""
    found = {}

    def put(field, value, src):
        if value not in (None, "", []) and field not in found:
            found[field] = (value, src)

    # company name: the most frequent company-style name, preferring one printed next to the CIN
    names, near_cin = collections.Counter(), {}
    for src, text in docs.items():
        for n in _company_names(text):
            names[n.upper()] += 1
            near_cin.setdefault(n.upper(), (n, src))
        for m in CIN_RX.finditer(text):
            put("cin", m.group(1), src)
            around = text[max(0, m.start() - 300): m.end() + 300]
            for n in _company_names(around):
                names[n.upper()] += 3
    if names:
        best = names.most_common(1)[0][0]
        n, src = near_cin.get(best, (best, ""))
        put("name", n if not n.isupper() else n.title().replace(" Of ", " of ").replace("Ltd", "Ltd"), src)
        if "former" not in found:
            for src_, text in docs.items():
                m = re.search(r"\((?:formerly|earlier)\s+(?:known\s+as\s+)?([^)]{4,90})\)", text, re.I)
                if m:
                    put("former", "(Formerly known as " + re.sub(r"\s+", " ", m.group(1)).strip() + ")", src_)
                    break

    peer_docs = [k for k in docs if re.search(r"peer\s*review", k + " " + docs[k][:3000], re.I)]
    for src in peer_docs:                         # the RFS is signed by the peer-reviewed auditor: take the firm from its certificate
        text = docs[src]
        m = re.search(r"(?:certif\w+\s+that|issued\s+to|in favour of|M/s\.?)\s*([A-Z][A-Za-z0-9&.,' ]{2,80}?(?:& Co\.?|& Associates|LLP|Associates|Company))", text)
        if m:
            put("aud_firm", re.sub(r"\s+", " ", m.group(1)).strip(" ,."), src)
        m = FRN_RX.search(text) or re.search(r"\b(\d{6}[A-Z])\b", text)
        if m:
            put("aud_frn", m.group(1).upper(), src)
        m = PEER_RX.search(text) or re.search(r"certificate\s+(?:no\.?|number)\s*[:\-.]?\s*(\d{4,6})", text, re.I)
        if m:
            put("aud_peer", m.group(1), src)
    for src, text in docs.items():
        for field, val in _kv(text).items():
            if field.startswith("aud_") and peer_docs and src not in peer_docs:
                continue                      # statutory auditor's details - not the peer-reviewed auditor who signs the RFS
            put(field, val, src)

    for src, text in docs.items():
        flat = re.sub(r"[ \t]+", " ", text)
        # registered office
        m = re.search(r"registered\s+office\s*(?:of the company)?\s*(?:is situated at|at|:|-)?\s*(.{10,260}?\b[1-8]\d{2}\s?\d{3}\b)"
                      r"(\s*[,.-]?\s*(?:" + _STATES_RX + r"))?(\s*[,.-]?\s*India)?", flat, re.I | re.S)
        if m:
            put("address", re.sub(r"\s+", " ", "".join(g or "" for g in m.groups())).strip(" ,.:-"), src)
        # incorporation
        m = re.search(r"incorporated\s+(?:on|as)?\s*(.{0,160}?)(?:under the (?:provisions of the )?companies act,?\s*(\d{4}))", flat, re.I | re.S)
        if m:
            d = _date(m.group(0))
            put("incorporated", (d.isoformat() if d else "") + (f" under the Companies Act, {m.group(2)}" if m.group(2) else ""), src)
        # year end: latest 'year ended 31 March YYYY'
        for m in re.finditer(r"(?:year|period)\s+ended\s+(?:on\s+)?(.{0,25})", flat, re.I):
            d = _date(m.group(1))
            if d and d.month == 3 and d.day == 31:
                cur = found.get("latest")
                if not cur or d > cur[0]:
                    found["latest"] = (d, src)
        # authorised capital and face value
        m = re.search(r"authori[sz]ed\s+(?:share\s+)?capital(.{0,260})", flat, re.I | re.S)
        if m:
            seg = m.group(1)
            fv = re.search(r"(?:Rs\.?|₹|INR)\s*(\d{1,4}(?:\.\d+)?)\s*(?:/-)?\s*each", seg, re.I)
            if fv:
                put("face_value", float(fv.group(1)), src)
            amts = [a for a in (_amount(x) for x in re.findall(r"\d{1,3}(?:,\d{2,3})+(?:\.\d+)?|\d{6,}(?:\.\d+)?", seg)) if a]
            if amts and (not found.get("authorised_capital") or max(amts) > found["authorised_capital"][0]):
                found["authorised_capital"] = (max(amts), src)   # highest across documents: authorised capital only increases
        # closing stock (inventories) of the latest year: 'Inventories | note | current | previous'
        for line in text.splitlines():
            cells = [c.strip() for c in line.split("|")]
            if any(re.fullmatch(r"inventor(?:y|ies)|stock[- ]in[- ]trade", c, re.I) for c in cells):
                nums = [_amount(c) for c in cells if re.fullmatch(r"-?[\d,]+(?:\.\d+)?", c)]
                nums = [n for n in nums if n is not None]
                if nums and float(nums[0]).is_integer() and nums[0] < 100 and len(nums) > 1:
                    nums = nums[1:]                       # drop the note number
                if nums and nums[0] > 0:
                    put("closing_stock", round(_amount_rupees(nums[0]), 2), src)
                    break
        if re.search(r"115\s*BAA", flat, re.I):
            put("tax_rate", 0.25168, src)
        # auditor
        # auditor's report / financials name the statutory auditor; the RFS examination report is signed by the peer-reviewed
        # auditor, so firm / FRN come from a peer review certificate when one is uploaded, and the examination report's
        # date and UDIN are never taken from statutory audit reports
        if not peer_docs:
            m = re.search(r"For\s+([A-Z][A-Za-z0-9&.,' ]{2,80}?)\s*\n?\s*Chartered\s+Accountants", text, re.I)
            if m:
                put("aud_firm", re.sub(r"\s+", " ", m.group(1)).strip(" ,."), src)
            m = FRN_RX.search(flat)
            if m:
                put("aud_frn", m.group(1).upper(), src)
        statutory_only = bool(peer_docs) and src not in peer_docs
        for rx, field in ((MNO_RX, "aud_mno"), (PEER_RX, "aud_peer")):
            m = rx.search(flat)
            if m and not statutory_only:
                put(field, m.group(1).upper(), src)
        m = None if statutory_only else re.search(r"([A-Z][A-Za-z.]+(?:\s+[A-Z][A-Za-z.]+){0,3})\s*\n?\s*\(?\s*Partner\s*\)?", text)
        if m and not re.search(r"chartered|accountants|for\b|managing|review|designated|name", m.group(1), re.I):
            put("aud_partner", m.group(1).strip(), src)
        m = re.search(r"\bPlace\s*[:\-]\s*([A-Za-z ]{3,30})", flat)
        if m:
            put("aud_place", m.group(1).strip().split("  ")[0].strip(), src)

    if "cin" in found and "name" in found and "PTC" in found["cin"][0] and not re.search(r"private", found["name"][0], re.I):
        put("former", "(Formerly known as " + re.sub(r"\s+(Limited|Ltd\.?)$", " Private Limited", found["name"][0], flags=re.I) + ")", found["cin"][1])

    # directors: a name next to a DIN, with the designation printed near it
    dirs, seen = [], set()
    for src, text in docs.items():
        for m in DIN_RX.finditer(text):
            din = m.group(1)
            if din in seen:
                continue
            before = text[max(0, m.start() - 160): m.start()]
            prev_din = list(DIN_RX.finditer(before))                 # do not look past the previous director's DIN
            if prev_din:
                before = before[prev_din[-1].end():]
            line_end = text.find("\n", m.end())
            after = text[m.end(): line_end if line_end != -1 else m.end() + 80]   # rest of the DIN line only
            lines = [ln.strip(" ,:()-") for ln in before.splitlines() if ln.strip(" ,:()-")]
            name, k_name = "", None
            for k in range(len(lines) - 1, -1, -1):
                ln2 = re.sub("|".join(_DESIG), "", lines[k], flags=re.I).strip(" ,:()-")
                if re.fullmatch(r"(?:Mr\.?|Mrs\.?|Ms\.?|Dr\.?)?\s*[A-Z][A-Za-z.]+(?:\s+[A-Z][A-Za-z.]*){0,4}", ln2) and not re.search(r"sd/-|for and on behalf|board|limited|place|date", ln2, re.I):
                    name, k_name = re.sub(r"^(?:Mr\.?|Mrs\.?|Ms\.?|Dr\.?)\s*", "", ln2), k
                    break
            near = " ".join(lines[k_name:] if k_name is not None else lines[-2:]) + " " + after   # name line onwards + DIN line
            desig = next((re.search(d, near, re.I).group(0).title() for d in _DESIG if re.search(d, near, re.I)), "Director")
            if name:
                seen.add(din)
                dirs.append({"din": din, "name": name, "designation": desig.replace("Wholetime", "Whole-time"), "category": "Promoter",
                             "signatory": "Yes" if len(dirs) < 2 else "No", "_src": src})
    if dirs:
        found["directors"] = ([{k: v for k, v in d.items() if k != "_src"} for d in dirs], ", ".join(sorted({d["_src"] for d in dirs})))
    return found
