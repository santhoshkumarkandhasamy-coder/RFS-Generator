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
            if amts:
                put("authorised_capital", max(amts), src)     # the share count and the rupee amount both appear; capital is the larger
        if re.search(r"115\s*BAA", flat, re.I):
            put("tax_rate", 0.25168, src)
        # auditor
        m = re.search(r"For\s+([A-Z][A-Za-z0-9&.,' ]{2,80}?)\s*\n?\s*Chartered\s+Accountants", text, re.I)
        if m:
            put("aud_firm", re.sub(r"\s+", " ", m.group(1)).strip(" ,."), src)
        for rx, field in ((FRN_RX, "aud_frn"), (MNO_RX, "aud_mno"), (UDIN_RX, "aud_udin"), (PEER_RX, "aud_peer")):
            m = rx.search(flat)
            if m:
                put(field, m.group(1).upper(), src)
        m = re.search(r"([A-Z][A-Za-z.]+(?:\s+[A-Z][A-Za-z.]+){0,3})\s*\n?\s*\(?\s*Partner\s*\)?", text)
        if m and not re.search(r"chartered|accountants|for\b", m.group(1), re.I):
            put("aud_partner", m.group(1).strip(), src)
        m = re.search(r"\bPlace\s*[:\-]\s*([A-Za-z ]{3,30})", flat)
        if m:
            put("aud_place", m.group(1).strip().split("  ")[0].strip(), src)
        m = re.search(r"\bDate\s*[:\-]\s*([0-9]{1,2}[./-][0-9]{1,2}[./-][0-9]{2,4}|[0-9]{1,2}(?:st|nd|rd|th)?\s+\w+,?\s+\d{4}|\w+\s+\d{1,2},\s+\d{4})", flat)
        if m:
            put("aud_date", m.group(1).strip(), src)

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
