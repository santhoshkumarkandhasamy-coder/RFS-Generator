"""Ledger -> Masters (note line) mapping.

Masters are the exact note-line texts used in the TB sheets (column G) and picked up by the
notes with SUMIF. Suggestions come from Tally group names / keywords (classify.py); the user can
override any ledger in the app, and overrides are saved as JSON for reuse.
"""
import re
from .classify import classify
from . import book_masters as BM

LBL = {
    "SC": ("Equity share capital", "Cr", "BS"), "RS": ("Reserves and surplus", "Cr", "BS"),
    "LTB_SEC": ("Term loans from banks and NBFCs (secured)", "Cr", "BS"),
    "DTL": ("Deferred tax liability (net)", "Cr", "BS"), "OLTL": ("Rental and security deposits received", "Cr", "BS"),
    "LTP_GRAT": ("Provision for gratuity", "Cr", "BS"),
    "STB_BANK": ("Overdraft facilities from banks (secured)", "Cr", "BS"),
    "STB_OTH": ("Bill discounting and trade finance loans (secured)", "Cr", "BS"),
    "STB_UNS": ("Purchase bill finance and channel finance (unsecured)", "Cr", "BS"),
    "TP_MSME": ("Dues to micro and small enterprises", "Cr", "BS"), "TP_OTH": ("Dues to other creditors", "Cr", "BS"),
    "OCL_STAT": ("Statutory dues payable", "Cr", "BS"), "OCL_EMP": ("Salary and other employee dues payable", "Cr", "BS"),
    "OCL_EXP": ("Expenses and other payables", "Cr", "BS"), "OCL_ADV": ("Advance from customers", "Cr", "BS"),
    "STP_TAX": ("Provision for income tax (net of advance tax)", "Cr", "BS"),
    "PPE_LAND": ("Land", "Dr", "BS"), "PPE_BLDG": ("Buildings", "Dr", "BS"), "PPE_FURN": ("Furniture and fittings", "Dr", "BS"),
    "PPE_VEH": ("Vehicles", "Dr", "BS"), "PPE_OFFEQ": ("Office equipment", "Dr", "BS"), "PPE_ELEC": ("Electrical installations", "Dr", "BS"),
    "PPE_COMP": ("Computers", "Dr", "BS"), "INT_SW": ("Computer software", "Dr", "BS"), "CWIP": ("Capital work-in-progress", "Dr", "BS"),
    "IAUD": ("Intangible assets under development", "Dr", "BS"), "DTA": ("Deferred tax asset (net)", "Dr", "BS"),
    "LTLA": ("Security deposits", "Dr", "BS"), "TR": ("Trade receivables", "Dr", "BS"),
    "CASH": ("Cash on hand", "Dr", "BS"), "BANK": ("Balances with banks in current accounts", "Dr", "BS"),
    "BANK_FD": ("Fixed deposits with banks", "Dr", "BS"),
    "STLA_SUP": ("Advance to suppliers", "Dr", "BS"), "STLA_EMP": ("Advances to employees", "Dr", "BS"), "STLA_OTH": ("Other advances", "Dr", "BS"),
    "OCA_GOVT": ("Balances with government authorities", "Dr", "BS"), "OCA_TAX": ("Advance tax, TDS and TCS receivable", "Dr", "BS"),
    "OCA_PRE": ("Prepaid expenses", "Dr", "BS"), "OCA_UNB": ("Unbilled revenue", "Dr", "BS"), "OCA_EMD": ("Earnest money deposits", "Dr", "BS"),
    "OCA_INT": ("Interest accrued on deposits", "Dr", "BS"), "OCA_IPO": ("IPO expenses pending adjustment", "Dr", "BS"),
    "OCA_OTH": ("Other current assets", "Dr", "BS"),
    "XX_OPSTK": ("Opening stock", "Dr", "BS"), "XX_PLA": ("Net profit as per books", "Dr", "BS"), "XX_BR": ("Inter-branch accounts", "Dr", "BS"),
    "REV_PROD": ("Sale of products", "Cr", "P&L"), "REV_SERV": ("Sale of services", "Cr", "P&L"), "REV_OTH": ("Other operating revenue", "Cr", "P&L"),
    "OI_INT": ("Interest income", "Cr", "P&L"), "OI_OTH": ("Other non-operating income", "Cr", "P&L"),
    "PUR": ("Purchases of stock-in-trade", "Dr", "P&L"), "DIR": ("Direct expenses", "Dr", "P&L"),
    "EMP_SAL": ("Salaries, wages and incentives", "Dr", "P&L"), "EMP_DIR": ("Directors' remuneration", "Dr", "P&L"),
    "EMP_PF": ("Contribution to provident and other funds", "Dr", "P&L"), "EMP_GRAT": ("Gratuity", "Dr", "P&L"),
    "EMP_WEL": ("Staff welfare expenses", "Dr", "P&L"),
    "FIN_INT": ("Interest on borrowings", "Dr", "P&L"), "FIN_OTH": ("Other borrowing costs", "Dr", "P&L"),
    "DEP": ("Depreciation and amortisation", "Dr", "P&L"),
    "OE_POWER": ("Electricity charges", "Dr", "P&L"), "OE_RENT": ("Rent", "Dr", "P&L"), "OE_RATES": ("Rates and taxes", "Dr", "P&L"),
    "OE_INS": ("Insurance", "Dr", "P&L"), "OE_RM": ("Repairs and maintenance", "Dr", "P&L"), "OE_TRAV": ("Travelling and conveyance", "Dr", "P&L"),
    "OE_COMM": ("Brokerage and commission", "Dr", "P&L"), "OE_ADV": ("Advertisement and business promotion", "Dr", "P&L"),
    "OE_PROF": ("Legal and professional charges", "Dr", "P&L"), "OE_AUD": ("Payment to auditors", "Dr", "P&L"),
    "OE_TEL": ("Communication, postage and courier", "Dr", "P&L"), "OE_PRINT": ("Printing and stationery", "Dr", "P&L"),
    "OE_FRT": ("Freight and transportation", "Dr", "P&L"), "OE_BAD": ("Bad debts and balances written off", "Dr", "P&L"),
    "OE_FOREX": ("Loss on foreign exchange fluctuation", "Dr", "P&L"), "OE_MISC": ("Miscellaneous expenses", "Dr", "P&L"),
    "TAX_CUR": ("Current tax", "Dr", "P&L"), "TAX_DEF": ("Deferred tax", "Dr", "P&L"),
}

MASTER_INFO = {v[0]: (v[1], v[2]) for v in LBL.values()}
MASTER_INFO.update({
    "Credit card dues": ("Cr", "BS"),
    "Interest accrued on dues to MSME": ("Cr", "BS"),
    "Interest on delayed payments to MSME": ("Dr", "P&L"),
    "Directors' remuneration payable": ("Cr", "BS"),
})
for _m, (_nat, _grp, _note) in BM.catalogue().items():      # every master of the final RFS format
    MASTER_INFO.setdefault(_m, (_nat, _grp))
CUSTOM = []                         # custom masters [{name, note, nature}] set by the app
DIRECTOR_PREFIX = "Loan from director - "
EXCLUDE = "Not considered"
CRED_PRIM = {"Capital Account", "Loans (Liability)", "Current Liabilities", "Sales Accounts", "Indirect Incomes"}


def all_labels(extra=()):
    return sorted(set(MASTER_INFO) | {c["name"] for c in CUSTOM} | set(extra)) + [EXCLUDE]


def label_info(label):
    if label.startswith(DIRECTOR_PREFIX):
        return ("Cr", "BS")
    i = BM.info(label, CUSTOM)
    if i:
        return i[0], i[1]
    return MASTER_INFO[label]


def key(x):
    return f"{x['primary']}|{x['parent']}|{x['name']}"


def _clean(name):
    return re.sub(r"\s*\((director|dir)\)\s*$", "", name, flags=re.I).strip()


def suggest(primary, parent, name):
    """returns (master if debit balance, master if credit balance)"""
    dr, cr = classify(primary, parent, name)
    d = LBL[dr][0] if dr in LBL else DIRECTOR_PREFIX + _clean(name)
    c = LBL[cr][0] if cr in LBL else DIRECTOR_PREFIX + _clean(name)
    nl = name.lower()
    if cr == "LTB_UNS":
        c = DIRECTOR_PREFIX + _clean(name)
    if "msme" in nl and "interest" in nl:
        if primary == "Indirect Expenses":
            d = c = "Interest on delayed payments to MSME"
        elif primary == "Current Liabilities":
            c = "Interest accrued on dues to MSME"
    if "director" in nl and "remuneration" in nl and "payable" in nl:
        c = "Directors' remuneration payable"
    if "credit card" in nl:
        d = c = "Credit card dues"
    for side in ("d", "c"):                  # director loans: full name of the director from the Company tab
        v = d if side == "d" else c
        if v.startswith(DIRECTOR_PREFIX):
            v = DIRECTOR_PREFIX + person_name(v[len(DIRECTOR_PREFIX):])
            if side == "d":
                d = v
            else:
                c = v
    return d, c


PEOPLE = []                         # director names from the Company tab (set by the app)


def person_name(short):
    """'Lenin' -> 'C K Lenin Christopher' when a director's name contains it"""
    toks = [t for t in re.split(r"[^a-z]+", short.lower()) if len(t) > 2 and t not in ("director", "loan", "dir", "mr", "mrs", "ms")]
    for p in PEOPLE:
        pl = p.lower()
        if toks and all(t in pl for t in toks):
            return p
    return short


def master_for(x, which, overrides=None):
    """which = 3 (closing) or 0 (opening). Returns (label, nature, group)."""
    v = x["v"][which]
    other = x["v"][0 if which == 3 else 3]
    s = v if abs(v) > 0.001 else (other if abs(other) > 0.001 else (-1 if x["primary"] in CRED_PRIM else 1))
    ov = (overrides or {}).get(key(x)) or (overrides or {}).get("@" + x["name"].strip())
    d, c = (ov["dr"], ov["cr"]) if ov else suggest(x["primary"], x["parent"], x["name"])
    lbl = d if s >= 0 else c
    if lbl == EXCLUDE:
        return None
    nat, grp = label_info(lbl)
    return lbl, nat, grp
