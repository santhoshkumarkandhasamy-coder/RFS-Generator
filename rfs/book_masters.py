"""Masters (TB column G labels) and the note line each one feeds, in the order of the final RFS format.

Every note line is a SUMIF on its master; lines that are nil in every period are dropped (except
Schedule III rows that must always be shown). Custom masters added by the user are placed at the end
of the note they are assigned to.
"""

# note key -> list of (master, nature)
NOTE_LINES = {
    "SC": [("Equity share capital", "Cr")],
    "RS": [("Reserves and surplus", "Cr")],
    "LTB_SEC": [("Term loans from banks and NBFCs (secured)", "Cr")],
    "LTB_UNS": [],                                     # "Loan from director - <name>" masters + custom
    "OLTL": [("Rental and security deposits received", "Cr")],
    "LTP": [],                                         # gratuity (non-current) is a special line
    "STB": [("Overdraft facilities from banks (secured)", "Cr"),
            ("Bill discounting and trade finance loans (secured)", "Cr"),
            ("Purchase bill finance and channel finance (unsecured)", "Cr")],
    "TP": [("Dues to micro and small enterprises", "Cr"), ("Dues to other creditors", "Cr")],
    "OCL": [("Statutory dues payable", "Cr"), ("Salary and other employee dues payable", "Cr"),
            ("Directors' remuneration payable", "Cr"), ("Expenses and other payables", "Cr"), ("Credit card dues", "Cr"),
            ("Interest accrued on dues to MSME", "Cr"), ("Advance from customers", "Cr")],
    "STP": [("Provision for income tax (net of advance tax)", "Cr")],
    "PPE": [("Land", "Dr"), ("Buildings", "Dr"), ("Plant and machinery", "Dr"), ("Furniture and fittings", "Dr"),
            ("Vehicles", "Dr"), ("Office equipment", "Dr"), ("Electrical installations", "Dr"), ("Computers", "Dr")],
    "INTANG": [("Computer software", "Dr"), ("Other intangible assets", "Dr")],
    "CWIP": [("Capital work-in-progress", "Dr"), ("Intangible assets under development", "Dr")],
    "DTA": [("Deferred tax asset (net)", "Dr")],
    "LTLA": [("Security deposits", "Dr")],
    "TR": [("Trade receivables", "Dr")],
    "CASH": [("Cash on hand", "Dr"), ("Balances with banks in current accounts", "Dr")],
    "BANK": [("Fixed deposits with banks", "Dr")],
    "STLA": [("Advance to suppliers", "Dr"), ("Advances to employees", "Dr"), ("Other advances", "Dr")],
    "OCA": [("Balances with government authorities", "Dr"), ("Advance tax, TDS and TCS receivable", "Dr"),
            ("Prepaid expenses", "Dr"), ("Unbilled revenue", "Dr"), ("Earnest money deposits", "Dr"),
            ("Interest accrued on deposits", "Dr"), ("IPO expenses pending adjustment", "Dr"), ("Other current assets", "Dr")],
    "REV": [("Sale of products", "Cr"), ("Sale of services", "Cr"), ("Other operating revenue", "Cr")],
    "OI": [("Interest income", "Cr"), ("Other non-operating income", "Cr")],
    "PUR": [("Purchases of stock-in-trade", "Dr"), ("Direct expenses", "Dr")],
    "EMP": [("Salaries, wages and incentives", "Dr"), ("Directors' remuneration", "Dr"),
            ("Contribution to provident and other funds", "Dr"), ("Gratuity", "Dr"), ("Staff welfare expenses", "Dr")],
    "FIN": [("Interest on borrowings", "Dr"), ("Interest on delayed payments to MSME", "Dr"), ("Other borrowing costs", "Dr")],
    "DEP": [("Depreciation and amortisation", "Dr")],
    "OE": [("Electricity charges", "Dr"), ("Rent", "Dr"), ("Rates and taxes", "Dr"), ("Insurance", "Dr"),
           ("Repairs and maintenance", "Dr"), ("Travelling and conveyance", "Dr"), ("Brokerage and commission", "Dr"),
           ("Advertisement and business promotion", "Dr"), ("Legal and professional charges", "Dr"),
           ("Payment to auditors", "Dr"), ("Communication, postage and courier", "Dr"), ("Printing and stationery", "Dr"),
           ("Freight and transportation", "Dr"), ("Bad debts and balances written off", "Dr"),
           ("Loss on foreign exchange fluctuation", "Dr"), ("Miscellaneous expenses", "Dr")],
    "TAX": [("Current tax", "Dr"), ("Deferred tax", "Dr")],
}

SPECIAL = {
    "Provision for gratuity": ("Cr", "BS"),
    "Opening stock": ("Dr", "BS"),
    "Net profit as per books": ("Dr", "BS"),
    "Inter-branch accounts": ("Dr", "BS"),
    "Deferred tax liability (net)": ("Cr", "BS"),
}
DIRECTOR_PREFIX = "Loan from director - "

NOTE_NAMES = {
    "LTB_SEC": "Long-term borrowings - secured", "LTB_UNS": "Long-term borrowings - unsecured", "OLTL": "Other long-term liabilities",
    "LTP": "Long-term provisions", "STB": "Short-term borrowings", "OCL": "Other current liabilities", "STP": "Short-term provisions",
    "PPE": "Property, Plant and Equipment (class)", "INTANG": "Intangible assets (class)", "LTLA": "Long-term loans and advances",
    "CASH": "Cash and cash equivalents", "BANK": "Other bank balances", "STLA": "Short-term loans and advances",
    "OCA": "Other current assets", "REV": "Revenue from operations", "OI": "Other income", "PUR": "Purchases / direct expenses",
    "EMP": "Employee benefits expense", "FIN": "Finance costs", "OE": "Other expenses",
}

PL_NOTES = {"REV", "OI", "PUR", "EMP", "FIN", "DEP", "OE", "TAX"}


def catalogue(custom=()):
    """master -> (nature, BS/P&L, note key)"""
    out = {}
    for k, lines in NOTE_LINES.items():
        for m, nat in lines:
            out[m] = (nat, "P&L" if k in PL_NOTES else "BS", k)
    for m, (nat, g) in SPECIAL.items():
        out[m] = (nat, g, None)
    for c in custom or ():
        out[c["name"]] = (c.get("nature", "Dr"), "P&L" if c["note"] in PL_NOTES else "BS", c["note"])
    return out


def info(master, custom=()):
    if master.startswith(DIRECTOR_PREFIX):
        return ("Cr", "BS", "LTB_UNS")
    return catalogue(custom).get(master)


def note_lines(key, custom=()):
    lines = list(NOTE_LINES.get(key, []))
    lines += [(c["name"], c.get("nature", "Dr")) for c in custom or () if c["note"] == key]
    return lines
