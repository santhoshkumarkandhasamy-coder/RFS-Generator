"""Generic rules: Tally primary group / parent group / ledger name -> Schedule III note line.

Works on Tally's standard groups (Capital Account, Loans (Liability), Current Liabilities, Fixed Assets,
Investments, Current Assets, Sales / Purchase Accounts, Direct / Indirect Incomes and Expenses ...) and common
ledger-name keywords, so it applies to any company. Returns (code if debit balance, code if credit balance);
codes are defined in mapping.LBL. Every suggestion can be changed in the app's mapping review.
"""


def _has(t, *words):
    return any(w in t for w in words)


def _fixed_asset(t):
    if _has(t, "work in progress", "work-in-progress", "cwip"):
        return "CWIP"
    if _has(t, "under development"):
        return "IAUD"
    if _has(t, "software", "intangible", "trademark", "trade mark", "patent", "copyright", "licence", "license", "website"):
        return "INT_SW"
    if "land" in t and "building" not in t:
        return "PPE_LAND"
    if _has(t, "building", "office premises", "flat", "factory shed"):
        return "PPE_BLDG"
    if _has(t, "computer", "laptop", "server", "printer"):
        return "PPE_COMP"
    if _has(t, "furniture", "fixture", "interior"):
        return "PPE_FURN"
    if _has(t, "vehicle", "motor car", " car", "bike", "two wheeler", "scooter", "truck", "lorry"):
        return "PPE_VEH"
    if _has(t, "electrical", "electric fitting", "air condition", "generator", "ups"):
        return "PPE_ELEC"
    return "PPE_OFFEQ"


def _expense(t, n):
    if n in ("depreciation",) or "depreciation" in t or "amortisation" in t or "amortization" in t:
        return "DEP"
    if _has(t, "income tax", "current tax", "provision for tax", "tax expense", "deferred tax"):
        return "TAX_DEF" if "deferred" in t else "TAX_CUR"
    if _has(" " + n + " ", "interest", " int ", "int paid", "int.") and not _has(n, "receiv", "income"):
        return "FIN_INT"
    if _has(t, "bank charges", "processing fee", "processing charges", "commitment fee", "loan charges", "finance charges", "financial charges"):
        return "FIN_OTH" if not _has(n, "interest") else "FIN_INT"
    if "director" in n and _has(n, "remuneration", "salary", "sitting fee"):
        return "EMP_DIR"
    if _has(n, "provident", "pf ", "pf-", "epf", "esi", "employer contribution", "labour welfare", "admin charges"):
        return "EMP_PF"
    if "gratuity" in n:
        return "EMP_GRAT"
    if _has(t, "staff welfare", "employee welfare", "food", "canteen", "group insurance", "health insurance", "medical", "uniform") or \
            (_has(t, "employee", "staff", "salary") and _has(n, "insurance", "education", "training", "welfare")):
        return "EMP_WEL"
    if _has(t, "salary", "salaries", "wages", "bonus", "incentive", "stipend", "leave encashment", "employee benefit"):
        return "EMP_SAL"
    if "insurance" in t:
        return "OE_INS"
    if _has(n, "audit fee", "statutory audit", "tax audit", "auditor"):
        return "OE_AUD"
    if _has(t, "professional", "legal", "consultancy", "consulting", "retainer"):
        return "OE_PROF"
    if _has(t, "electricity", "power and fuel", "power & fuel", "fuel"):
        return "OE_POWER"
    if _has(t, "rent", "lease rent", "hire charges"):
        return "OE_RENT"
    if "round off" in n or "rounding off" in n:
        return "OE_MISC"
    if _has(t, "rates and taxes", "rates & taxes", "ineligible", "registration", "filing fee", "roc fee", "stamp duty", "professional tax", "late fee"):
        return "OE_RATES"
    if _has(t, "repair", "maintenance", "amc"):
        return "OE_RM"
    if _has(t, "travel", "conveyance", "tour", "lodging", "boarding"):
        return "OE_TRAV"
    if _has(t, "brokerage", "commission"):
        return "OE_COMM"
    if _has(t, "advertis", "business promotion", "sales promotion", "marketing", "exhibition"):
        return "OE_ADV"
    if _has(t, "telephone", "internet", "mobile", "postage", "courier", "communication"):
        return "OE_TEL"
    if _has(t, "printing", "stationery"):
        return "OE_PRINT"
    if _has(t, "freight", "carriage", "transport", "loading", "unloading", "packing"):
        return "OE_FRT"
    if _has(t, "bad debt", "written off", "write off", "irrecoverable"):
        return "OE_BAD"
    if _has(t, "forex", "foreign exchange", "exchange fluctuation", "exchange loss"):
        return "OE_FOREX"
    return "OE_MISC"


def classify(p, g, n):
    """p = Tally primary group, g = immediate parent group, n = ledger name."""
    pl, gl, nl = (p or "").lower(), (g or "").lower(), (n or "").lower()
    t = " | ".join((pl, gl, nl))

    # memorandum / control items
    if _has(pl, "profit & loss a/c", "profit and loss a/c"):
        return ("XX_PLA", "XX_PLA")
    if _has(pl, "branch / divisions", "branch/divisions"):
        return ("XX_BR", "XX_BR")
    if nl == "opening stock" or _has(pl, "stock-in-hand", "stock in hand") or _has(gl, "stock-in-hand", "stock in hand"):
        return ("XX_OPSTK", "XX_OPSTK")
    if "deferred tax" in t and "expense" not in t and not _has(pl, "indirect", "direct", "expense"):
        return ("DTA", "DTL")

    # equity
    if _has(pl, "reserves", "surplus") or (_has(pl, "capital account") and _has(t, "reserve", "surplus", "premium", "retained")):
        return ("RS", "RS")
    if _has(pl, "capital account", "share capital"):
        return ("SC", "SC")

    # borrowings
    if _has(pl, "loans (liability)", "loans(liability)", "borrowing"):
        if _has(t, "bank od", "overdraft", " od ", "od a/c", "cash credit", " cc ", "cc a/c", "working capital"):
            return ("BANK", "STB_BANK")
        if _has(t, "channel finance", "purchase bill", "dealer finance"):
            return ("STLA_OTH", "STB_UNS")
        if _has(t, "short term", "short-term", "bill discount", "invoice discount", "supply chain finance", "packing credit"):
            return ("STLA_OTH", "STB_OTH")
        if _has(t, "credit card"):
            return ("STLA_OTH", "OCL_EXP")
        if _has(t, "unsecured", "director", "promoter", "relative", "shareholder"):
            return ("STLA_OTH", "LTB_UNS")
        return ("STLA_OTH", "LTB_SEC")

    # current liabilities / provisions
    if _has(pl, "current liabilities", "provisions", "liabilities"):
        if _has(t, "sundry creditor", "trade payable", "creditors"):
            msme = _has(pl + " " + gl, "msme", "micro", "small enterprise") or "msme" in nl
            return ("STLA_SUP", "TP_MSME" if msme and "non msme" not in t and "non-msme" not in t else "TP_OTH")
        if "gratuity" in t or "leave encashment" in t:
            return ("STLA_OTH", "LTP_GRAT")
        if _has(t, "provision for tax", "provision for income tax", "income tax payable", "income tax provision"):
            return ("OCA_TAX", "STP_TAX")
        if _has(t, "tds recoverable", "tds receivable", "income tax refund", "advance tax"):
            return ("OCA_TAX", "OCL_STAT")
        if _has(nl, "audit fee", "rent payable", "expenses payable", "professional fee"):
            return ("STLA_OTH", "OCL_EXP")
        if _has(t, "duties & taxes", "duties and taxes", "gst", "rcm", "reverse charge", "import duty", "customs", "tax on import", "tds", "tcs", "pf payable", "esi payable", "professional tax", "statutory"):
            return ("OCA_GOVT", "OCL_STAT")
        if _has(t, "deposit received", "rental deposit", "security deposit", "dealer deposit", "long term liabilit", "long-term liabilit"):
            return ("STLA_OTH", "OLTL")
        if _has(t, "advance from customer", "customer advance", "advance received"):
            return ("STLA_SUP", "OCL_ADV")
        if _has(t, "salary", "wages", "bonus", "employee", "staff"):
            return ("STLA_EMP", "OCL_EMP")
        if "credit card" in t:
            return ("STLA_OTH", "OCL_EXP")
        return ("STLA_OTH", "OCL_EXP")

    # fixed assets
    if _has(pl, "fixed assets", "property, plant", "tangible", "intangible"):
        c = _fixed_asset(nl)                   # ledger name first, then its groups
        if c == "PPE_OFFEQ":
            c = _fixed_asset(t)
        return (c, c)

    # investments (no separate note in the layout - fixed deposits to bank balances, others to other current assets)
    if "investment" in pl:
        return ("BANK_FD", "OCL_EXP") if _has(t, "fixed deposit", "term deposit", " fd") else ("OCA_OTH", "OCL_EXP")

    # current assets
    if _has(pl, "current assets", "misc. expenses (asset)", "misc expenses (asset)", "suspense", "assets"):
        if _has(t, "sundry debtor", "trade receivable", "debtors"):
            return ("TR", "OCL_ADV")
        if _has(t, "cash-in-hand", "cash in hand", "cash on hand", "petty cash"):
            return ("CASH", "OCL_EXP")
        if _has(t, "fixed deposit", "term deposit", "margin money", "bank guarantee") or nl.startswith("fd "):
            return ("BANK_FD", "OCL_EXP")
        if _has(t, "bank accounts", "bank account", "current account", "savings account"):
            return ("BANK", "STB_BANK")
        if _has(t, "earnest money", "emd"):
            return ("OCA_EMD", "OCL_EXP")
        if _has(t, "security deposit", "rental deposit", "rent deposit", "electricity deposit", "deposits (asset)", "long term loans", "long-term loans"):
            return ("LTLA", "OCL_EXP")
        if _has(t, "ipo", "issue expenses", "share issue"):
            return ("OCA_IPO", "OCL_EXP")
        if "unbilled" in t:
            return ("OCA_UNB", "OCL_EXP")
        if "prepaid" in t and "tax" not in t:
            return ("OCA_PRE", "OCL_EXP")
        if _has(t, "interest accrued", "accrued interest", "interest receivable"):
            return ("OCA_INT", "OCL_EXP")
        if _has(t, "salary advance", "staff advance", "employee", "staff", "imprest"):
            return ("STLA_EMP", "OCL_EMP")
        if _has(t, "tds", "tcs", "advance tax", "income tax", "prepaid tax", "self assessment"):
            return ("OCA_TAX", "OCL_STAT")
        if _has(t, "gst", "igst", "cgst", "sgst", "input tax", "electronic cash", "electronic credit", "duty", "government"):
            return ("OCA_GOVT", "OCL_STAT")
        if _has(t, "advance to supplier", "supplier advance", "vendor advance", "advance to vendor"):
            return ("STLA_SUP", "OCL_EXP")
        if _has(t, "loans & advances", "loans and advances", "advance"):
            return ("STLA_OTH", "OCL_EXP")
        return ("OCA_OTH", "OCL_EXP")

    # revenue
    if _has(pl, "sales accounts", "revenue", "sales"):
        if _has(t, "freight", "frieght", "scrap", "export incentive", "duty drawback", "other operating"):
            c = "REV_OTH"
        elif _has(t, "service", "consult", "commission income", "amc", "maintenance", "support", "subscription", "unbilled", "fees"):
            c = "REV_SERV"
        else:
            c = "REV_PROD"
        return (c, c)
    if "direct incomes" in pl and "indirect" not in pl:
        return ("REV_OTH", "REV_OTH")
    if _has(pl, "indirect incomes", "other income", "income"):
        c = "OI_INT" if "interest" in t else "OI_OTH"
        return (c, c)

    # purchases and expenses
    if _has(pl, "purchase"):
        return ("PUR", "PUR")
    if "direct expense" in pl and "indirect" not in pl:
        c = "OE_FRT" if _has(t, "freight", "carriage", "transport") else "DIR"
        return (c, c)
    if _has(pl, "indirect expense", "expense"):
        c = _expense(t, nl)
        return (c, c)

    # unknown primary group - decide on the words in the names
    if _has(t, "sales", "revenue", "income"):
        c = "OI_INT" if "interest" in t else "OI_OTH"
        return (c, c)
    if _has(t, "expense", "charges", "salary", "wages"):
        c = _expense(t, nl)
        return (c, c)
    return ("OCA_OTH", "OCL_EXP")
