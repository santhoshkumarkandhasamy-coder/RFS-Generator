"""Import the inputs of an existing RFS workbook in this app's format (MasterData, ageing, shareholder register, Adj Entry RFS,
notes and annexures) into the app's project. Rows are found by their labels, so earlier versions of the file also work."""
import datetime as dt
import re
import openpyxl

PER_COLS = {"S": 0, "Y1": 1, "Y2": 2, "Y3": 3}


def _rs(v):
    """'=87182/MasterData!$B$15' or '=ROUND(2592600/...,4)' -> 87182 (amount in Rs as typed); numbers as is"""
    if isinstance(v, (int, float)):
        return v
    m = re.match(r"^=\s*(?:ROUND\()?\(?([-0-9.+*/() ]+?)\)?/MasterData!\$B\$15", str(v or ""))
    if m:
        try:
            return eval(m.group(1), {"__builtins__": {}})
        except Exception:
            return None
    return None


def _iso(d):
    return d.date().isoformat() if isinstance(d, dt.datetime) else (d.isoformat() if isinstance(d, dt.date) else None)


def _find(ws, text, cols="BC", start=1):
    t = text.lower()
    for r in range(start, ws.max_row + 1):
        for c in cols:
            v = ws[f"{c}{r}"].value
            if isinstance(v, str) and v.strip().lower().startswith(t):
                return r
    return None


def import_workbook(path):
    """returns (project_update, notes_update, messages)"""
    wb = openpyxl.load_workbook(path)
    wv = openpyxl.load_workbook(path, data_only=True)
    pr, N, msg = {}, {}, []
    S = wb.sheetnames
    stub = False
    if "MasterData" in S:
        md, mv = wb["MasterData"], wv["MasterData"]
        lab = {str(md[f"A{r}"].value).strip().lower(): r for r in range(1, 60) if md[f"A{r}"].value}
        g = lambda key: next((mv[f"B{r}"].value for k, r in lab.items() if k.startswith(key)), None)
        pr.update({"cin": g("cin"), "name": g("company name"), "former": g("formerly"), "address": g("registered address"),
                   "face_value": g("face value"), "bonus_factor": g("bonus"), "tax_rate": mv["B17"].value})
        y1 = g("year ended (latest)")
        if isinstance(y1, dt.datetime):
            pr["latest"] = y1.date().isoformat()
        sb = g("period ended (stub")
        if isinstance(sb, dt.datetime):
            pr["stub_end"] = sb.date().isoformat()
            stub = True
        for r in range(1, 60):
            a = str(md[f"A{r}"].value or "")
            if a.startswith("Closing stock") and "September" not in a and isinstance(mv[f"B{r}"].value, (int, float)):
                if stub and r == 19:
                    pr["stub_closing_stock"] = mv[f"B{r}"].value
                else:
                    pr["closing_stock"] = mv[f"B{r}"].value
            if a.startswith("Closing stock") and stub and r != 16 and isinstance(mv[f"B{r}"].value, (int, float)):
                pr["stub_closing_stock"] = mv[f"B{r}"].value
        pr["auditor"] = {"firm": g("audit firm"), "frn": g("firm reg"), "peer": g("peer review"), "partner": g("engagement partner"),
                         "mno": g("membership"), "sign_date": g("date of signing"), "udin": g("udin"), "place": g("place")}
        pr["auditor"] = {k: ("" if v is None else str(v)) for k, v in pr["auditor"].items()}
        people = []
        for r in range(4, 30):
            if md[f"F{r}"].value in (None, "") and md[f"D{r}"].value in (None, ""):
                break
            if md[f"F{r}"].value:
                people.append({"din": str(md[f"E{r}"].value or ""), "name": str(md[f"F{r}"].value or "").title(), "designation": md[f"G{r}"].value or "",
                               "category": md[f"H{r}"].value or "", "signatory": md[f"I{r}"].value or ""})
        pr["directors"] = people
        r0 = next((r for r in range(25, 60) if str(md[f"A{r}"].value or "").lower().startswith("authorised capital history")), None)
        if r0:
            N["auth_history"] = [{"date": _iso(mv[f"A{r}"].value), "amount": mv[f"B{r}"].value, "remarks": md[f"D{r}"].value}
                                 for r in range(r0 + 2, r0 + 12) if isinstance(mv[f"B{r}"].value, (int, float))]
        r0 = next((r for r in range(25, 70) if str(md[f"A{r}"].value or "").lower().startswith("statutory auditor")), None)
        if r0:
            pr["stat_auditor"] = {"firm": md[f"B{r0 + 1}"].value or "", "frn": md[f"B{r0 + 2}"].value or "", "partner": md[f"B{r0 + 3}"].value or "",
                                  "reports": [[md[f"A{r}"].value, md[f"B{r}"].value] for r in range(r0 + 4, r0 + 10) if md[f"A{r}"].value]}
        msg.append(f"Company details, {len(people)} directors / signatories, auditor")
    keys = (["S"] if stub else []) + ["Y1", "Y2", "Y3"]
    dk = {"S": "dS", "Y1": "d26", "Y2": "d25", "Y3": "d24"}
    # ageing (blocks: heading 'As at ...', header row, 4 rows)
    for sh, key, nb in (("Debtor Ageing", "ageing_dr", 5), ("Creditor Ageing", "ageing_cr", 4)):
        if sh not in S:
            continue
        ws = wv[sh]
        blocks = [r for r in range(1, ws.max_row + 1) if str(wb[sh][f"B{r}"].value or "").strip().lower() == "particulars"]
        out = {}
        for k, r in zip(keys, blocks):
            vals = [[ws.cell(r + 1 + i, 3 + j).value or 0 for j in range(nb)] for i in range(4)]
            if any(any(isinstance(x, (int, float)) and x for x in row) for row in vals):
                out[dk[k]] = vals
        pr[key] = out
        msg.append(f"{sh}: {len(out)} date(s)")
    if "Shareholder Register" in S:
        ws = wv["Shareholder Register"]
        sk = {"S": "sS", "Y1": "s26", "Y2": "s25", "Y3": "s24"}
        sh = []
        prom = set()
        if "Notes_Equity & Reserves" in S:
            ne = wb["Notes_Equity & Reserves"]
            r1, r2 = _find(ne, "(b) details of shareholding of promoters"), _find(ne, "(c) details of shareholders")
            if r1 and r2:
                prom = {re.sub(r"^\d+\.\s*", "", str(ne[f"C{r}"].value)).strip().lower() for r in range(r1, r2)
                        if re.match(r"^\d+\.", str(ne[f"C{r}"].value or ""))}
        for r in range(5, ws.max_row + 1):
            n = ws[f"B{r}"].value
            if not n or str(n).strip().lower() == "total":
                break
            sh.append({"name": n, **{sk[k]: ws.cell(r, 3 + i).value or 0 for i, k in enumerate(keys)}, "promoter": str(n).strip().lower() in prom})
        pr["shareholders"] = sh
        msg.append(f"Shareholder register: {len(sh)} shareholders")
    # restatement entries (typed into tab 8 as amounts in lakhs)
    if "Adj Entry RFS" in S:
        ws, wvv = wb["Adj Entry RFS"], wv["Adj Entry RFS"]
        rows, eno, narr = [], 0, ""
        for r in range(3, ws.max_row + 1):
            c = ws[f"C{r}"].value
            a = ws[f"A{r}"].value
            if c:
                if a is not None and not str(a).startswith("=A"):
                    eno += 1
                d = wvv[f"A{r}"].value
                b = str(ws[f"B{r}"].value or "").strip()
                rows.append({"entry": eno, "date": _iso(d), "master": c, "particulars": b[3:].strip() if b.startswith("To ") else b,
                             "dr": float(wvv[f"E{r}"].value or 0), "cr": float(wvv[f"F{r}"].value or 0), "narration": ""})
            elif rows and str(ws[f"B{r}"].value or "").startswith("("):
                for x in rows:
                    if x["entry"] == eno and not x["narration"]:
                        x["narration"] = str(ws[f"B{r}"].value).strip()[1:-1]
                        break
        pr["adj_rows"] = rows
        msg.append(f"Restatement entries: {eno} (tab 8)")
    # notes inputs
    if "Notes_Liabilities" in S:
        ws = wb["Notes_Liabilities"]
        r = _find(ws, "current maturities", "C") or _find(ws, '="current maturities', "C")
        if not r:
            r = next((x for x in range(1, ws.max_row + 1) if "current maturities" in str(ws[f"C{x}"].value or "").lower()), None)
        if r:
            cm = {}
            for i, k in enumerate(keys + ["OPEN"]):
                v = ws.cell(r, 4 + i).value
                rs = _rs(v)
                cm[k] = "all" if isinstance(v, str) and "SUMIF" in v else (str(rs) if rs is not None else None)
            N["current_maturities"] = cm
    if "Borrowing Terms" in S:
        ws = wb["Borrowing Terms"]
        L = []
        for r in range(11, ws.max_row + 1):
            if not ws[f"C{r}"].value or str(ws[f"C{r}"].value).strip().lower() == "total":
                break
            led = re.findall(r'\$A:\$A,"([^"]+)"', str(ws[f"J{r}"].value or ""))
            lim = ws[f"F{r}"].value
            L.append({"lender": ws[f"C{r}"].value, "nature": ws[f"D{r}"].value, "date": _iso(ws[f"E{r}"].value),
                      "limit_rs": lim * 100000 if isinstance(lim, (int, float)) else None, "roi": ws[f"G{r}"].value, "tenor": ws[f"H{r}"].value,
                      "security": ws[f"I{r}"].value, "ledgers": ", ".join(led), "guarantee": "guarantee" in str(ws[f"I{r}"].value or "").lower()})
        N["borrowings"] = L
        msg.append(f"Borrowing terms: {len(L)} facilities")
    if "Notes_P&L" in S:
        ws = wb["Notes_P&L"]
        r0, r1 = _find(ws, "other income", "C"), None
        oi = []
        if r0:
            for r in range(r0 + 1, r0 + 30):
                f = str(ws[f"D{r}"].value or "")
                if str(ws[f"C{r}"].value or "").strip().lower() == "total":
                    break
                m = re.search(r'\$G:\$G,"([^"]+)",[^,]+\$A:\$A,"([^"]+)"', f)
                if m:
                    lbl = str(ws[f"C{r}"].value)
                    oi.append({"label": lbl, "master": m.group(1), "ledgers": m.group(2), "interest": "interest" in lbl.lower(),
                               "fd": "fixed deposit" in lbl.lower(), "ppe_sale": "sale of property" in lbl.lower()})
        if oi:
            N["oi_lines"] = oi
    if "RPT" in S:
        ws = wb["RPT"]
        r0 = _find(ws, "directors' remuneration", "C")
        rem = []
        if r0:
            for r in range(r0 + 1, r0 + 20):
                n = str(ws[f"C{r}"].value or "")
                if not n or n.lower() == "total" or n.lower().startswith("not allocated"):
                    break
                rem.append({"name": n, **{k: _rs(ws.cell(r, 4 + i).value) for i, k in enumerate(keys)}})
        N["remuneration"] = rem
        r1, r2 = _find(ws, "key managerial personnel", "C"), _find(ws, "(b) transactions", "C")
        if r1 and r2:
            nat = [ws[f"D{r}"].value for r in range(r1 + 1, r2) if ws[f"D{r}"].value and str(ws[f"C{r}"].value or "").startswith("=PROPER")]
            N["kmp_nature"] = {p["name"]: nat[i] for i, p in enumerate(pr.get("directors", [])) if i < len(nat)}
    if "RPT" in S:
        ws = wb["RPT"]
        r1, r2 = _find(ws, "entity in which the directors", "C"), _find(ws, "(b) transactions", "C")
        ents = []
        if r1 and r2:
            for r in range(r1 + 1, r2):
                n = ws[f"C{r}"].value
                if isinstance(n, str) and n.strip() and not n.startswith("="):
                    e = {"name": n.strip(), "relation": ws[f"D{r}"].value or ""}
                    for rr in range(r2, ws.max_row + 1):          # sales row and balance rows of the entity
                        if str(ws[f"C{rr}"].value or "").strip() == e["name"]:
                            vals = [ws.cell(rr, 4 + i).value for i in range(len(keys))]
                            if any(v not in (None, "") for v in vals) and "sales_S" not in e and not any("SUMIF" in str(v) for v in vals):
                                for k, v in zip(keys, vals):
                                    e["sales_" + k] = _rs(v) if _rs(v) is not None else (v if isinstance(v, str) and not v.startswith("=") else None)
                            for r3 in (rr + 1, rr + 2):
                                lab = str(ws[f"C{r3}"].value or "").lower()
                                m = re.search(r'\$A:\$A,"([^"]+)"', str(ws[f"D{r3}"].value or ""))
                                if m and "receivable" in lab:
                                    e["tr_ledger"] = m.group(1)
                                if m and "payable" in lab:
                                    e["tp_ledger"] = m.group(1)
                    ents.append(e)
        N["rpt_entities"] = ents
        rg = _find(ws, "borrowings secured by personal guarantees", "C")
        if rg and N.get("borrowings"):
            rows_g = {int(x) for x in re.findall(r"'Borrowing Terms'!J(\d+)", str(ws[f"D{rg}"].value or ""))}
            for i, b in enumerate(N["borrowings"]):
                b["guarantee"] = (11 + i) in rows_g
    if "MasterData" in S:
        md = wb["MasterData"]
        pr["cin_note"] = md["C2"].value or ""
        r0 = next((r for r in range(30, 80) if str(md[f"A{r}"].value or "").lower().startswith(("stub period (", "inputs pending"))), None)
        if r0:
            pr["pending"] = "\n".join(str(md[f"A{r}"].value) for r in range(r0 + 1, r0 + 20) if md[f"A{r}"].value)
    for sh, key in (("Debtor Ageing", "debtor_ageing_note"), ("Creditor Ageing", "creditor_ageing_note"), ("Shareholder Register", "shareholder_note")):
        if sh in S:
            ws = wb[sh]
            last_num = max((r for r in range(1, ws.max_row + 1) if any(isinstance(ws.cell(r, c).value, (int, float)) for c in range(3, 8))), default=1)
            lines = [str(ws[f"B{r}"].value) for r in range(last_num + 1, ws.max_row + 1)
                     if isinstance(ws[f"B{r}"].value, str) and len(ws[f"B{r}"].value) > 30 and not ws[f"B{r}"].value.startswith("=")]
            if lines:
                N.setdefault("texts_extra", {})[key] = "\n".join(lines)
    if "Gratuity and contingencies" in S:
        ws = wb["Gratuity and contingencies"]
        fy = [k for k in keys if k != "S"]
        G = {"assumptions": [], "dbo": {}}
        ra, rb = _find(ws, "a) actuarial assumptions", "C"), _find(ws, "b) change in the present value", "C")
        if ra and rb:
            for r in range(ra + 2, rb):
                lbl = ws[f"C{r}"].value
                vals = [ws.cell(r, 4 + PER_COLS[k]).value for k in fy]
                if lbl or any(v not in (None, "") for v in vals):
                    G["assumptions"].append({"label": lbl or "", **{k: v for k, v in zip(fy, vals)}})
        rows = {"present value of obligation as at the beginning": "open", "current service cost": "service", "interest cost": "interest",
                "benefits paid": "paid", "actuarial (gain) / loss on obligations": "actuarial", "present value of obligation as at the end": "closing"}
        if rb:
            for r in range(rb + 2, rb + 9):
                lbl = str(ws[f"C{r}"].value or "").lower()
                for k_, key in rows.items():
                    if lbl.startswith(k_) and key not in G["dbo"]:
                        G["dbo"][key] = [_rs(ws.cell(r, 4 + PER_COLS[k]).value) for k in fy]
        rd = _find(ws, "add: gratuity paid directly", "C")
        if rd:
            G["dbo"]["direct_paid"] = [_rs(ws.cell(rd, 4 + PER_COLS[k]).value) for k in fy]
        re_ = _find(ws, "experience adjustments on plan", "C")
        if re_:
            G["dbo"]["experience"] = [_rs(ws.cell(re_, 4 + PER_COLS[k]).value) for k in fy]
        rc = _find(ws, "current (included", "C")
        if rc:
            G["current"] = {k: _rs(ws.cell(rc, 4 + i).value) for i, k in enumerate(keys + ["OPEN"])}
        rs_ = _find(ws, "discount rate", "C", (_find(ws, "h) sensitivity", "C") or ws.max_row))
        if rs_:
            G["sensitivity"] = [[_rs(ws.cell(rs_ + i, 5).value), _rs(ws.cell(rs_ + i, 6).value)] for i in range(3)]
        m = re.search(r"independent actuary \(([^)]+)\)", str(ws[f"C{(_find(ws, 'the company provides for gratuity', 'C') or 1)}"].value or ""))
        if m:
            G["actuary"] = m.group(1)
        N["gratuity"] = G
        cont = {}
        for key, lbl in (("claims", "(a) claims against"), ("guarantees", "(b) guarantees issued"), ("other", "(c) other money")):
            r = _find(ws, lbl, "C")
            if r:
                cont[key] = {k: _rs(ws.cell(r, 4 + i).value) for i, k in enumerate(keys)}
        N["contingent"] = cont
        msg.append("Gratuity disclosure, contingent liabilities")
    if "Tax Computation" in S:
        ws = wb["Tax Computation"]
        T = {"other_rows": [], "permanent": []}
        r = _find(ws, "depreciation as per the income-tax act", "C")
        if r:
            T["it_dep"] = {k: (abs(_rs(ws.cell(r, 4 + i).value)) if _rs(ws.cell(r, 4 + i).value) else None) for i, k in enumerate(keys)}
        r0, r1 = _find(ws, "other differences", "C"), _find(ws, "total timing differences", "C")
        if r0 and r1:
            for r in range(r0 + 1, r1):
                lbl = str(ws[f"C{r}"].value or "")
                if lbl and "provision for gratuity" not in lbl.lower():
                    T["other_rows"].append({"label": lbl, **{k: _rs(ws.cell(r, 4 + i).value) for i, k in enumerate(keys)}})
        r = _find(ws, "assessment year", "C")
        if r:
            T["itr"] = {k: [ws.cell(r + j, 4 + i).value if not str(ws.cell(r + j, 4 + i).value or "").startswith("=") else None for j in range(7)]
                        for i, k in enumerate(keys)}
        N["tax"] = T
        msg.append("Tax computation inputs and return details")
    if "Additional Disclosures" in S:
        ws = wb["Additional Disclosures"]
        r = _find(ws, "for taxation matters", "C")
        if r:
            N["auditor_fees"] = {"tax_audit": {k: _rs(ws.cell(r, 4 + i).value) for i, k in enumerate(keys)}}
        fx = {}
        for key, lbl in (("imports", "a. cif value"), ("exports", "b. earnings in foreign"), ("expenditure", "c. expenditure in foreign")):
            r = _find(ws, lbl, "C")
            if r:
                fx[key] = {k: (_rs(ws.cell(r, 4 + i).value) if _rs(ws.cell(r, 4 + i).value) is not None else
                               (ws.cell(r, 4 + i).value if isinstance(ws.cell(r, 4 + i).value, str) else None)) for i, k in enumerate(keys)}
                if key == "expenditure":
                    fx["expenditure_label"] = ws[f"C{r}"].value
        N["foreign"] = fx
    return pr, N, msg


def import_mapping(path):
    """ledger name -> masters used in the workbook's TB sheets, as name-based overrides {"@ledger": {"dr": m, "cr": m}}"""
    from . import book_masters as BM
    wb = openpyxl.load_workbook(path)
    seen = {}
    for n in wb.sheetnames:
        if not n.startswith("TB "):
            continue
        ws = wb[n]
        for r in range(8, ws.max_row + 1):
            a, g, h = ws[f"A{r}"].value, ws[f"G{r}"].value, ws[f"H{r}"].value
            if not a or not g or not isinstance(h, str) or "memorandum" in str(a) or str(g).startswith(("Repeated", "Not considered")):
                continue
            info = BM.info(g) or ("Dr", "BS", None)
            natural = not h.startswith("=-")
            side = info[0] if natural else ("Cr" if info[0] == "Dr" else "Dr")
            seen.setdefault(str(a).strip(), {})[side] = g
            seen.setdefault(n + "@" + str(a).strip(), {})[side] = g           # year-specific (TB sheet name)
    return {"@" + k: {"dr": v.get("Dr") or v.get("Cr"), "cr": v.get("Cr") or v.get("Dr")} for k, v in seen.items()}


def detect_fa_register(path, latest, stub=None, opening=None):
    """column mapping of a fixed asset register sheet ('FA Register') in the workbook; None if not found"""
    wb = openpyxl.load_workbook(path, read_only=True)
    if "FA Register" not in wb.sheetnames:
        return None
    ws = wb["FA Register"]
    rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400), values_only=True))
    hdr = next((i for i, r in enumerate(rows[:30]) if any(isinstance(v, str) and "classification" in v.lower() for v in r)), None)
    if hdr is None:
        return None
    H = {i: str(v or "").strip() for i, v in enumerate(rows[hdr])}
    col = lambda i: openpyxl.utils.get_column_letter(i + 1)
    find = lambda *keys: next((i for i, t in H.items() if any(k in t.lower() for k in keys)), None)
    ci, ai, di = find("fa classification", "classification"), find("amount", "cost"), find("date of purchase", "purchase date", "put to use")
    dl = find("deletion")
    stub_dep = find("stub period - pro rata", "pro rata from date of acquisition")
    if None in (ci, ai, di):
        return None
    wdv = {}
    for i, t in H.items():
        m = re.search(r"wdv as on (\d{1,2})/(\d{1,2})/(\d{4})", t.lower())
        if m:
            wdv[dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))] = i
    want = {"OPEN": opening, "Y3": dt.date(latest.year - 2, latest.month, latest.day), "Y2": dt.date(latest.year - 1, latest.month, latest.day), "Y1": latest}
    if stub:
        want["S"] = stub
    per = {}
    for k, d in want.items():
        if d not in wdv:
            return None
        i = wdv[d]
        per[k] = {"wdv": col(i)}
        if k != "OPEN":
            per[k]["dep"] = col(i - 2)                 # register layout: depreciation, accumulated depreciation, WDV
    if stub and stub_dep is not None:
        per["S"]["dep_reg"], per["S"]["dep"] = per["S"]["dep"], col(stub_dep)
    first = hdr + 2
    last = max((i + 1 for i, r in enumerate(rows) if i > hdr and isinstance(r[ai], (int, float)) and r[ai]), default=first)
    classes = sorted({str(r[ci]) for r in rows[hdr + 1:last] if r[ci]})
    cmap = {}
    if "FA Restatement" in wb.sheetnames:
        fr = wb["FA Restatement"]
        for r in fr.iter_rows(min_row=5, max_row=40, max_col=2, values_only=True):
            if r[0] and r[1] and str(r[0]) in classes:
                cmap[str(r[0])] = r[1]
    M = {"sheet": "FA Register", "header_row": hdr + 1, "first_row": first, "last_row": last, "class": col(ci), "cost": col(ai), "date": col(di),
         "del": col(dl) if dl is not None else "", "periods": per}
    return M, cmap, classes


def import_opening_mapping(path):
    """opening balance sheet masters of the earliest TB ('Masters (Opening)' column J) -> {"@OPEN@ledger": {...}}"""
    from . import book_masters as BM
    wb = openpyxl.load_workbook(path)
    tbs = sorted(n for n in wb.sheetnames if n.startswith("TB "))
    if not tbs:
        return {}
    ws = wb[tbs[0]]
    if str(ws["J7"].value or "").lower().startswith("masters (opening)") is False:
        return {}
    seen = {}
    for r in range(8, ws.max_row + 1):
        a, j, k = ws[f"A{r}"].value, ws[f"J{r}"].value, ws[f"K{r}"].value
        if not a or not j or not isinstance(k, str):
            continue
        info = BM.info(j) or ("Dr", "BS", None)
        side = info[0] if not k.startswith("=-") else ("Cr" if info[0] == "Dr" else "Dr")
        seen.setdefault(str(a).strip(), {})[side] = j
    return {"@OPEN@" + k: {"dr": v.get("Dr") or v.get("Cr"), "cr": v.get("Cr") or v.get("Dr")} for k, v in seen.items()}


def import_rpt_overrides(path, keys):
    """director loans typed in the RPT note (audited-FS figures) instead of linked to the TB movements"""
    wb = openpyxl.load_workbook(path)
    if "RPT" not in wb.sheetnames:
        return []
    ws = wb["RPT"]
    out = []
    for kind, head in (("accepted", "unsecured loans accepted"), ("repaid", "unsecured loans repaid")):
        r0 = _find(ws, head, "C")
        if not r0:
            continue
        for r in range(r0 + 1, r0 + 15):
            n = str(ws[f"C{r}"].value or "")
            if not n or n.lower() == "total":
                break
            row = {"kind": kind, "director": n}
            for i, k in enumerate(keys):
                v = ws.cell(r, 4 + i).value
                if isinstance(v, str) and "ROUND(" in v and "Notes_Liabilities" not in v:
                    row[k] = _rs(v)
            if len(row) > 2:
                out.append(row)
    return out


def import_texts(path, keys):
    """paragraphs of the workbook -> the app's editable text keys (so the output reads as the workbook did)"""
    wb = openpyxl.load_workbook(path)
    T = {}

    def txt(sheet, cell):
        v = wb[sheet][cell].value if sheet in wb.sheetnames else None
        return v if isinstance(v, str) and not v.startswith("=") else None

    def after(sheet, head, col="C", n=1):
        if sheet not in wb.sheetnames:
            return None
        ws = wb[sheet]
        r = _find(ws, head, col)
        return txt(sheet, f"{col}{r + n}") if r else None
    if "SAP_4" in wb.sheetnames:
        ws = wb["SAP_4"]
        T["sap_background"] = txt("SAP_4", "C10")
        o = txt("SAP_4", "C11")
        if o:
            T["sap_objects"] = o
        pol = {"a)": "pol_basis", "b)": "pol_estimates", "c)": "pol_revenue", "d)": "pol_cash", "e)": "pol_cfs", "f)": "pol_ppe", "g)": "pol_dep",
               "h)": "pol_inv", "i)": "pol_emp", "j)": "pol_fx", "k)": "pol_tax", "l)": "pol_lease", "m)": "pol_borrow", "n)": "pol_prov", "o)": "pol_eps",
               "p)": "pol_segment", "q)": "pol_impair"}
        prev = None
        for r in range(12, ws.max_row + 1):
            v = ws[f"C{r}"].value
            if not isinstance(v, str) or v.startswith("r) Useful"):
                continue
            k = pol.get(v[:2])
            if k:
                T[k] = v
                prev = k
            elif prev == "pol_basis" and v.startswith("The Restated"):
                T["pol_compiled"] = v
            elif v.startswith("Operating Cycle"):
                T["pol_cycle"] = v
    if "SOA_5" in wb.sheetnames:
        ws = wb["SOA_5"]
        part = "pl"
        for r in range(1, ws.max_row + 1):
            v = str(ws[f"C{r}"].value or "")
            if v.lower().startswith("reconciliation of restated shareholders"):
                part = "bs"
            m = re.match(r"^([A-H])\) (.+):$", v)
            if m:
                h = m.group(2).lower()
                if part == "pl":
                    key = next((k for w_, k in (("depreciation", "dep"), ("gratuity", "grat"), ("micro", "msme"), ("current tax", "curtax"),
                                                ("deferred", "deftax")) if w_ in h), None)
                else:
                    key = next((k for w_, k in (("property", "ppe"), ("gratuity", "grat"), ("micro", "msme"), ("income tax", "taxprov"),
                                                ("deferred", "dta")) if w_ in h), None)
                body = txt("SOA_5", f"C{r + 1}")
                if key and body:
                    T[f"soa_{part}_{key}"] = body
                    T[f"soa_{part}_head_{key}"] = m.group(2)
            if v.startswith("* For the period") or (v.startswith("*") and part == "pl" and "soa_pl_star" not in T):
                T["soa_pl_star" if part == "pl" else "soa_bs_star"] = v
            elif v.startswith("*") and part == "bs":
                T["soa_bs_star"] = v
        T["soa_regroup"] = after("SOA_5", "material regrouping")
        T["soa_nonadj"] = after("SOA_5", "non-adjusting items")
    ne = "Notes_Equity & Reserves"
    if ne in wb.sheetnames:
        ws = wb[ne]
        for r in range(1, 60):
            v = ws[f"C{r}"].value
            if isinstance(v, str):
                if v.startswith("Terms / rights"):
                    T["sc_terms"] = v
                elif v.startswith("Note: There has been no change") or v.startswith("Note: The authorised"):
                    T["sc_auth_note"] = v
                elif v.startswith("As per records of the Company"):
                    T["sc_records"] = v
    if "Notes_Liabilities" in wb.sheetnames:
        ws = wb["Notes_Liabilities"]
        for r in range(1, ws.max_row + 1):
            v = ws[f"C{r}"].value
            if isinstance(v, str) and v.startswith("Loans from directors are interest free"):
                T["ltb_terms"] = v
            if isinstance(v, str) and v.startswith("Note: The above information regarding micro"):
                T["msme_note"] = v
            if isinstance(v, str) and v.startswith("=\"Overdraft") or (isinstance(v, str) and "secured by" in v and r > 40 and "stb_security" not in T
                                                                       and not v.startswith("Loans from")):
                T["stb_security"] = v
    if "Notes_PPE" in wb.sheetnames:
        ws = wb["Notes_PPE"]
        r0 = _find(ws, "notes:", "C")
        if r0:
            notes = [ws[f"C{r}"].value for r in range(r0 + 1, r0 + 12) if isinstance(ws[f"C{r}"].value, str) and ws[f"C{r}"].value.strip()
                     and not ws[f"C{r}"].value.startswith(("Checks", "FY ", "March 31, 2023:", "Period ended"))]
            T["ppe_notes"] = "\n".join(notes)
        r = _find(ws, "ageing of intangible assets under development", "C")
        if r:
            T["cwip_project"] = txt("Notes_PPE", f"C{r + 2}")
            nt = next((ws[f"C{x}"].value for x in range(r + 2, r + 20) if isinstance(ws[f"C{x}"].value, str) and ws[f"C{x}"].value.startswith("Ageing as at")), None)
            if nt:
                T["cwip_note"] = nt
    if "Notes_Assets" in wb.sheetnames:
        ws = wb["Notes_Assets"]
        for r in range(1, ws.max_row + 1):
            v = ws[f"C{r}"].value
            if isinstance(v, str):
                if v.startswith("Deferred tax has been recognised"):
                    T["dta_note"] = v
                elif v.startswith("On account of difference between depreciation"):
                    T["dta_dep_label"] = v
                elif v.startswith("On account of expenses disallowed"):
                    T["dta_grat_label"] = v
                elif v.startswith("Closing stock as at"):
                    T["inv_note"] = v
                elif v.startswith("Fixed deposits include"):
                    T["cash_note"] = v
    if "Notes_P&L" in wb.sheetnames:
        ws = wb["Notes_P&L"]
        r = _find(ws, "other income as a percentage", "C")
        if r:
            T["oi_note"] = txt("Notes_P&L", f"C{r + 1}")
    if "RPT" in wb.sheetnames:
        ws = wb["RPT"]
        T["rpt_intro"] = txt("RPT", "C10")
        r0 = _find(ws, "notes:", "C")
        if r0:
            T["rpt_notes"] = "\n".join(ws[f"C{r}"].value for r in range(r0 + 1, r0 + 12) if isinstance(ws[f"C{r}"].value, str) and ws[f"C{r}"].value.strip())
    g = "Gratuity and contingencies"
    if g in wb.sheetnames:
        ws = wb[g]
        for r in range(1, ws.max_row + 1):
            v = ws[f"C{r}"].value
            if isinstance(v, str):
                if v.startswith("Note: The audited financial statements") or (v.startswith("Note:") and r < 20):
                    T["cont_note"] = v
                elif v.startswith("Contributions to provident fund"):
                    T["eb_dc"] = v
                elif v.startswith("The Company provides for gratuity"):
                    T["eb_grat_intro"] = v
                elif v.startswith("The sensitivity analysis"):
                    T["eb_sens"] = v
                elif v.startswith("The gratuity scheme is unfunded"):
                    T["eb_plan"] = v
        r0 = _find(ws, "notes:", "C")
        if r0:
            T["eb_notes"] = "\n".join(ws[f"C{r}"].value for r in range(r0 + 1, r0 + 12) if isinstance(ws[f"C{r}"].value, str) and ws[f"C{r}"].value.strip())
    if "Tax Computation" in wb.sheetnames:
        ws = wb["Tax Computation"]
        for r in range(1, ws.max_row + 1):
            v = ws[f"C{r}"].value
            if isinstance(v, str) and v.startswith("Notes: (1)"):
                T["tax_notes"] = v
    if "Accounting ratios" in wb.sheetnames:
        ws = wb["Accounting ratios"]
        r0 = _find(ws, "notes:", "C")
        if r0:
            T["ar_notes"] = "\n".join(ws[f"C{r}"].value for r in range(r0 + 1, r0 + 15) if isinstance(ws[f"C{r}"].value, str) and ws[f"C{r}"].value.strip())
    if "Analytical Ratios" in wb.sheetnames:
        ws = wb["Analytical Ratios"]
        rk = ["cr", "de", "dscr", "roe", "itr", "trt", "tpt", "nct", "npr", "roce", "roi"]
        blocks = [r for r in range(1, ws.max_row + 1) if str(ws[f"C{r}"].value or "") == "Ratio"]
        for p, r0 in zip(keys, blocks):
            for i, k in enumerate(rk):
                v = ws[f"I{r0 + 1 + i}"].value
                if isinstance(v, str) and v.strip() not in ("", "-") and not v.startswith("="):
                    T[f"ar_reason:{p}:{k}"] = v
        for r in range(1, ws.max_row + 1):
            v = ws[f"C{r}"].value
            if isinstance(v, str) and v.startswith("Notes: (1)"):
                T["anr_notes"] = v
    if "Borrowing Terms" in wb.sheetnames:
        ws = wb["Borrowing Terms"]
        for r in range(1, ws.max_row + 1):
            v = ws[f"C{r}"].value
            if isinstance(v, str) and v.startswith("Notes: (i)"):
                T["bt_notes"] = v
    if "Additional Disclosures" in wb.sheetnames:
        ws = wb["Additional Disclosures"]
        for r in range(1, ws.max_row + 1):
            v = ws[f"C{r}"].value
            if not isinstance(v, str):
                continue
            m = re.match(r"^\(([ivx]+)\)", v)
            if m and not v.startswith(("(iii) Loans and advances", "(xiii)", "(xv)", "(xi) Utilisation")):
                T[f"add({m.group(1)})"] = v
            elif v.startswith("A. The Company has not advanced"):
                T["add(xi)A"] = v
            elif v.startswith("B. The Company has not received"):
                T["add(xi)B"] = v
            elif v.startswith("Note: Values for"):
                T["add(xv)_note"] = v
    if "CFS_3" in wb.sheetnames:
        for r in range(40, 70):
            v = wb["CFS_3"][f"B{r}"].value
            if isinstance(v, str) and v.startswith("Note: The Cash Flow"):
                T["cfs_note"] = v
    if "BS_1" in wb.sheetnames:
        for r in range(40, 70):
            v = wb["BS_1"][f"B{r}"].value
            if isinstance(v, str) and v.startswith("The above statement should be read"):
                T["bs_readwith"] = v
    return {k: v for k, v in T.items() if v}
