"""Input / working sheets of the RFS workbook: MasterData, TB sheets, Debtor / Creditor Ageing,
Shareholder Register and Adj Entry RFS."""
import datetime
from openpyxl.styles import Font, Alignment
from .book_core import *          # noqa: F401,F403
from . import book_masters as BM


def _num(v):
    try:
        return float(v) if v not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def master_data(cx):
    I = cx.I
    C = I["company"]
    A = I.get("auditor", {})
    ws = cx.sheet(MD, {"A": 44, "B": 30, "C": 4, "D": 6, "E": 12, "F": 28, "G": 20, "H": 12, "I": 10})
    y1 = cx.DATE["Y1"]
    rows = [
        (1, "Company Information", None, None),
        (2, "CIN", C.get("cin", ""), None),
        (3, "Company Name", (C.get("name") or "").upper(), None),
        (4, "Formerly known as", C.get("former", ""), None),
        (5, "Registered Address", C.get("address", ""), None),
        (7, "Face value per equity share (Rs)", _num(C.get("fv", 10)), NUM0),
        (8, f"Bonus / split factor after {long_date(y1)} (1 = none)", _num(C.get("bonus", 1)) or 1, NUM0),
        (9, "Period ended (stub period)" if cx.stub else "Period ended (stub period) - not applicable", cx.stub, DMY),
        (10, "RFS Details", None, None),
        (11, "Year ended (latest)", cx.DATE["Y1"], DMY),
        (12, "Year ended (previous)", cx.DATE["Y2"], DMY),
        (13, "Year ended (earliest)", cx.DATE["Y3"], DMY),
        (14, "Opening balance sheet date", cx.opening, DMY),
        (15, "Scale", _num(C.get("scale", 100000)) or 100000, NUM0),
        (16, f"Closing stock-in-trade as at {long_date(y1)} (Rs)", _num((I.get("closing_stock") or {}).get("Y1")), NUM0),
        (18, "Gratuity (prior period) relating to the earliest year (Rs)", I.get("gratuity_prior"), NUM0),
        (19, f"Closing stock-in-trade as at {long_date(cx.stub)} (Rs)" if cx.stub else "Closing stock-in-trade as at the stub period end (Rs) - not applicable",
         _num((I.get("closing_stock") or {}).get("S")) if cx.stub else None, NUM0),
        (20, "Audit Firm Name", A.get("firm", ""), None),
        (21, "Firm Reg No", A.get("frn", ""), None),
        (22, "Peer Review Certificate No", A.get("peer", ""), None),
        (23, "Engagement Partner Name", A.get("partner", ""), None),
        (24, "Membership No", A.get("mno", ""), None),
        (25, "Date of Signing the Examination Report", A.get("sign_date") or "xx/xx/xxxx", None),
        (26, "UDIN : ", A.get("udin") or "XXXXXXXXXX", None),
        (27, "Place", C.get("place", ""), None),
    ]
    for r, lbl, val, fmt in rows:
        w(ws, f"A{r}", lbl, fB if r in (1, 10) else fN)
        if r in (1, 10):
            continue
        c = w(ws, f"B{r}", val, fBlue, fmt)
        if r in (8, 16, 18) or 20 <= r <= 26:
            c.fill = Y_IN
        if r == 19 and cx.stub:
            c.fill = Y_URG
    w(ws, "A6", "Authorised Capital (Rs) - current")
    ac = I.get("auth_capital") or []
    n_ac = max(3, len(ac))
    w(ws, "B6", f"=SUM(B31:B{30 + n_ac})", fN, NUM0)
    cx.AC_LAST = 30 + n_ac
    w(ws, "A17", C.get("tax_rate_label") or "Tax rate - Sec 115BAA (22% + 10% SC + 4% cess)")
    tr = C.get("tax_rate")
    w(ws, "B17", "=22%*1.1*1.04" if tr in (None, "", 0.25168) else float(tr), fBlue, '0.000%;\\(0.000%\\);\\-')
    if C.get("cin_note"):
        w(ws, "C2", C["cin_note"], Font(name=TN, size=9, italic=True))
    w(ws, "F15", "All amounts in Indian Rupees Lakhs, except as otherwise stated", fI)
    w(ws, "F16", f"As per audited financial statements FY {fy_label(y1)} (not carried in Tally TB)", fI)
    if I.get("gratuity_prior_note"):
        w(ws, "F18", I["gratuity_prior_note"], fI)
    if cx.stub:
        w(ws, "F19", f"As per the management (closing stock as at {long_date(cx.stub)} - not carried in the Tally TB)", fI)
    w(ws, "C22", "Peer reviewed auditor - details as furnished by the management (audit firm, FRN, peer review certificate, engagement partner).",
      Font(name=TN, size=9, italic=True))
    w(ws, "C23", "Engagement partner signing the examination report")
    # director / signatory table
    w(ws, "D2", "Director/Signatory Details", fB)
    for i, h in enumerate(["Sr. No", "DIN", "Name", "Designation", "Category", "Signatory"]):
        w(ws, f"{CL(4 + i)}3", h, fB, bd=B_TB)
    people = list(I.get("people") or [])
    while len(people) < 4:
        people.append({})
    cx.NPEOPLE = len(people)
    for k, d in enumerate(people):
        r = 4 + k
        w(ws, f"D{r}", k + 1)
        for col, key in zip("EFGHI", ["din", "name", "designation", "category", "signatory"]):
            v = d.get(key, "")
            w(ws, f"{col}{r}", (v or "").upper() if key == "name" else (v if v not in (None,) else ""), fBlue)
    # authorised capital history
    w(ws, "A29", "Authorised capital history", fB)
    for col, h in zip("ABCD", ["Date of resolution", "Increase in authorised capital (Rs)", "No. of equity shares", "Remarks"]):
        w(ws, f"{col}30", h, fB, bd=B_TB, al=Alignment(wrap_text=True))
    ws.row_dimensions[30].height = 68.25
    for k in range(n_ac):
        r = 31 + k
        if k < len(ac):
            a = ac[k]
            w(ws, f"A{r}", ddate(a.get("date")), fBlue, DMY)
            w(ws, f"B{r}", _num(a.get("amount")), fBlue, NUM0)
            w(ws, f"C{r}", f"=B{r}/$B$7", fN, NUM0)
            w(ws, f"D{r}", a.get("remarks", ""))
    r = max(35, cx.AC_LAST + 2)
    SA = I.get("stat_auditor") or {}
    w(ws, f"A{r}", "Statutory auditor (audited financial statements)", fB)
    r += 1
    for lbl, key in (("Firm", "firm"), ("Firm Reg No", "frn"), ("Partner / M.No.", "partner")):
        w(ws, f"A{r}", lbl)
        w(ws, f"B{r}", SA.get(key, ""), fBlue)
        r += 1
    for lbl, val in SA.get("reports") or []:
        w(ws, f"A{r}", lbl)
        w(ws, f"B{r}", val, fBlue)
        r += 1
    pend = [p for p in (I.get("pending") or []) if p]
    if pend:
        r += 5
        w(ws, f"A{r}", "Inputs pending", fB)
        for p in pend:
            r += 1
            w(ws, f"A{r}", p, fRed)
    return ws


def nat_of(cx, master):
    i = BM.info(master, cx.custom)
    return i[0] if i else "Dr"


def grp_of(cx, master):
    i = BM.info(master, cx.custom)
    return i[1] if i else "BS"


def tb_sheets(cx):
    """TB sheets: Tally layout (opening / transactions / closing) + Group, Masters, Net Amount (natural side +), In Lakhs.
    The earliest TB also carries the opening balance sheet (Masters (Opening), Opening Net Amount, Opening In Lakhs)."""
    I = cx.I
    order = cx.PK
    nxt = {cx.PK[i]: cx.PK[i - 1] for i in range(1, len(cx.PK))}       # next (later) period with a TB
    cx.tb_memo = {}
    for p in order:
        T = I["tbs"][p]
        ws = cx.wb.create_sheet(cx.TBN[p])
        for k, v in {"A": 66.2, "B": 13.4, "C": 12.3, "D": 13.0, "E": 13.8, "F": 7.0, "G": 44.0, "H": 15.0, "I": 12.0, "J": 8.0}.items():
            ws.column_dimensions[k].width = v
        early = p == cx.EARLY
        side_col = "M" if early else "J"
        if early:
            ws.column_dimensions["J"].width = 44
            ws.column_dimensions["K"].width = 15
            ws.column_dimensions["L"].width = 12
            ws.column_dimensions["M"].width = 8
        title = T.get("title") or [I["company"].get("name", ""), "Trial Balance", ""]
        ptxt = title[2] if len(title) > 2 and title[2] else (
            f"1-Apr-{str(cx.DATE[cx.PREV[p]].year)[2:]} to {cx.DATE[p].day}-{cx.DATE[p]:%b}-{str(cx.DATE[p].year)[2:]}")
        w(ws, "A1", title[0], fTBB)
        w(ws, "A2", title[1] if len(title) > 1 else "Trial Balance", fTBB)
        w(ws, "A3", ptxt, fTB)
        w(ws, "B4", title[0], fTBB)
        w(ws, "A5", "Particulars", fTBB)
        w(ws, "B5", ptxt, fTBB)
        for m in ("A1:C1", "A2:C2", "A3:C3", "B4:E4", "B5:E5", "C6:D6"):
            ws.merge_cells(m)
        w(ws, "B6", "Opening", fTBB)
        w(ws, "C6", "Transactions", fTBB, al=A_C)
        w(ws, "E6", "Closing", fTBB)
        w(ws, "I6", f"={cx.SCALE}", fTB, "#,##0")
        heads = ["Balance", "Debit", "Credit", "Balance", "Group", "Masters", "Net Amount", "In Lakhs"]
        heads += ["Masters (Opening)", "Opening Net Amount", "Opening In Lakhs", "Dr / Cr"] if early else ["Dr / Cr"]
        for i, h in enumerate(heads):
            w(ws, f"{CL(2 + i)}7", h, fTBB, bd=B_B)
        r = 8
        has_pla = False
        for x in T["rows"]:
            name = x.get("name", "")
            bold = x.get("group") or x.get("bold")
            w(ws, f"A{r}", name, fTBB if bold else fTB)
            op, dr, cr, cl = (_num(x.get(k)) for k in ("op", "dr", "cr", "cl"))
            for col, v in (("B", abs(op)), ("C", dr), ("D", cr), ("E", abs(cl))):
                if abs(v) > 0.0000001:
                    w(ws, f"{col}{r}", round(v, 2), fTBB if bold else fTB, "#,##0.00")
            if x.get("dup"):
                w(ws, f"G{r}", "Repeated in Tally export - not considered", fTBI)
            elif not bold and x.get("master"):
                m = x["master"]
                nat, g = nat_of(cx, m), grp_of(cx, m)
                natural = (cl >= 0) == (nat == "Dr")
                w(ws, f"F{r}", g, fTB)
                w(ws, f"G{r}", m, fTB)
                w(ws, f"H{r}", f"=E{r}" if natural or abs(cl) < 0.005 else f"=-E{r}", fTB, "#,##0.00")
                w(ws, f"I{r}", f"=H{r}/$I$6", fTB, "#,##0.00")
                if abs(cl) >= 0.005:
                    w(ws, f"{side_col}{r}", "Dr" if cl > 0 else "Cr", fTB)
                v = (abs(cl) if natural else -abs(cl)) / cx.scale_v
                cx.tbval[(p, m)] = cx.tbval.get((p, m), 0) + v
                if m == "Net profit as per books" and abs(cl) > 0.005:
                    has_pla = True
                if early:
                    om = x.get("op_master") or m
                    onat = nat_of(cx, om)
                    onatural = (op >= 0) == (onat == "Dr")
                    w(ws, f"J{r}", om, fTB)
                    w(ws, f"K{r}", f"=B{r}" if onatural or abs(op) < 0.005 else f"=-B{r}", fTB, "#,##0.00")
                    w(ws, f"L{r}", f"=K{r}/$I$6", fTB, "#,##0.00")
                    ov = (abs(op) if onatural else -abs(op)) / cx.scale_v
                    cx.tbval[("OPEN", om)] = cx.tbval.get(("OPEN", om), 0) + ov
            elif not bold and x.get("master") is None and not x.get("group"):
                w(ws, f"G{r}", "Not considered (excluded in mapping)", fTBI)
            r += 1
        last = r - 1
        tot_dr = sum(_num(x.get("dr")) for x in T["rows"] if not x.get("group") and not x.get("dup"))
        w(ws, f"A{r}", "Grand Total", fTBB)
        r += 2
        cx.tb_memo[p] = None
        if not has_pla:
            # profit for the period is not carried in a Profit & Loss A/c ledger: memorandum rows from the P&L ledgers
            if p == "S":
                stk = f"{MD}!$B$19"
            elif p == "Y1":
                stk = f"{MD}!$B$16"
            else:
                stk = f"SUMIF({q(cx.TBN[nxt[p]])}!$G:$G,\"Opening stock\",{q(cx.TBN[nxt[p]])}!$E:$E)"
            w(ws, f"A{r}", "Profit & Loss A/c - profit for the period as per books (memorandum, computed from the P&L ledgers and closing stock)", fTB)
            w(ws, f"F{r}", "BS", fTB)
            w(ws, f"G{r}", "Net profit as per books", fTB)
            w(ws, f"H{r}", f'=SUMIFS($E$8:$E${last},$F$8:$F${last},"P&L",${side_col}$8:${side_col}${last},"Cr")-SUMIFS($E$8:$E${last},$F$8:$F${last},"P&L",'
                           f'${side_col}$8:${side_col}${last},"Dr")-SUMIF($G$8:$G${last},"Opening stock",$E$8:$E${last})+{stk}', fTB, "#,##0.00")
            w(ws, f"I{r}", f"=H{r}/$I$6", fTB, "#,##0.00")
            cx.tb_memo[p] = r
            r += 1
            w(ws, f"A{r}", "Profit for the period carried to Reserves and surplus (memorandum - not yet transferred in the books)", fTB)
            w(ws, f"F{r}", "BS", fTB)
            w(ws, f"G{r}", "Reserves and surplus", fTB)
            w(ws, f"H{r}", f"=H{r - 1}", fTB, "#,##0.00")
            w(ws, f"I{r}", f"=H{r}/$I$6", fTB, "#,##0.00")
            r += 1
            cx.tbval[(p, "Net profit as per books")] = 1
            cx.tbval[(p, "Reserves and surplus")] = cx.tbval.get((p, "Reserves and surplus"), 0) + 1e-6
        w(ws, f"A{r}", "Check: TB closing balances (Dr - Cr, including opening stock) should be nil", fTBI)
        w(ws, f"H{r}", f'=ROUND(SUMIF(${side_col}$8:${side_col}${last},"Dr",$E$8:$E${last})-SUMIF(${side_col}$8:${side_col}${last},"Cr",$E$8:$E${last}),0)',
          fTBI, "#,##0")
        ws.freeze_panes = "B8"


def _ageing(cx, name, heads, rows_lbl, data, note):
    ws = cx.sheet(name, {"A": 4, "B": 40, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16})
    w(ws, "B2", f"{'Debtor' if 'Debtor' in name else 'Creditor'} Ageing schedule (amount in Rs.)", fB)
    starts = {}
    r = 3
    for p in cx.PK:
        w(ws, f"B{r}", cx.dtxt(p, "As at ", nl=False), fB)
        vals = (data or {}).get(p)
        if not vals:
            w(ws, f"H{r}", f"{long_date(cx.DATE[p])}: ageing report to be furnished by the management", fRedI)
        r += 1
        w(ws, f"B{r}", "Particulars", fB, bd=B_TB)
        for j, h in enumerate(heads):
            w(ws, f"{CL(3 + j)}{r}", h, fB, bd=B_TB)
        r += 1
        starts[p] = r
        for i, lbl in enumerate(rows_lbl):
            w(ws, f"B{r}", lbl)
            for j in range(len(heads)):
                v = vals[i][j] if vals and i < len(vals) and j < len(vals[i]) else None
                c = w(ws, f"{CL(3 + j)}{r}", _num(v) if vals else None, fBlue, ACC)
                if p == cx.PK[0] or not vals:
                    c.fill = Y_IN
            r += 1
        r += 2
    if note:
        w(ws, f"B{r + 1}", note, fI)
    return starts


def ageing_sheets(cx):
    I = cx.I
    cx.DEB = _ageing(cx, "Debtor Ageing", ["Less than 6 months", "6 months - 1 year", "1-2 years", "2-3 years", "More than 3 years"],
                     ["(i) Undisputed - considered good", "(ii) Undisputed - considered doubtful", "(iii) Disputed - considered good",
                      "(iv) Disputed - considered doubtful"], I.get("debtor_ageing"), cx.texts.get("debtor_ageing_note"))
    cx.CRE = _ageing(cx, "Creditor Ageing", ["Less than 1 year", "1-2 years", "2-3 years", "More than 3 years"],
                     ["(i) MSME", "(ii) Others", "(iii) Disputed dues - MSME", "(iv) Disputed dues - Others"], I.get("creditor_ageing"),
                     cx.texts.get("creditor_ageing_note"))


def shareholder_sheet(cx):
    SH = cx.I.get("shareholders") or []
    ws = cx.sheet("Shareholder Register", {"A": 4, "B": 40, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16})
    w(ws, "B2", "Register of members - number of equity shares held", fB)
    w(ws, "B4", "Name of shareholder", fB, bd=B_TB)
    cx.SHC = {}
    for i, p in enumerate(cx.PK):
        c = CL(3 + i)
        cx.SHC[p] = c
        w(ws, f"{c}4", f'=TEXT({cx.MDC[p]},"dd-mm-yyyy")', fB, bd=B_TB)
    cx.SHROW = {}
    r = 5
    for h in SH:
        w(ws, f"B{r}", h.get("name", ""))
        for p in cx.PK:
            w(ws, f"{cx.SHC[p]}{r}", _num(h.get(p)), fBlue, ACC0)
        cx.SHROW[h.get("name", "")] = r
        r += 1
    last = max(r - 1, 5)
    w(ws, f"B{r}", "Total", fB, bd=B_TB2)
    for p in cx.PK:
        c = cx.SHC[p]
        w(ws, f"{c}{r}", f"=SUM({c}5:{c}{last})", fB, ACC0, bd=B_TB2)
    cx.SHTOT = r
    if cx.texts.get("shareholder_note"):
        w(ws, f"B{r + 3}", cx.texts["shareholder_note"], fI)


def adj_sheet(cx):
    """Adj Entry RFS: Date | Particulars | Masters | Group | Dr | Cr (Rs. lakhs); read by the notes with SUMIFS."""
    ws = cx.sheet(ADJ, {"A": 12, "B": 70, "C": 40, "D": 8, "E": 14, "F": 14, "H": 50, "L": 14})
    w(ws, "A1", "Restatement adjustment entries (Rs. in lakhs)", fB)
    for col, h in zip("ABCDEF", ["Date", "Particulars", "Masters", "Group", "Dr", "Cr"]):
        w(ws, f"{col}2", h, fB, bd=B_TB)
    w(ws, "H2", "Working", fB)
    r = 3
    for ent in cx.entries:
        r0 = r
        for i, ln in enumerate(ent["lines"]):
            d = ent["date"]
            if i == 0:
                w(ws, f"A{r}", d if isinstance(d, str) and d.startswith("=") else ddate(d), fBlue if ent.get("manual") else fN, DMY)
            else:
                w(ws, f"A{r}", f"=A{r0}", fN, DMY)
            to = ln.get("to")
            if to is None:
                to = not _num(ln.get("dr")) and bool(ln.get("cr"))
            w(ws, f"B{r}", ("   To " if to else "") + (ln.get("particulars") or ln["master"]))
            grey = Font(name=TN, size=10, italic=True, color="808080")
            w(ws, f"C{r}", ln["master"], grey if ent.get("grey") else fN)
            w(ws, f"D{r}", grp_of(cx, ln["master"]), grey if ent.get("grey") else fN)
            for col, side in (("E", "dr"), ("F", "cr")):
                v = ln.get(side)
                if isinstance(v, str) and v.startswith("="):
                    w(ws, f"{col}{r}", v, fN, ACC)
                elif _num(v):
                    w(ws, f"{col}{r}", _num(v), fBlue, ACC)
            if i == 0 and ent.get("working"):
                w(ws, f"H{r}", ent["working"])
            for lab, val in (ent.get("calc") or {}).items() if i == 0 else []:
                pass
            r += 1
        if ent.get("narration"):
            w(ws, f"B{r}", "(" + str(ent["narration"]).strip().strip("()") + ")", fI, al=A_LW)
            ws.row_dimensions[r].height = max(15, est_height(ent["narration"], 70))
        r += 2
    cx.ADJ_END = max(150, r + 30)
    w(ws, f"B{r}", "Further adjustments, if any, may be added above row " + str(cx.ADJ_END) + " using the same Masters names as in the TB sheets.", fGreyI)
    ws.freeze_panes = "A3"
