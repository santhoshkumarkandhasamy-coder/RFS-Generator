"""RFS workbook builder in the engagement's final format (KMC layout, Annexures I - VI, 29 sheets).

build(I, out_path) writes the workbook from the input dict I collected in the app (no data is built in):

  company        cin, cin_note, name, former, address, fv, bonus, scale, tax_rate, place, incorporation_date, roc, objects, exchange
  dates          stub (optional ISO date), fy [latest, previous, earliest year ends], opening
  people         [{din, name, designation, category, signatory}]  rows 1-4 sign the statements (director, director, CFO, CS)
  auth_capital   [{date, amount (Rs), remarks}]
  auditor        {firm, frn, peer, partner, mno, sign_date, udin};  stat_auditor {firm, frn, partner, reports [[label, text]]}
  closing_stock  {"Y1": Rs, "S": Rs}  (earlier years: opening stock of the next trial balance)
  tbs            {"S"/"Y1"/"Y2"/"Y3": {"title": [..], "rows": [{name, op, dr, cr, cl (Dr +), master, op_master, group, dup}]}}
  custom_masters [{name, note, nature}]
  adjustments    [{date, lines [{master, particulars, dr, cr}], narration, working}]  (Rs. lakhs)
  restatement    {gratuity {dbo {OPEN/Y3/..: Rs}, direct_reserve}, msme {p: Rs}, tax {p: Rs | "computed"}}  automatic entries
  fa             {mode "standard", assets [...]} | {mode "mapped", source, sheet, map {...}, class_map {...}};  direct_reserve {p: Rs}
  shareholders   [{name, promoter, S, Y1, Y2, Y3}];  debtor_ageing / creditor_ageing {p: [[...]]} (Rs)
  current_maturities {p: Rs | "all"};  borrowings [{lender, nature, date, limit, roi, tenor, security, ledgers [], guarantee}]
  other_income_lines [{label, master, ledgers [], interest, fd, ppe_sale}]
  rpt, gratuity, contingent, tax, auditor_fees, foreign, loans_to_related, analytical, texts {key: text}
"""
from .book_core import Ctx, ddate, ADJ
from . import book_inputs as BI, book_notes as BN, book_fa as BF, book_statements as BS, book_annex as BA, book_auto as BAU
from . import book_masters as BM

ORDER = ["MasterData", "BS_1", "PL_2", "CFS_3", "SAP_4", "SOA_5", "Notes_Equity & Reserves", "Notes_Liabilities", "Notes_PPE", "Notes_Assets", "Notes_P&L",
         "RPT", "Gratuity and contingencies", "Tax Computation", "Accounting ratios", "Analytical Ratios", "Borrowing Terms", "Capitalisation",
         "Additional Disclosures", "TB*", "Debtor Ageing", "Creditor Ageing", "Shareholder Register", ADJ, "FA Restatement", "FA Register"]


def _num(v):
    try:
        return float(v) if v not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def default_oi_lines(I, cx):
    """one line per ledger mapped to the Other income masters (label = ledger name), from all trial balances"""
    out, seen = [], set()
    for p in cx.PK:
        for x in I["tbs"][p]["rows"]:
            m = x.get("master")
            if m in ("Interest income", "Other non-operating income") and not x.get("group") and not x.get("dup"):
                if (m, x["name"]) in seen:
                    continue
                seen.add((m, x["name"]))
                nl = x["name"].lower()
                out.append({"label": x["name"], "master": m, "ledgers": [x["name"]], "interest": m == "Interest income",
                            "fd": m == "Interest income" and ("fd" in nl or "deposit" in nl), "ppe_sale": "sale" in nl and ("asset" in nl or "car" in nl or "vehicle" in nl)})
    return out


def build(I, out_path):
    cx = Ctx(I)
    cx.scale_v = _num((I.get("company") or {}).get("scale")) or 100000.0
    # python-side movements of the director loan masters (Dr / Cr) for the movement table
    cx.tbmov = {}
    first_seen = {}
    for p in cx.PK:
        for i, x in enumerate(I["tbs"][p]["rows"]):
            m = x.get("master")
            if m and not x.get("group") and not x.get("dup"):
                cx.tbmov[(p, m, "C")] = cx.tbmov.get((p, m, "C"), 0) + _num(x.get("dr"))
                cx.tbmov[(p, m, "D")] = cx.tbmov.get((p, m, "D"), 0) + _num(x.get("cr"))
                first_seen.setdefault(m, (cx.PK.index(p), i))
    cx.dir_order = lambda m: first_seen.get(m, (9, 9))
    cx.oi_lines = I.get("other_income_lines") or default_oi_lines(I, cx)
    # borrowing terms rows (RPT refers to the guaranteed facilities)
    cx.BTC = {k: chr(ord("J") + i) for i, k in enumerate(cx.PK)}
    cx.BT_GUAR = [11 + i for i, b in enumerate(I.get("borrowings") or []) if b.get("guarantee")]

    BI.master_data(cx)
    BI.tb_sheets(cx)
    has_fa = BF.fa_register(cx)
    if has_fa:
        BF.fa_restatement(cx)
    # restatement entries: automatic (FA register, gratuity, MSME interest, current tax) + manual
    entries = []
    if has_fa and (I.get("fa") or {}).get("auto_entries", True):
        entries += BF.ppe_entries(cx)
    entries += BAU.auto_entries(cx)
    for e in I.get("adjustments") or []:
        e = dict(e)
        e["manual"] = True
        entries.append(e)
    cx.entries = entries
    for e in entries:
        for ln in e["lines"]:
            cx.adj_masters.add(ln["master"])
            info = BM.info(ln["master"], cx.custom)
            if info and info[1] == "P&L":
                cx.adj_pl_masters.add(ln["master"])
    cx.ADJ_END = max(150, 3 + sum(len(e["lines"]) + 3 for e in entries) + 30)
    BI.ageing_sheets(cx)
    BI.shareholder_sheet(cx)
    BN.notes_equity(cx)
    cx.setref("sc_begin", cx.refs["sc_end"] - 2)
    BN.notes_liabilities(cx)
    BF.notes_ppe(cx)
    BN.notes_assets(cx)
    BN.notes_pl(cx)
    BS.bs1(cx)
    BS.pl2(cx)
    BS.cfs3(cx)
    BS.sap4(cx)
    BS.soa5(cx)
    BA.rpt(cx)
    BA.gratuity(cx)
    BA.tax_comp(cx)
    BA.acc_ratios(cx)
    BA.analytical(cx)
    BA.borrowing_terms(cx)
    BA.capitalisation(cx)
    BA.additional(cx)
    BI.adj_sheet(cx)
    # sheet order as in the final RFS
    wb = cx.wb
    names = []
    for n in ORDER:
        if n == "TB*":
            names += [cx.TBN[p] for p in cx.PK]
        elif n in wb.sheetnames:
            names.append(n)
    names += [n for n in wb.sheetnames if n not in names]
    wb._sheets = [wb[n] for n in names]
    wb.active = 1
    cx.resolve()
    wb.save(out_path)
    return out_path
