"""Automatic restatement entries (gratuity as per actuarial valuation, MSME interest, current tax), written to
'Adj Entry RFS' with formulas so the amounts stay linked to the TB sheets."""
from .book_core import *          # noqa: F401,F403


def _n(v):
    try:
        return float(v) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


def L(master, amt, part=None, to=False):
    return {"master": master, "particulars": part or master, "dr": f"=MAX({amt},0)", "cr": f"=MAX(-({amt}),0)", "to": to}


def auto_entries(cx):
    RST = cx.I.get("restatement") or {}
    S, T = cx.SCALE, cx.TAX
    chron = list(reversed(cx.PK))
    E = []
    book = lambda m, p: cx.tb(p, qs(m))
    # ---- gratuity: provision restated to the actuarial obligation at every balance sheet date
    g = RST.get("gratuity") or {}
    dbo = g.get("dbo") or {}
    if dbo:
        PG = "Provision for gratuity"
        if _n(dbo.get("OPEN")) is not None:
            x0 = f"({_n(dbo['OPEN']):g}/{S}-{book(PG, 'OPEN')})"
            E.append({"date": cx.opening, "lines": [L("Reserves and surplus", x0), L(PG, f"-{x0}", to=True)],
                      "narration": f"Being gratuity liability relating to periods up to {long_date(cx.opening)}, as per actuarial valuation, recognised in the "
                                   "opening balance sheet"})
            E.append({"date": cx.opening, "lines": [L("Deferred tax asset (net)", f"ROUND({x0}*{T},2)"), L("Reserves and surplus", f"-ROUND({x0}*{T},2)", to=True)],
                      "narration": "Being deferred tax asset on the above opening gratuity liability"})
        prev = "OPEN"
        direct = g.get("direct_reserve") or {}
        for p in chron:
            if _n(dbo.get(p)) is None:
                prev = p
                continue
            pd_ = _n(dbo.get(prev)) if prev in dbo else None
            if pd_ is None:
                prev = p
                continue
            d = f"N({_n(direct.get(p)) or 0:g})/{S}"
            dlt = f"(({_n(dbo[p]):g}-{pd_:g})/{S}-({book(PG, p)}-{book(PG, prev)}))"
            E.append({"date": cx.DATE[p], "lines": [L("Gratuity", f"{dlt}+{d}", f"Gratuity (Employee benefits expense - {_tag(cx, p)})"),
                                                    L(PG, f"-{dlt}", to=True), L("Reserves and surplus", f"-{d}", to=True)],
                      "narration": f"Being gratuity expense for the {cx.per_word(p)} ended {long_date(cx.DATE[p])} restated as per actuarial valuation"})
            E.append({"date": cx.DATE[p], "lines": [L("Deferred tax asset (net)", f"ROUND(({dlt})*{T},2)"),
                                                    L("Deferred tax", f"-ROUND(({dlt})*{T},2)", f"Deferred tax (P&L - {_tag(cx, p)})", True)],
                      "narration": "Being deferred tax on the change in the gratuity provision"})
            prev = p
    # ---- MSME interest (Section 16) - not deductible (Section 23): no deferred tax
    MI, MA = "Interest on delayed payments to MSME", "Interest accrued on dues to MSME"
    for p in chron:
        v = _n((RST.get("msme") or {}).get(p))
        if v is None:
            continue
        x = f"({v:g}/{S}-{book(MI, p)})"
        E.append({"date": cx.DATE[p], "lines": [L(MI, x, f"{MI} (P&L - {_tag(cx, p)})"), L(MA, f"-{x}", to=True)],
                  "narration": f"Being interest under Section 16 of the MSMED Act on dues to micro and small enterprises for the {cx.per_word(p)} ended "
                               f"{long_date(cx.DATE[p])}, not provided in the books - no deferred tax, as the interest is disallowed under Section 23 of the "
                               "MSMED Act"})
    # ---- current tax: as per return / computation, or on the restated taxable income ("computed")
    CT, PT = "Current tax", "Provision for income tax (net of advance tax)"
    for p in chron:
        v = (RST.get("tax") or {}).get(p)
        if v in (None, ""):
            continue
        if v == "computed":
            x = f"('Tax Computation'!{cx.NC[p]}{{{{tax_payable}}}}-{book(CT, p)})"
            how = "on the restated taxable income (Tax Computation)"
        else:
            x = f"({_n(v):g}/{S}-{book(CT, p)})"
            how = "as per the return of income / computation of total income"
        E.append({"date": cx.DATE[p], "lines": [L(PT, f"-{x}"), L(CT, x, f"Current tax (P&L - {_tag(cx, p)})", True)],
                  "narration": f"Being current tax for the {cx.per_word(p)} ended {long_date(cx.DATE[p])} restated {how}"})
    return E


def _tag(cx, p):
    return f"period ended {cx.DATE[p]:%d.%m.%Y}" if p == "S" else "FY " + fy_label(cx.DATE[p])


def _side(v):
    """expression of a Dr / Cr cell of an entry line (formula, number or blank)"""
    if isinstance(v, str) and v.startswith("="):
        return "(" + v[1:] + ")"
    try:
        return f"{float(v):g}" if v not in (None, "") else "0"
    except (TypeError, ValueError):
        return "0"


def _book_residual(cx, entries):
    """python-side estimate (Rs. lakhs) of the movement in the books' reserves other than the profit, less what the entries with
    plain amounts (and the gratuity / PPE 'through reserves' inputs) already reverse - decides whether an entry is needed"""
    from .book_inputs import grp_of
    I, sc, RS = cx.I, cx.scale_v, "Reserves and surplus"
    rows = lambda p: [x for x in I["tbs"][p]["rows"] if x.get("master") and not x.get("group") and not x.get("dup")]
    cl = lambda x: _n(x.get("cl")) or 0.0
    nxt = {cx.PK[i]: cx.PK[i - 1] for i in range(1, len(cx.PK))}

    def has_pla(p):
        return any(x["master"] == "Net profit as per books" and abs(cl(x)) > 0.005 for x in rows(p))

    def profit(p):
        """profit as per books: the P&L A/c ledger, or computed from the P&L ledgers and closing stock (memorandum rows)"""
        if has_pla(p):
            return cx.tbval.get((p, "Net profit as per books"), 0)
        stk = (_n((I.get("closing_stock") or {}).get({"S": "S"}.get(p, "Y1"))) or 0) if p in ("S", "Y1") else \
            sum(abs(cl(x)) for x in rows(nxt[p]) if x["master"] == "Opening stock")
        v = sum(-cl(x) for x in rows(p) if grp_of(cx, x["master"]) == "P&L") - sum(abs(cl(x)) for x in rows(p) if x["master"] == "Opening stock") + stk
        return v / sc

    def rs(p):
        v = cx.tbval.get((p, RS), 0)
        if p != "OPEN" and not has_pla(p):
            v = round(v - 1e-6, 9) + profit(p)
        return v

    known = {}
    g = ((I.get("restatement") or {}).get("gratuity") or {})
    for p, v in (g.get("direct_reserve") or {}).items() if g.get("dbo") else []:
        known[p] = known.get(p, 0) + (_n(v) or 0) / sc
    fa = I.get("fa") or {}
    if fa.get("auto_entries", True) and (fa.get("assets") or fa.get("source")):
        for p, v in (fa.get("direct_reserve") or {}).items():
            known[p] = known.get(p, 0) - (_n(v) or 0) / sc
    out = {}
    for p in cx.PK:
        lo, hi = ddate(cx.DATE[cx.PREV[p]]), ddate(cx.DATE[p])
        k = known.get(p, 0)
        for e in entries:
            d = e.get("date")
            if not e.get("manual") or isinstance(d, str) and d.startswith("=") or not (lo < ddate(d) <= hi):
                continue
            k += sum((_n(ln.get("cr")) or 0) - (_n(ln.get("dr")) or 0) for ln in e["lines"] if ln.get("master") == RS)
        out[p] = rs(p) - rs(cx.PREV[p]) - profit(p) + k
    return out


def _pl_leg(x, sign, tag):
    """credit to Other non-operating income (reserves credited in the books) or debit to Miscellaneous expenses (debited)"""
    if sign > 0:
        return L("Other non-operating income", f"-({x})", f"Other non-operating income (P&L - {tag})")
    return L("Miscellaneous expenses", f"-({x})", f"Miscellaneous expenses (P&L - {tag})")


def reserve_reversal(cx, entries):
    """Movements of the books' Reserves and surplus other than the profit for the period (prior period items, gratuity or PPE
    adjustments passed through reserves ...) are reversed: in restated statements every item belongs to the profit of the period
    it relates to. One entry per period, its amount a formula = the movement not yet reversed by the other entries, so the check row
    in Note 'Reserves and surplus' is nil and the balance sheet tallies.

    restatement.reserve_reversal = {"off": bool, p: {"relates": "same" | "OPEN" | earlier period, "master": balance sheet account}}
      same  -> recognised in the Statement of Profit and Loss of the same period (Other non-operating income / Miscellaneous expenses)
      other -> recognised in that earlier period (P&L, or the opening reserves for OPEN) against the balance sheet account given, and
               the books' reserve entry is reversed against the same account in the period it was passed"""
    cfg = (cx.I.get("restatement") or {}).get("reserve_reversal") or {}
    if cfg.get("off"):
        return []
    RS, S = "Reserves and surplus", cx.SCALE
    RSq = qs(RS)
    est = _book_residual(cx, entries)
    out = []
    for p in reversed(cx.PK):
        if abs(est.get(p, 0)) < 0.00005:          # nothing moved through reserves, or already reversed by the entries given
            continue
        lo, hi = ddate(cx.DATE[cx.PREV[p]]), ddate(cx.DATE[p])
        other = ""
        for e in entries:
            d = e.get("date")
            if isinstance(d, str) and d.startswith("="):
                continue
            if not (lo < ddate(d) <= hi):
                continue
            for ln in e["lines"]:
                if ln.get("master") == RS:
                    other += f"+{_side(ln.get('cr'))}-{_side(ln.get('dr'))}"
        x = f"({cx.tb(p, RSq)}-{cx.tb(cx.PREV[p], RSq)}-{cx.tb(p, qs('Net profit as per books'))}{other})"
        c = cfg.get(p) or {}
        rel, bsm = c.get("relates") or "same", c.get("master")
        tag = _tag(cx, p)
        if rel == "same" or not bsm or rel not in cx.PREV and rel != "OPEN" or (rel != "OPEN" and cx.PK.index(rel) <= cx.PK.index(p)):
            out.append(_order({"date": cx.DATE[p], "auto_rr": True, "lines": [L(RS, x, f"Reserves and surplus (movement in the books - {tag})"), _pl_leg(x, est[p], tag)],
                "narration": f"Being the amount credited / (debited) directly to Reserves and surplus in the books during the {cx.per_word(p)} ended "
                             f"{long_date(cx.DATE[p])}, recognised in the restated Statement of Profit and Loss of that {cx.per_word(p)}"}, est[p]))
            continue
        # recognised in the earlier period it relates to, against the balance sheet account the books used
        if rel == "OPEN":
            first = [L(bsm, x), L(RS, f"-({x})", f"Reserves and surplus (opening balance as at {long_date(cx.opening)})")]
            when = f"the opening reserves as at {long_date(cx.opening)}"
        else:
            rt = _tag(cx, rel)
            first = [L(bsm, x), _pl_leg(x, est[p], rt)]
            when = f"the Statement of Profit and Loss of the {cx.per_word(rel)} ended {long_date(cx.DATE[rel])}"
        out.append(_order({"date": cx.DATE[rel], "auto_rr": True, "lines": first,
                    "narration": f"Being the amount credited / (debited) directly to Reserves and surplus in the books during the {cx.per_word(p)} ended "
                                 f"{long_date(cx.DATE[p])} ({bsm}), recognised in {when} to which it relates"}, est[p]))
        out.append(_order({"date": cx.DATE[p], "auto_rr": True, "lines": [L(RS, x, f"Reserves and surplus (movement in the books - {tag})"), L(bsm, f"-({x})")],
                    "narration": f"Being the above entry passed through Reserves and surplus in the books during the {cx.per_word(p)} ended "
                                 f"{long_date(cx.DATE[p])} reversed"}, est[p]))
    return out


def _order(e, sg):
    """the sign of the movement is known from the books: debit lines first, "To" on the credit lines"""
    for ln in e["lines"]:
        neg = str(ln["dr"]).startswith("=MAX(-(")
        ln["to"] = (sg > 0) == neg
    e["lines"].sort(key=lambda ln: ln["to"])
    return e
