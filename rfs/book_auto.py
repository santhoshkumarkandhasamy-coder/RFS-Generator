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
