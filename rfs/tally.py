"""Tally trial balance reader: signs balances from Dr/Cr formats, rebuilds the group tree
so only ledgers are counted, and drops ledgers that Tally lists twice."""
import sys, signal, threading, collections
import openpyxl
sys.setrecursionlimit(20000)
def sign(c):
    v=c.value
    if v in (None,''): return 0.0
    return -float(v) if 'Cr' in c.number_format else float(v)
def load(path):
    ws=openpyxl.load_workbook(path).active
    out=[]; primary=None; groups=[]
    start = 8
    for r in range(1, min(ws.max_row, 30) + 1):          # data starts after the header block
        if any(isinstance(ws.cell(r, c).value, (int, float)) for c in range(2, 6)) and ws.cell(r, 1).value not in (None, ''):
            start = r; break
    for r in range(start,ws.max_row+1):
        a=ws.cell(r,1)
        if a.value in (None,'') or str(a.value).strip()=='Grand Total': continue
        name=str(a.value).strip()
        kind='P' if a.font.b else ('L' if a.font.i else 'G')
        v=[sign(ws.cell(r,2)),float(ws.cell(r,3).value or 0),float(ws.cell(r,4).value or 0),sign(ws.cell(r,5))]
        if name=='Profit & Loss A/c': kind='P'
        if kind=='P': primary=name; lastgrp=''
        if kind=='G': lastgrp=name
        out.append(dict(r=r,name=name,kind=kind,primary=primary,group=lastgrp if kind!='P' else '',v=v))
    import collections
    cnt=collections.Counter((x['kind'] if x['kind']!='G' else 'L',x['name'],tuple(round(a,2) for a in x['v'])) for x in out)
    for x in out:
        x['dup']= x['kind']!='P' and any(abs(a)>0 for a in x['v']) and cnt[('L' if x['kind']=='G' else x['kind'],x['name'],tuple(round(a,2) for a in x['v']))]>1
    return out

TOL=0.011
def build(rows):
    n=len(rows)
    zero=lambda v: all(abs(x)<TOL for x in v)
    key=lambda v: tuple(round(x,2) for x in v)
    fail=set(); gmemo={}; steps=[0]
    def seq(p,rem,stop_at=None):
        # returns (end, kids) for nodes from p summing to rem; kids entries (p,e,t) or ('skip',p)
        steps[0]+=1
        if steps[0]>300000: raise TimeoutError('tree search budget exceeded')
        if zero(rem) and (stop_at is None or p==stop_at): return (p,[])
        if p>=n or rows[p]['kind']=='P': return None
        k=(p,key(rem),stop_at)
        xr=rows[p]['v']
        if xr[1]>rem[1]+TOL or xr[2]>rem[2]+TOL:   # transactions are additive and non-negative: a child cannot exceed what is left
            if rows[p]['dup']:
                r2=seq(p+1,rem,stop_at)
                if r2: return (r2[0],[('skip',p)]+r2[1])
            return None
        if k in fail: return None
        x=rows[p]
        if x['dup'] and x.get('second'):
            r=seq(p+1,rem,stop_at)
            if r: return (r[0],[('skip',p)]+r[1])
        opts=[]
        if x['kind']=='G':
            g=group(p)
            if g: opts.append(g)
        opts.append((p+1,None))
        for e,t in opts:
            r=seq(e,[a-b for a,b in zip(rem,x['v'])],stop_at)
            if r: return (r[0],[(p,e,t)]+r[1])
        if x['dup'] and not x.get('second'):
            r=seq(p+1,rem,stop_at)
            if r: return (r[0],[('skip',p)]+r[1])
        fail.add(k); return None
    def group(k):
        if k in gmemo: return gmemo[k]
        gmemo[k]=None
        if zero(rows[k]['v']): gmemo[k]=None; return None
        r=seq(k+1,rows[k]['v'])
        if r and r[1] and any(c[0]!='skip' for c in r[1]): gmemo[k]=(r[0],r[1])
        return gmemo[k]
    out=[]; skipped=[]
    def walk(kids,path):
        for c in kids:
            if c[0]=='skip': skipped.append(rows[c[1]]['r']); continue
            p,e,t=c; x=rows[p]
            if t is None:
                if zero(x['v']): continue
                out.append(dict(x,path=path))
            else: walk(t,path+[x['name']])
    tops=[k for k,x in enumerate(rows) if x['kind']=='P']
    for i,k in enumerate(tops):
        nxt=tops[i+1] if i+1<len(tops) else n
        if nxt==k+1:
            if not zero(rows[k]['v']): out.append(dict(rows[k],path=[]))
            continue
        r=seq(k+1,rows[k]['v'],stop_at=nxt)
        if r: walk(r[1],[rows[k]['name']])
        else: raise TimeoutError('primary group does not reconcile: '+rows[k]['name'])
    return out,skipped


def _timeout(*a):
    raise TimeoutError


def read_tb(path):
    """Parse one Tally TB export. Returns list of row dicts with keys
    r, name, kind, primary, group, v(opening, debit, credit, closing signed), type, inc, parent."""
    return _finalize(load(path))


def _finalize(rows):
    seen = set()
    for x in rows:
        k = (x['name'], tuple(round(a, 2) for a in x['v']))
        x['second'] = x['dup'] and k in seen
        seen.add(k)
    tops = [k for k, x in enumerate(rows) if x['kind'] == 'P']
    for x in rows:
        x['type'] = x['kind']; x['inc'] = 0; x['parent'] = ''
    for i, k in enumerate(tops):
        nxt = tops[i + 1] if i + 1 < len(tops) else len(rows)
        sub = rows[k:nxt]
        ok = True
        use_alarm = hasattr(signal, 'SIGALRM') and threading.current_thread() is threading.main_thread()
        try:
            if use_alarm:
                signal.signal(signal.SIGALRM, _timeout); signal.alarm(10)
            try:
                L, sk = build(sub)
            except (TimeoutError, RecursionError):
                # retry without the second listing of ledgers that Tally prints twice
                keep = [x for x in sub if not x.get('second')]
                try:
                    L, sk = build(keep)
                    sk = sk + [x['r'] for x in sub if x.get('second')]
                except (TimeoutError, RecursionError):
                    # ... or without the first listing
                    later = {}
                    for x in sub:
                        if x.get('dup'):
                            later[(x['name'], tuple(round(a, 2) for a in x['v']))] = x['r']
                    firsts = [x['r'] for x in sub if x.get('dup') and later.get((x['name'], tuple(round(a, 2) for a in x['v']))) != x['r']]
                    keep = [dict(x, second=False) for x in sub if x['r'] not in firsts]
                    L, sk = build(keep)
                    sk = sk + firsts
        except (TimeoutError, RecursionError):
            ok = False
        finally:
            if use_alarm: signal.alarm(0)
        if ok:
            byr = {x['r']: x for x in L}
            for x in sub:
                if x['r'] in byr:
                    x['inc'] = 1
                    x['parent'] = byr[x['r']]['path'][-1] if byr[x['r']]['path'] else ''
                if x['r'] in sk: x['type'] = 'X'
        else:   # fallback: italic ledgers, later duplicates dropped, parent = nearest group above
            last = sub[0]['name']
            for x in sub[1:]:
                if x['kind'] == 'G': last = x['name']
                if x['kind'] == 'L':
                    if x['second']: x['type'] = 'X'
                    else: x['inc'] = 1; x['parent'] = last
        if len(sub) == 1 and any(abs(a) > 0.001 for a in sub[0]['v']):
            sub[0]['inc'] = 1
    return rows


def tb_check(rows):
    """sum of included ledgers: (opening, debit, credit, closing) - opening/closing must be ~0, Dr = Cr"""
    return [round(sum(x['v'][t] for x in rows if x['inc']), 2) for t in range(4)]


# ------------------------------------------------------------------ any-format trial balances
import os, re
TALLY_PRIMARIES = {
    "capital account", "loans (liability)", "current liabilities", "fixed assets", "investments", "current assets",
    "branch / divisions", "branch/divisions", "suspense a/c", "sales accounts", "purchase accounts", "direct incomes",
    "indirect incomes", "direct expenses", "direct expense", "indirect expenses", "misc. expenses (asset)", "profit & loss a/c",
}


def _num(t):
    """'1,23,456.00 Dr' -> 123456.0 ; '(500)' -> -500 ; 'Cr' suffix -> negative"""
    if t is None:
        return None, None
    t = str(t).strip().replace("\u20b9", "").replace("Rs.", "").replace("INR", "")
    if t in ("", "-", "--", "nil", "Nil"):
        return 0.0, None
    side = None
    m = re.search(r"\b(Dr|Cr)\.?$", t, re.I)
    if m:
        side = m.group(1).capitalize(); t = t[:m.start()].strip()
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()").replace(",", "").replace(" ", "")
    try:
        v = float(t)
    except ValueError:
        return None, side
    return (-v if neg else v), side


def _find_header(table):
    keys = ("particular", "ledger", "account", "name", "opening", "debit", "credit", "closing", "balance")
    best, best_i = 0, None
    for i, row in enumerate(table[:40]):
        nxt = table[i + 1] if i + 1 < len(table) else []
        txt = [(row[j] if j < len(row) else "") + " " + (nxt[j] if j < len(nxt) else "") for j in range(max(len(row), len(nxt)))]
        score = sum(1 for k in keys if any(k in c.lower() for c in txt))
        if score > best:
            best, best_i = score, i
    return best_i if best >= 2 else None


def _columns(table, h):
    """map role -> column index using the header row (+ next row for two-line headers)"""
    row = table[h]; nxt = table[h + 1] if h + 1 < len(table) else []
    n = max(len(r) for r in table[h:h + 30])
    up = [""] * n
    cur = ""
    for j in range(n):                       # carry merged top-header text to the right
        top = row[j].strip() if j < len(row) else ""
        cur = top or cur
        up[j] = cur.lower()
    low = [(nxt[j].lower().strip() if j < len(nxt) else "") for j in range(n)]
    two_line = any(x in ("debit", "credit", "balance", "dr", "cr") for x in low)
    role = {}
    for j in range(n):
        top, sub = up[j], (low[j] if two_line else "")
        own = (row[j].lower() if j < len(row) else "")
        lab = (own + " " + sub).strip() if two_line else own
        if "name" not in role and any(k in lab or k in top for k in ("particular", "ledger", "account name", "name of")):
            role.setdefault("name", j); continue
        if "group" in lab and "group" not in role:
            role["group"] = j; continue
        if "opening" in top or "opening" in lab:
            if "credit" in lab or sub in ("credit", "cr"): role["op_cr"] = j
            elif "debit" in lab or sub in ("debit", "dr"): role["op_dr"] = j
            else: role.setdefault("op", j)
        elif "closing" in top or "closing" in lab or ("balance" in lab and "opening" not in lab):
            if "credit" in lab or sub in ("credit", "cr"): role["cl_cr"] = j
            elif "debit" in lab or sub in ("debit", "dr"): role["cl_dr"] = j
            else: role.setdefault("cl", j)
        elif "debit" in lab or lab.strip() in ("dr", "dr."):
            role.setdefault("dr", j)
        elif "credit" in lab or lab.strip() in ("cr", "cr."):
            role.setdefault("cr", j)
    role.setdefault("name", 0)
    return role, (h + 2 if two_line else h + 1)


def parse_generic(tables):
    """Flat trial balance table(s) -> rows in the same structure as load()."""
    table = max(tables, key=lambda t: sum(1 for r in t if any(_num(c)[0] for c in r[1:] if c)))
    h = _find_header(table)
    if h is None:
        raise ValueError("Could not find the header row (Particulars / Debit / Credit / Closing) in the trial balance.")
    role, start = _columns(table, h)
    out, primary, group = [], None, ""

    def get(r, k):
        j = role.get(k)
        return r[j] if j is not None and j < len(r) else ""

    def signed(val, side, default_side="Dr"):
        if val is None:
            return 0.0
        s = side or default_side
        return -abs(val) if s == "Cr" else (val if side is None and val < 0 else abs(val))

    for r in table[start:]:
        name = get(r, "name").strip()
        if not name or name.lower() in ("grand total", "total", "particulars"):
            continue
        dr, _ = _num(get(r, "dr")); cr, _ = _num(get(r, "cr"))
        if "op" in role:
            v, sd = _num(get(r, "op")); op = signed(v, sd)
        else:
            a, _ = _num(get(r, "op_dr")); b, _ = _num(get(r, "op_cr")); op = (a or 0) - (b or 0)
        if "cl" in role:
            v, sd = _num(get(r, "cl")); cl = signed(v, sd)
        elif "cl_dr" in role or "cl_cr" in role:
            a, _ = _num(get(r, "cl_dr")); b, _ = _num(get(r, "cl_cr")); cl = (a or 0) - (b or 0)
        else:   # only Debit / Credit columns given: treat as closing balances
            cl = op + (dr or 0) - (cr or 0)
        dr, cr = dr or 0.0, cr or 0.0
        if all(abs(x) < 0.005 for x in (op, dr, cr, cl)) and not name:
            continue
        if "cl" in role and "dr" not in role and "cr" not in role:   # closing only -> derive movement
            dr, cr = (max(cl - op, 0), max(op - cl, 0))
        g = get(r, "group").strip() if "group" in role else ""
        kind = "P" if name.lower() in TALLY_PRIMARIES and not g else ("L" if g else "G")
        if kind == "P":
            primary, group = name, ""
        out.append(dict(r=len(out) + 8, name=name, kind=kind, primary=(g or primary or "Unclassified"),
                        group=("" if kind == "P" else (g or group)), v=[op, dr, cr, cl]))
    if not out:
        raise ValueError("No ledger rows found in the trial balance.")
    if not any(x["kind"] == "P" for x in out):     # no group headings at all: every row is a ledger
        for x in out:
            x["kind"] = "L"
        out.insert(0, dict(r=0, name="Unclassified", kind="P", primary="Unclassified", group="", v=[0, 0, 0, 0]))
        out[0]["v"] = [sum(x["v"][i] for x in out[1:]) for i in range(4)]
    cnt = collections.Counter((x["kind"], x["name"], tuple(round(a, 2) for a in x["v"])) for x in out)
    for i, x in enumerate(out):
        x["r"] = i + 8
        x["dup"] = x["kind"] != "P" and any(abs(a) > 0 for a in x["v"]) and cnt[(x["kind"], x["name"], tuple(round(a, 2) for a in x["v"]))] > 1
    return out


def write_normalised(rows, path, title="Trial Balance"):
    """Writes rows as a Tally-style TB (bold primaries, italic ledgers, Dr/Cr formats) so the builder can copy it."""
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook(); ws = wb.active; ws.title = "Trial Balance"
    ws["A1"] = title; ws["A1"].font = Font(name="Arial", size=12, bold=True)
    ws["A2"] = "Trial Balance (normalised from uploaded file)"; ws["A2"].font = Font(name="Arial", size=10)
    for c, h in zip("ABCDE", ["Particulars", "Opening", "Debit", "Credit", "Closing"]):
        ws[f"{c}6"] = h; ws[f"{c}6"].font = Font(name="Arial", size=9, bold=True)
    for c, h in zip("BCDE", ["Balance", "", "", "Balance"]):
        ws[f"{c}7"] = h or None
    for x in rows:
        r = x["r"]
        f = Font(name="Arial", size=9, bold=x["kind"] == "P", italic=x["kind"] == "L")
        ws.cell(r, 1, x["name"]).font = f
        for col, val, sided in ((2, x["v"][0], True), (3, x["v"][1], False), (4, x["v"][2], False), (5, x["v"][3], True)):
            if abs(val) < 0.005:
                continue
            c = ws.cell(r, col, abs(val) if sided else val)
            c.font = Font(name="Arial", size=9)
            c.number_format = ('""0.00" Cr"' if val < 0 else '""0.00" Dr"') if sided else '""0.00'
    ws.column_dimensions["A"].width = 60
    for c in "BCDE":
        ws.column_dimensions[c].width = 15
    wb.save(path)
    return path


def _is_tally_xlsx(path):
    import openpyxl
    try:
        ws = openpyxl.load_workbook(path).active
    except Exception:
        return False
    italics = sum(1 for r in range(1, min(ws.max_row, 400) + 1) if ws.cell(r, 1).font and ws.cell(r, 1).font.i)
    sided = sum(1 for r in range(1, min(ws.max_row, 400) + 1) if "Dr" in str(ws.cell(r, 5).number_format) or "Cr" in str(ws.cell(r, 5).number_format))
    return italics > 5 and sided > 5


def read_any_tb(path, work_dir):
    """Any format -> (rows, path_of_tally_style_xlsx). Tally Excel exports are used as-is."""
    from . import docreader
    os.makedirs(work_dir, exist_ok=True)
    p = docreader.to_modern(path, work_dir)
    if p.lower().endswith((".xlsx", ".xlsm")) and _is_tally_xlsx(p):
        return read_tb(p), p
    rows = parse_generic(docreader.extract_tables(p, work_dir))
    norm = os.path.join(work_dir, os.path.splitext(os.path.basename(path))[0] + "_normalised.xlsx")
    write_normalised(rows, norm, os.path.basename(path))
    return read_tb(norm), norm
