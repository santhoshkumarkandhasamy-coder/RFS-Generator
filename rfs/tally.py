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
        else: print('  FAIL PRIMARY',rows[k]['name'])
    return out,skipped


def _timeout(*a):
    raise TimeoutError


def read_tb(path):
    """Parse one Tally TB export. Returns list of row dicts with keys
    r, name, kind, primary, group, v(opening, debit, credit, closing signed), type, inc, parent."""
    rows = load(path)
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
            L, sk = build(sub)
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
