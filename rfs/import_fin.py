"""Shareholders and trade receivables / payables ageing from the company's audited financial statements workbook
(the Excel financials prepared for each year: party-wise ageing sheets, the share capital note with the list of
shareholders and the promoters' shareholding).

read_financials(path) -> {"debtor_ageing": {iso: [[5 buckets] x 4 rows]}, "creditor_ageing": {iso: [[4 buckets] x 4 rows]},
                          "shareholders": {name: {"name", "promoter", iso: shares}}, "msg": [...]}
Amounts in Rs; dates are the balance sheet dates (ISO).
"""
import datetime
import re

import openpyxl

DR_ROWS = ["undisputed good", "undisputed doubtful", "disputed good", "disputed doubtful"]


def _num(v):
    try:
        return float(v) if v not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def _fy_end(v):
    """'FY 23-24' / 'FY 2023-24' / '2023-24' / a date -> year end date"""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    m = re.search(r"(20\d{2}|\d{2})\s*[-/]\s*(\d{2,4})", str(v or ""))
    if not m:
        return None
    y = int(m.group(2)[-2:])
    return datetime.date(2000 + y, 3, 31)


def _bucket(h):
    """column heading -> bucket: 0 <6m, 1 6m-1y, 2 1-2y, 3 2-3y, 4 >3y, 'lt1' less than 1 year; None otherwise"""
    s = re.sub(r"\s+", " ", str(h or "")).lower()
    if not s or "total" in s:
        return None
    if re.search(r"(less than|<|upto|up to|within) ?6 ?m", s):
        return 0
    if re.search(r"6 ?m\w* ?(-|to) ?1 ?y", s):
        return 1
    if re.search(r"(less than|<|upto|up to|within) ?1 ?y", s):
        return "lt1"
    if re.search(r"1 ?(-|to) ?2 ?y", s):
        return 2
    if re.search(r"2 ?(-|to) ?3 ?y", s):
        return 3
    if re.search(r"(more than|>|above|over) ?3 ?y", s):
        return 4
    return None


def _party_ageing(ws):
    """party-wise ageing list: a heading row with FY / Particulars / bucket columns, one row per party and year.
    Returns ({date: {row_key: [5 buckets]}}, has_classification)"""
    for hr in range(1, min(ws.max_row, 15) + 1):
        heads = {c: ws.cell(hr, c).value for c in range(1, min(ws.max_column, 30) + 1)}
        bk = {c: _bucket(h) for c, h in heads.items() if _bucket(h) is not None}
        fyc = next((c for c, h in heads.items() if re.fullmatch(r"\s*(fy|year|financial year|period)\s*", str(h or ""), re.I)), None)
        if len(bk) >= 3 and fyc:
            break
    else:
        return None
    clc = next((c for c, h in heads.items() if re.search(r"classif|categor|msme|status|nature", str(h or ""), re.I)), None)
    out = {}
    for r in range(hr + 1, ws.max_row + 1):
        fy = ws.cell(r, fyc).value
        d = _fy_end(fy) if fy not in (None, "") and not re.search(r"total", str(fy), re.I) else None
        if not d:
            if out and fy in (None, ""):          # end of the list (blank line, then totals / pivot tables)
                if all(ws.cell(r, c).value in (None, "") for c in bk):
                    continue
                if re.search(r"total", str(ws.cell(r, fyc + 1).value or ""), re.I):
                    break
            continue
        cl = str(ws.cell(r, clc).value or "").lower() if clc else ""
        v = [0.0] * 5
        for c, b in bk.items():
            x = _num(ws.cell(r, c).value)
            if b == "lt1":
                v[0] += x
            else:
                v[b] += x
        row = out.setdefault(d, {})
        key = ("disputed " if "disput" in cl else "") + ("msme" if re.search(r"(?<!non )(?<!non-)msme|micro|small", cl) and "non" not in cl else
                                                         "doubtful" if "doubt" in cl else "other")
        row[key] = [a + b for a, b in zip(row.get(key, [0.0] * 5), v)]
    return out


def _shareholders(ws):
    """share capital note: 'list of shareholders ... 5%' table (In Nos / In % per date) and 'shares held by promoters'"""
    found, prom = {}, set()
    r = 1
    while r <= ws.max_row:
        b = " ".join(str(ws.cell(r, c).value or "") for c in range(1, 4))
        if re.search(r"share ?holders? (having|holding)|list of share ?holders|name of share ?holders", b, re.I):
            # heading row with the dates
            hr = r if any(isinstance(ws.cell(r, c).value, (datetime.date, datetime.datetime)) for c in range(2, 12)) else r + 1
            for k in range(r, r + 3):
                if any(isinstance(ws.cell(k, c).value, (datetime.date, datetime.datetime)) for c in range(2, 12)):
                    hr = k
                    break
            cols, seen = [], []
            for c in range(2, 14):
                v = ws.cell(hr, c).value
                if isinstance(v, (datetime.date, datetime.datetime)):
                    d = v.date() if isinstance(v, datetime.datetime) else v
                    while d in seen:                     # a copied heading: the next column is the year before
                        d = d.replace(year=d.year - 1)
                    seen.append(d)
                    cols.append((c, d))
            namec = next((c for c in range(1, 4) if re.search(r"name", str(ws.cell(hr, c).value or ws.cell(r, c).value or ""), re.I)), 2)
            k = hr + 1
            while k <= ws.max_row:
                nm = ws.cell(k, namec).value
                txt = str(nm or "").strip()
                if re.search(r"^(as per|total|note|\*)", txt, re.I) or re.search(r"promoter", txt, re.I):
                    break
                if txt and not re.search(r"in nos|in %|^nos?$", txt, re.I) and any(isinstance(ws.cell(k, c).value, (int, float)) for c, _ in cols):
                    h = found.setdefault(_key(txt), {"name": _clean(txt), "promoter": False})
                    for c, d in cols:
                        h[d.isoformat()] = _num(ws.cell(k, c).value)
                k += 1
            r = k
            continue
        if re.search(r"(shares )?held by promoters?|name of promoter", b, re.I):
            k = r + 1
            while k <= ws.max_row and k < r + 60:
                txt = str(ws.cell(k, 2).value or "").strip()
                if re.search(r"^(as per|total|note|\*)", txt, re.I):
                    break
                if txt and not re.search(r"name of promoter|no\.? of shares", txt, re.I) and isinstance(ws.cell(k, 3).value, (int, float)):
                    prom.add(_key(txt))
                k += 1
            r = k
            continue
        r += 1
    for k_ in prom:
        if k_ in found:
            found[k_]["promoter"] = True
    return found


def _clean(s):
    s = re.sub(r"\s+", " ", str(s)).strip()
    s = re.sub(r"\.(?=\S)", ". ", s)                 # C.K.Lenin -> C. K. Lenin
    return re.sub(r"\s+", " ", s.replace(".", " ")).strip()


def _key(s):
    return re.sub(r"[^a-z]", "", str(s).lower())


def is_financials(path):
    try:
        names = openpyxl.load_workbook(path, read_only=True).sheetnames
    except Exception:
        return False
    if "MasterData" in names and "BS_1" in names:          # an RFS workbook
        return False
    return any(re.search(r"(debtor|receivable|creditor|payable).*age?ing|age?ing.*(debtor|receivable|creditor|payable)", n, re.I) for n in names) \
        or any(re.search(r"notes? to (bs|balance)", n, re.I) for n in names)


def read_financials(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    res = {"debtor_ageing": {}, "creditor_ageing": {}, "shareholders": {}, "msg": []}
    for ws in wb.worksheets:
        n = ws.title.lower()
        kind = "dr" if re.search(r"debtor|receivable", n) else "cr" if re.search(r"creditor|payable", n) else None
        if kind and re.search(r"age?ing", n):
            pa = _party_ageing(ws)
            if not pa:
                continue
            for d, rows in pa.items():
                iso = d.isoformat()
                if kind == "dr":
                    good = [a + b for a, b in zip(rows.get("other", [0.0] * 5), rows.get("msme", [0.0] * 5))]
                    m = [good, rows.get("doubtful", [0.0] * 5), rows.get("disputed other", [0.0] * 5), rows.get("disputed doubtful", [0.0] * 5)]
                    res["debtor_ageing"][iso] = [[round(x, 2) for x in r] for r in m]
                else:
                    lt = lambda v: [v[0] + v[1], v[2], v[3], v[4]]
                    m = [lt(rows.get("msme", [0.0] * 5)), lt(rows.get("other", [0.0] * 5)), lt(rows.get("disputed msme", [0.0] * 5)),
                         lt(rows.get("disputed other", [0.0] * 5))]
                    res["creditor_ageing"][iso] = [[round(x, 2) for x in r] for r in m]
            res["msg"].append(f"{ws.title}: {'receivables' if kind == 'dr' else 'payables'} ageing for " +
                              ", ".join(f"{d:%d-%m-%Y}" for d in sorted(pa)))
        if re.search(r"notes? to (bs|balance)", n) or n.startswith("note"):
            sh = _shareholders(ws)
            for k, v in sh.items():
                cur = res["shareholders"].setdefault(k, {"name": v["name"], "promoter": False})
                cur.update({kk: vv for kk, vv in v.items() if kk not in ("name", "promoter")})
                cur["promoter"] = cur["promoter"] or v["promoter"]
            if sh:
                res["msg"].append(f"{ws.title}: {len(sh)} shareholders")
    return res
