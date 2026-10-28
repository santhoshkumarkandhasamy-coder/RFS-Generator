"""RFS Generator - Streamlit app.

Upload the three Tally trial balances, fill in company details, review the ledger mapping,
enter restatement adjustments, and download the Restated Financial Statements workbook.
Run:  streamlit run app.py
"""
import datetime as dt
import json
import os
import re
import tempfile

import pandas as pd
import streamlit as st

from rfs import tally, builder, mapping, exporters, docreader, regulatory, classify, extract, restate
from rfs import book_core, book_masters, book_inputs, book_notes, book_fa, book_statements, book_annex, book_auto, book, ui_book


def _reload_if_updated():
    """After a git pull while the app is running, Streamlit re-runs app.py but keeps the old rfs modules in memory:
    reload any module whose file changed since it was loaded."""
    import importlib
    for mod in (book_masters, classify, mapping, tally, docreader, regulatory, extract, restate, builder, exporters, book_core, book_inputs,
                book_notes, book_fa, book_statements, book_annex, book_auto, book, ui_book):     # dependencies first
        try:
            m = os.path.getmtime(mod.__file__)
        except OSError:
            continue
        loaded = getattr(mod, "_loaded_mtime", None)
        if loaded is not None and m != loaded:
            importlib.reload(mod)
        mod._loaded_mtime = m


_reload_if_updated()

st.set_page_config(page_title="RFS Generator", layout="wide")
BASE_SLOTS = [("FY26", "Latest year"), ("FY25", "Previous year"), ("FY24", "Earliest year")]
SLOTS = list(BASE_SLOTS)          # a stub period slot is added in front when the Company tab says so

ss = st.session_state
ss.setdefault("tb", {})          # slot -> parsed rows
ss.setdefault("tbfile", {})      # slot -> saved path
ss.setdefault("workdir", tempfile.mkdtemp(prefix="rfs_"))
ss.setdefault("overrides", {})
ss.setdefault("project", {})
ss.setdefault("docs", {})          # supporting documents: name -> {"path", "type", "text"}
ss.setdefault("regcheck", {})      # results of the ICDR / DRHP check
FILE_TYPES = docreader.SUPPORTED


def fy_label(d):
    return f"FY {d.year - 1}-{str(d.year)[2:]}"


def P(key, default=None):
    return ss.project.get(key, default)


# ---------------------------------------------------------------- sidebar: project file
with st.sidebar:
    st.header("RFS Generator")
    st.caption("SEBI ICDR - Indian GAAP - Schedule III Division I")
    quick = st.file_uploader("Upload input files (any format)", type=FILE_TYPES + ["json"], accept_multiple_files=True, key="quickup",
                             help="Excel, CSV, Word, PowerPoint, PDF, .doc / .xls / .ppt, or a saved project (.json). Trial balances are "
                                  "placed in the year found in the file name or heading (e.g. FY 2025-26, 25-26); other files go to "
                                  "Supporting documents. You can also upload in tabs 3 and 4.")
    quick_msg = st.container()
    for up in quick or []:
        if up.name.lower().endswith(".json") and ss.get("_loaded") != up.name + str(up.size):
            proj = json.load(up)
            ss.project = proj.get("project", {})
            ss.overrides = proj.get("overrides", {})
            ss._loaded = up.name + str(up.size)
            quick_msg.success(f"Project {up.name} loaded. Upload the trial balances again.")
    st.markdown("**Steps**\n1. ICDR & DRHP check\n2. Company\n3. Trial balances\n4. Supporting documents\n5. Mapping\n6. Shareholders & ageing\n"
                "7. Restatement (FA register, gratuity, MSME, tax, borrowings)\n8. Adjustments\n9. Generate")
    if not ss.get("_settings_loaded"):
        regulatory.load_settings()
        ss._settings_loaded = True
    with st.expander("Settings - source URLs"):
        for k in list(regulatory.SOURCES):
            regulatory.SOURCES[k] = st.text_input(k, regulatory.SOURCES[k], key="url_" + k)
        for k in list(regulatory.EXCHANGE_SOURCES):
            regulatory.EXCHANGE_SOURCES[k] = st.text_input(k, regulatory.EXCHANGE_SOURCES[k], key="url_" + k)
            regulatory.DATA_ADDRESSES[k] = st.text_input(f"{k.split(' - ')[0]} data address (copied from the browser)",
                                                         regulatory.DATA_ADDRESSES[k], key="data_" + k, help=regulatory.DATA_HOW_TO,
                                                         placeholder="https://.../api/...")
        if st.button("Save settings", help="Remembered on this PC for next time"):
            try:
                regulatory.save_settings()
                st.success("Saved.")
            except Exception as e:
                st.error(f"Could not save: {e}")

TABS = st.tabs(["1. ICDR & DRHP check", "2. Company", "3. Trial balances", "4. Supporting documents", "5. Mapping review",
                "6. Shareholders & ageing", "7. Restatement", "8. Adjustments", "9. Notes inputs", "10. Texts", "11. Generate"])
reg_tab, docs_tab, rst_tab, notes_tab, texts_tab = TABS[0], TABS[3], TABS[6], TABS[8], TABS[9]
tabs = [TABS[1], TABS[2], TABS[4], TABS[5], TABS[7], TABS[10]]
mapping.CUSTOM[:] = [c for c in (ss.project.get("notes", {}).get("custom_masters") or []) if c.get("name") and c.get("note")]

# ---------------------------------------------------------------- 1. company
LABELS = {"name": "Company name", "former": "Formerly known as", "cin": "CIN", "address": "Registered address",
          "incorporated": "Incorporation (background)", "latest": "Latest year end", "authorised_capital": "Authorised capital (Rs)",
          "face_value": "Face value (Rs)", "tax_rate": "Tax rate (115BAA)", "aud_firm": "Auditor - firm", "aud_frn": "Auditor - FRN",
          "aud_partner": "Auditor - partner", "aud_mno": "Auditor - membership no.", "aud_udin": "Auditor - UDIN",
          "aud_peer": "Auditor - peer review no.", "aud_place": "Auditor - place", "aud_date": "Auditor - date of signing",
          "directors": "Directors"}


def fill_from_documents():
    """Company tab fields from the uploaded documents; only empty fields are filled."""
    texts = {k: v.get("text", "") for k, v in ss.docs.items() if v.get("text")}
    found = extract.company_details(texts)
    pr, aud, filled = ss.project, dict(ss.project.get("auditor", {})), []
    for f, (val, src) in found.items():
        if f.startswith("aud_"):
            k = {"date": "sign_date"}.get(f[4:], f[4:])
            if not aud.get(k):
                aud[k] = val; filled.append(f)
        elif f == "directors":
            if not any(str(d.get("name", "")).strip() for d in pr.get("directors", [])):
                pr["directors"] = val; filled.append(f)
        elif f == "incorporated":
            if not pr.get("background"):
                d, _, act = val.partition(" ")
                when = dt.date.fromisoformat(d).strftime("%B %d, %Y") if d[:4].isdigit() else ""
                pr["background"] = (f"The Company was incorporated on {when} {act}".strip() + ".") if when else ""
                filled.append(f)
        elif f == "latest":
            pr["latest"] = val.isoformat(); filled.append(f)
        elif f in ("authorised_capital",) and not pr.get(f):
            pr[f] = val; filled.append(f)
        elif f in ("face_value", "tax_rate"):
            pr[f] = val; filled.append(f)
        elif not pr.get(f):
            pr[f] = val; filled.append(f)
    pr["auditor"] = aud
    ss._fill_n = ss.get("_fill_n", 0) + 1          # new widget keys so the filled values show
    ss._filled = [{"Field": LABELS.get(f, f), "Value": (f"{len(v)} found" if f == "directors" else str(v)),
                   "Found in": src, "Filled": "Yes" if f in filled else "No - field already had a value"} for f, (v, src) in found.items()]


with tabs[0]:
    fb1, fb2 = st.columns([2, 3])
    if fb1.button("Fill company details from uploaded documents", disabled=not ss.docs,
                  help="Reads the documents in Supporting documents / the sidebar upload (audited financials, auditor's report, "
                       "MCA master data, board report ...). Only empty fields are filled."):
        fill_from_documents()
        st.rerun()
    if not ss.docs:
        fb2.caption("Upload the audited financial statements and other documents first (sidebar or tab 4).")
    if ss.get("_filled") is not None:
        with st.expander(f"Found in documents ({len(ss._filled)} items) - check before generating", expanded=True):
            if ss._filled:
                st.dataframe(pd.DataFrame(ss._filled), hide_index=True, width="stretch")
            else:
                st.info("Nothing found. Scanned (image) PDFs have no readable text - upload the Word / Excel / text PDF version.")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Company")
        name = st.text_input("Company name", P("name", ""))
        former = st.text_input("Formerly known as (line under name)", P("former", ""))
        cin = st.text_input("CIN", P("cin", ""))
        address = st.text_area("Registered address", P("address", ""))
        segment = st.text_input("Business segment (for segment policy)", P("segment", ""))
        background = st.text_area("Background (incorporation, conversion, registered office) - blank: built from the fields below",
                                  P("background", ""), height=120)
        objects = st.text_area("Main objects of the Company", P("objects", ""), height=100)
        i1, i2 = st.columns(2)
        inc_date = i1.date_input("Date of incorporation", dt.date.fromisoformat(P("incorporation_date")) if P("incorporation_date") else None,
                                 min_value=dt.date(1900, 1, 1), key="inc_date")
        roc = i2.text_input("Registrar of Companies (e.g. Chennai)", P("roc", ""))
        cin_note = st.text_input("Note next to the CIN (e.g. earlier CIN before conversion)", P("cin_note", ""))
        exchange = st.text_input("Exchange / platform for the offer (e.g. SME Platform of NSE (NSE Emerge))", P("exchange", ""))
    with c2:
        st.subheader("RFS period and inputs")
        latest = st.date_input("Latest year end", dt.date.fromisoformat(P("latest", "2026-03-31")))
        st.caption(f"Years covered: {fy_label(latest)}, {fy_label(latest.replace(year=latest.year - 1))}, {fy_label(latest.replace(year=latest.year - 2))}")
        authorised = st.number_input("Authorised capital (Rs)", value=float(P("authorised_capital", 0.0)), step=100000.0)
        fv = st.number_input("Face value per share (Rs)", value=float(P("face_value", 10.0)))
        bonus = st.number_input("Bonus / split factor after latest year end (1 = none)", value=float(P("bonus_factor", 1.0)))
        tax_rate = st.number_input("Tax rate (e.g. 0.25168 for Sec 115BAA)", value=float(P("tax_rate", 0.25168)), format="%.5f")
        closing_stock = st.number_input("Closing stock at latest year end (Rs) - not in Tally TB", value=float(P("closing_stock", 0.0)))
        has_stub = st.checkbox("Include a stub period (interim period after the latest year end)", value=bool(P("stub_end")),
                               help="ICDR: restated financial information for the stub period is needed when the latest audited year is more "
                                    "than six months old at filing. Upload the stub period trial balance in tab 3.")
        if has_stub:
            sd = P("stub_end") or latest.replace(month=9, day=30).isoformat()
            stub_end = st.date_input("Stub period ended", dt.date.fromisoformat(sd))
            stub_stock = st.number_input("Closing stock at the stub period end (Rs)", value=float(P("stub_closing_stock", 0.0)))
        else:
            stub_end, stub_stock = None, 0.0
        st.subheader("Auditor (peer reviewed)")
        aud = P("auditor", {})
        a1, a2 = st.columns(2)
        auditor = {
            "firm": a1.text_input("Firm name", aud.get("firm", "")), "frn": a2.text_input("FRN", aud.get("frn", "")),
            "peer": a1.text_input("Peer review certificate no.", aud.get("peer", "")), "partner": a2.text_input("Partner", aud.get("partner", "")),
            "mno": a1.text_input("Membership no.", aud.get("mno", "")), "udin": a2.text_input("UDIN", aud.get("udin", "")),
            "sign_date": a1.text_input("Date of signing", aud.get("sign_date", "")), "place": a2.text_input("Place", aud.get("place", "")),
        }
    with st.expander("Statutory auditor of the audited financial statements and pending inputs"):
        sa = P("stat_auditor", {}) or {}
        s1, s2, s3 = st.columns(3)
        sa = {"firm": s1.text_input("Statutory auditor - firm", sa.get("firm", "")), "frn": s2.text_input("Statutory auditor - FRN", sa.get("frn", "")),
              "partner": s3.text_input("Partner / M.No.", sa.get("partner", "")), "reports": sa.get("reports") or []}
        rep_df = st.data_editor(pd.DataFrame(sa["reports"] or [["Audit report FY ", ""]], columns=["Report", "Details (date, UDIN, opinion)"]),
                                num_rows="dynamic", width="stretch", key="sa_rep")
        sa["reports"] = [list(r) for r in rep_df.fillna("").values.tolist() if any(r)]
        pending = st.text_area("Inputs pending from the management (one per line - listed in red in MasterData)", P("pending", ""))
    st.subheader("Directors / signatories")
    st.caption("Row 1 and 2: directors signing the statements; row 3: CFO (and director); row 4: Company Secretary; then independent / other directors.")
    dir_df = st.data_editor(pd.DataFrame(P("directors", [{"din": "", "name": "", "designation": "Director", "category": "Promoter", "signatory": "Yes"}])),
                            num_rows="dynamic", width="stretch", key=f"dirs{ss.get('_fill_n', 0)}")
    ss.project.update({"name": name, "former": former, "cin": cin, "address": address, "segment": segment, "background": background,
                       "objects": objects, "latest": latest.isoformat(), "authorised_capital": authorised, "face_value": fv,
                       "bonus_factor": bonus, "tax_rate": tax_rate, "closing_stock": closing_stock, "auditor": auditor,
                       "stub_end": stub_end.isoformat() if stub_end else None, "stub_closing_stock": stub_stock,
                       "directors": dir_df.fillna("").to_dict("records"), "incorporation_date": inc_date.isoformat() if inc_date else None,
                       "roc": roc, "cin_note": cin_note, "exchange": exchange, "stat_auditor": sa, "pending": pending})

years = {"d26": latest, "d25": latest.replace(year=latest.year - 1), "d24": latest.replace(year=latest.year - 2)}
YKEY = {"FY26": "d26", "FY25": "d25", "FY24": "d24", "FYS": "dS"}
if stub_end:
    years["dS"] = stub_end
    SLOTS = [("FYS", "Stub period")] + BASE_SLOTS
DKEYS = [YKEY[s_] for s_, _ in SLOTS]                 # date keys, latest first


def plabel(slot):
    """Display name of a period slot."""
    return f"Period ended {years['dS']:%d-%b-%Y}" if slot == "FYS" else fy_label(years[YKEY[slot]])


def load_tb(slot, f):
    """Save an uploaded trial balance for a year slot and read it (raises if it cannot be read)."""
    import uuid
    ext = os.path.splitext(f.name)[1].lower()
    tdir = os.path.join(ss.workdir, "tb_" + slot, uuid.uuid4().hex[:8])   # own folder per upload: a rejected file never replaces the stored TB
    os.makedirs(tdir, exist_ok=True)
    path = os.path.join(tdir, f"TB_{slot}{ext}")
    with open(path, "wb") as fh:
        fh.write(f.getbuffer())
    rows, tbfile = tally.read_any_tb(path, tdir)
    ss.tb[slot], ss.tbfile[slot] = rows, tbfile
    ss["tbname_" + slot] = f.name + str(f.size)
    ss["tbsrc_" + slot] = f.name


def add_doc(f):
    """Save an uploaded supporting document and extract its text."""
    ddir = os.path.join(ss.workdir, "docs"); os.makedirs(ddir, exist_ok=True)
    p = os.path.join(ddir, f.name)
    with open(p, "wb") as fh:
        fh.write(f.getbuffer())
    try:
        txt = docreader.extract_text(p, ddir)
        err = None
    except Exception as e:
        txt, err = "", str(e)
    ss.docs[f.name] = {"path": p, "type": os.path.splitext(f.name)[1].lstrip(".").upper(), "text": txt,
                       "category": "", "added": dt.date.today().isoformat()}
    return err


def year_slot(text):
    """Year slot whose financial year appears in a file name / heading (FY 2025-26, 25-26, 2025-2026, FY26, 31-Mar-2026)."""
    t = re.sub(r"[_\s./]+", "-", text.lower())
    for slot, _ in BASE_SLOTS:          # a stub period TB is recognised from the period in its heading, not its name
        y = years[YKEY[slot]].year
        pats = [f"{y - 1}-{str(y)[2:]}", f"{str(y - 1)[2:]}-{str(y)[2:]}", f"{y - 1}-{y}", f"fy-?{str(y)[2:]}(?!-?\\d)", f"mar[a-z]*-?{y}"]
        if any(re.search(r"(?<!\d)" + p_, t) for p_ in pats):
            return slot
    return None


_PERIOD = re.compile(r"(\d{1,2}-[A-Za-z]{3}-\d{2,4})\s*to\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})", re.I)


def tb_period(text):
    """(start, end) of a Tally heading like '1-Apr-25 to 31-Mar-26', else None."""
    m = _PERIOD.search(text or "")
    if not m:
        return None
    out = []
    for s_ in m.groups():
        for fmt in ("%d-%b-%y", "%d-%b-%Y"):
            try:
                out.append(dt.datetime.strptime(s_.title(), fmt).date()); break
            except ValueError:
                pass
    return tuple(out) if len(out) == 2 else None


def slot_for_period(period):
    """Year slot whose year end is the period's end; None for a stub / other period."""
    return next((s_ for s_, _ in SLOTS if period and years[YKEY[s_]] == period[1]), None)


def names_other_year(name):
    """True when the file name states a financial year that is none of the three RFS years (e.g. FY 26-27)."""
    t = re.sub(r"[_\s./']+", "-", name.lower())
    return bool(re.search(r"(?<!\d)(20)?\d{2}-\d{2}(?!\d)", t)) and year_slot(name) is None


def heading_text(f):
    """First lines of a spreadsheet / CSV (Tally puts the period in the heading)."""
    ext = os.path.splitext(f.name)[1].lower()
    try:
        if ext in (".xlsx", ".xlsm"):
            import openpyxl
            ws = openpyxl.load_workbook(f, read_only=True, data_only=True).worksheets[0]
            return " ".join(str(v) for r in ws.iter_rows(max_row=10, values_only=True) for v in r if v is not None)
        if ext == ".csv":
            return f.getvalue()[:2000].decode("utf-8", "ignore")
    except Exception:
        pass
    finally:
        f.seek(0)
    return ""


# ---------------------------------------------------------------- sidebar upload: route files to TB slots / documents
for up in quick or []:
    key = up.name + str(up.size)
    if up.name.lower().endswith(".json") or key in ss.setdefault("_quick_done", set()):
        continue
    ss._quick_done.add(key)
    head = heading_text(up)
    period = tb_period(head)
    named_tb = bool(re.search(r"trial|\btb\b|tb[-_ ]", up.name, re.I))
    if period:                                   # the TB's own heading decides the year
        slot = slot_for_period(period)
        if not slot:
            quick_msg.warning(f"{up.name} covers {period[0]:%d-%b-%Y} to {period[1]:%d-%b-%Y}, which is not one of the three RFS years "
                              f"({', '.join(plabel(s_) for s_, _ in SLOTS)}). For a stub period, tick 'Include a stub period' in the "
                              "Company tab and upload it again - added to supporting documents.")
    else:
        slot = year_slot(up.name) or year_slot(head)
        if not slot and named_tb and not names_other_year(up.name):
            slot = next((s_ for s_, _ in SLOTS if s_ not in ss.tb), None)
        elif not slot and named_tb:
            quick_msg.warning(f"{up.name}: the year in the file name is not one of the three RFS years - added to supporting documents.")
    if slot and up.name.lower().rsplit(".", 1)[-1] in FILE_TYPES:
        if slot:
            before = {k: ss.get(k) for k in ("tbname_" + slot, "tbsrc_" + slot)}, ss.tb.get(slot), ss.tbfile.get(slot)
            with st.spinner(f"Reading {up.name} as trial balance..."):
                try:
                    load_tb(slot, up)
                    op, drs, crs, cl = tally.tb_check(ss.tb[slot])
                    ok = abs(op) < 1 and abs(cl) < 1 and abs(drs - crs) < 1 and sum(1 for x in ss.tb[slot] if x["inc"]) >= 10
                    if named_tb or ok:
                        quick_msg.success(f"{up.name} -> trial balance {plabel(slot)}")
                        continue
                    raise ValueError("not a trial balance")
                except Exception as e:
                    keys, tb_, tf_ = before          # put back whatever was in the slot
                    for k, v in keys.items():
                        ss[k] = v
                    if tb_ is None:
                        ss.tb.pop(slot, None); ss.tbfile.pop(slot, None)
                    else:
                        ss.tb[slot], ss.tbfile[slot] = tb_, tf_
                    if named_tb:
                        quick_msg.warning(f"{up.name} could not be read as a trial balance ({e}); added to supporting documents.")
    with st.spinner(f"Reading {up.name}..."):
        err = add_doc(up)
    quick_msg.info(f"{up.name} -> supporting documents" + (f" (text not extracted: {err})" if err else ""))

# ---------------------------------------------------------------- 2. trial balances
with tabs[1]:
    st.write("Upload the trial balance for each year in any format: Excel (.xlsx / .xls), CSV, Word (.docx / .doc), "
             "PowerPoint (.pptx / .ppt) or PDF. A Tally Excel export (Detailed view) is read as-is. Other files need columns for "
             "Particulars and Opening / Debit / Credit / Closing (Dr/Cr suffix or separate Dr and Cr columns); the group and "
             "ledger structure is rebuilt from the totals.")
    cols = st.columns(len(SLOTS))
    for (slot, lbl), col in zip(SLOTS, cols):
        with col:
            f = st.file_uploader(f"{lbl} - {plabel(slot)}", type=FILE_TYPES, key="tb_" + slot)
            if f is not None and ss.get("tbname_" + slot) != f.name + str(f.size):
                with st.spinner("Reading trial balance..."):
                    try:
                        per = tb_period(heading_text(f))
                        if per and years[YKEY[slot]] != per[1]:
                            st.warning(f"{f.name} covers {per[0]:%d-%b-%Y} to {per[1]:%d-%b-%Y}, not {plabel(slot)} "
                                       f"(year ended {years[YKEY[slot]]:%d-%b-%Y}). Check that this is the right file for this year.")
                        load_tb(slot, f)
                    except Exception as e:
                        st.error(f"Could not read {f.name}: {e}")
            if slot in ss.tb:
                op, drs, crs, cl = tally.tb_check(ss.tb[slot])
                n = sum(1 for x in ss.tb[slot] if x["inc"])
                ok = abs(op) < 1 and abs(cl) < 1 and abs(drs - crs) < 1
                (st.success if ok else st.error)(f"{n} ledgers | Dr {drs:,.2f} | Cr {crs:,.2f} | closing net {cl:,.2f}")
                if not ok:
                    st.caption("The ledgers found do not reconcile to nil. Upload the Tally Excel export (Detailed) if available, "
                               "or a file with a Group column; otherwise check the mapping tab for missing / repeated ledgers.")
                if not ss.tbfile[slot].endswith(f"TB_{slot}.xlsx"):
                    st.caption(f"Source: {ss.get('tbsrc_' + slot, '')} - normalised into a Tally-style TB sheet in the output.")

ready = all(s in ss.tb for s, _ in SLOTS)

# ---------------------------------------------------------------- ICDR & DRHP check
with reg_tab:
    st.write("Check the latest position before preparing the RFS: amendments under the SEBI (ICDR) Regulations, 2018 and the "
             "latest DRHP / RHP filings to use as format references. Needs internet access.")
    st.caption(f"RFS layout in this app last reviewed against ICDR / recent offer documents on {regulatory.FORMAT_REVIEWED_ON:%d %B %Y}.")
    c1, c2 = st.columns(2)
    if c1.button("Check latest ICDR updates", type="primary"):
        with st.spinner("Reading SEBI regulations and circulars..."):
            items, errs = regulatory.icdr_updates(regulatory.SOURCES)
        ss.regcheck["icdr"] = {"checked": dt.datetime.now().isoformat(timespec="minutes"), "items": items, "errors": errs}
    srcs = c2.multiselect("DRHP / RHP sources", ["SEBI", "NSE", "BSE"], default=["SEBI", "NSE", "BSE"], key="offer_srcs",
                          help="SEBI: offer documents filed with SEBI / RoC. NSE: Main board and Emerge. BSE: Main board and BSE SME.")
    if c2.button("Check latest DRHP / RHP filings"):
        with st.spinner(f"Reading offer document listings ({', '.join(srcs)})..."):
            items, errs = regulatory.all_offer_documents(srcs, None, regulatory.SOURCES, regulatory.EXCHANGE_SOURCES)
        ss.regcheck["offer"] = {"checked": dt.datetime.now().isoformat(timespec="minutes"), "items": items, "errors": errs, "period": "Latest"}
    today = dt.date.today()
    months = [dt.date(today.year + (today.month - 1 - k) // 12, (today.month - 1 - k) % 12 + 1, 1) for k in range(24)]   # last 24 months
    m1, m2 = c2.columns([3, 2])
    mon = m1.selectbox("Month", months, format_func=lambda d: d.strftime("%B %Y"), key="offer_month", label_visibility="collapsed")
    if m2.button("Month-wise filings"):
        with st.spinner(f"Reading DRHP / RHP filings of {mon:%B %Y}..."):
            items, errs = regulatory.all_offer_documents(srcs, (mon.year, mon.month), regulatory.SOURCES, regulatory.EXCHANGE_SOURCES)
        ss.regcheck["offer"] = {"checked": dt.datetime.now().isoformat(timespec="minutes"), "items": items, "errors": errs,
                                "period": f"{mon:%B %Y}"}

    ic = ss.regcheck.get("icdr")
    if ic:
        st.subheader(f"ICDR amendments and circulars (checked {ic['checked']})")
        for e in ic["errors"]:
            st.warning(f"Could not read {e.split(':')[0]} - no internet access or the site refused the request. Open the source links below manually.")
            with st.expander("Technical detail"):
                st.code(e)
        if ic["items"]:
            df = pd.DataFrame([{"Date": i["date"], "Title": i["title"], "Source": i["source"], "Link": i["url"]} for i in ic["items"]])
            st.dataframe(df, hide_index=True, width="stretch", column_config={"Link": st.column_config.LinkColumn("Link", display_text="Open")})
            newer = [i for i in ic["items"] if i["date"] and i["date"] > regulatory.FORMAT_REVIEWED_ON]
            if newer:
                st.error(f"{len(newer)} ICDR-related update(s) are dated after the app's format review "
                         f"({regulatory.FORMAT_REVIEWED_ON:%d-%m-%Y}). Read them and adjust the RFS before issue.")
            else:
                st.success("No ICDR-related update after the app's format review date.")
    of = ss.regcheck.get("offer")
    if of:
        per = of.get("period", "Latest")
        st.subheader(f"{'Latest DRHP / RHP filings' if per == 'Latest' else 'DRHP / RHP filings - ' + per} (checked {of['checked']})")
        if per != "Latest" and not of["items"] and not of["errors"]:
            st.info(f"No DRHP / RHP filings found for {per}.")
        for e in of["errors"]:
            src = e.split(":")[0]
            if src in regulatory.DATA_ADDRESSES and not regulatory.DATA_ADDRESSES[src]:
                st.warning(f"{src.split(' - ')[0]}: set its data address once in the sidebar (Settings - source URLs) so the app can "
                           "read it directly. How to find it: " + regulatory.DATA_HOW_TO + "  Or use *Add a list downloaded from NSE / BSE* below.")
            elif src.startswith("NSE") or src.startswith("BSE"):
                st.warning(f"Could not read {src}. Check the address in Settings - source URLs, or open the page (Source links below), "
                           "use its Download (.csv / Excel) option, and add the file under *Add a list downloaded from NSE / BSE*.")
            else:
                st.warning(f"Could not read {src} - no internet access or the site refused the request. Open the source links below manually.")
            with st.expander("Technical detail"):
                st.code(e)
        with st.expander("Add a list downloaded from NSE / BSE", expanded=any(e.startswith(("NSE", "BSE")) for e in of["errors"])):
            ex1, ex2 = st.columns([1, 3])
            exch = ex1.selectbox("Exchange", ["NSE", "BSE", "BSE SME"], key="imp_exch")
            imp = ex2.file_uploader("Downloaded list (CSV, Excel, saved web page, PDF)", type=["csv", "xlsx", "xls", "html", "htm", "pdf", "docx"], key="imp_list")
            if imp is not None and ss.get("_imp_done") != imp.name + str(imp.size) + exch:
                p = os.path.join(ss.workdir, "import_" + imp.name)
                with open(p, "wb") as fh:
                    fh.write(imp.getbuffer())
                try:
                    new = regulatory.import_listing(p, exch, ss.workdir)
                    if of.get("period", "Latest") != "Latest":       # keep to the month being viewed
                        new = [it for it in new if it["date"] and it["date"].strftime("%B %Y") == of["period"]]
                    have = {(it["title"], it["url"], it.get("exchange")) for it in of["items"]}
                    of["items"] = regulatory.merge_offers(of["items"] + [it for it in new if (it["title"], it["url"], exch) not in have])
                    of["errors"] = [e for e in of["errors"] if not e.startswith(exch.split()[0])] if new else of["errors"]
                    ss._imp_done = imp.name + str(imp.size) + exch
                    st.success(f"{len(new)} {exch} filings added from {imp.name}.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not read {imp.name}: {e}")
        if of["items"]:
            items = of["items"]
            cache = regulatory.load_location_cache()
            for it in items:                         # locations already found earlier (cached on this PC)
                if it["url"] in cache:
                    it.update({k: cache[it["url"]].get(k, "") for k in ("city", "district", "state", "pin", "address", "error")})
            missing = [it for it in items if regulatory.needs_location(cache.get(it["url"]))]
            lc1, lc2 = st.columns([3, 2])
            lc1.caption("SEBI lists only the company name. To filter by state, the app reads the registered office "
                        "on the cover page of each offer document (downloads each PDF once; results are remembered on this PC).")
            auto = lc2.checkbox("Find locations automatically for every filing", value=True, key="auto_loc",
                                help="Reads each new filing's cover page as soon as the list loads. Untick to do it only with the button.")
            run_key = f"{of['checked']}|{of.get('period')}|{len(items)}"
            run_auto = auto and missing and ss.get("_auto_loc_for") != run_key     # once per loaded list (failed ones: use the button)
            if run_auto:
                ss._auto_loc_for = run_key
            if missing and (run_auto or lc2.button(f"Find company locations ({len(missing)} filings)")):
                bar = st.progress(0.0, text="Reading cover pages...")
                regulatory.locate_all(items, ss.workdir, cache, progress=lambda d, n: bar.progress(d / n, text=f"Reading cover pages... {d}/{n}"))
                bar.empty()
                st.rerun()

            f1, f2 = st.columns(2)
            states = sorted({it.get("state") for it in items if it.get("state")})
            sel_states = f1.multiselect("State", states, key="flt_state")
            word = f2.text_input("Company name contains", key="flt_word")
            shown = [it for it in items if (not sel_states or it.get("state") in sel_states)
                     and (not word or word.lower() in it["title"].lower())]
            if sel_states and missing:
                st.caption(f"{len(missing)} filing(s) have no location yet and are hidden by the filter - press Find company locations.")

            df = pd.DataFrame([{"Date": i["date"], "Company": i["title"], "Type": i["source"], "Source": i.get("exchange", "SEBI"),
                                "Filed / listed on": i.get("seen_on", ""), "State": i.get("state", ""),
                                "Location": ", ".join(x for x in (i.get("city", ""), i.get("pin", "")) if x), "Link": i["url"]} for i in shown])
            st.dataframe(df, hide_index=True, width="stretch", column_config={"Link": st.column_config.LinkColumn("Link", display_text="Open")})
            failed = [it for it in items if it.get("error")]
            if failed:
                with st.expander(f"State not found for {len(failed)} filing(s) - why"):
                    st.dataframe(pd.DataFrame([{"Company": it["title"], "Source": it.get("exchange", "SEBI"), "Reason": it["error"],
                                                "Link": it.get("pdf") or it["url"]} for it in failed]),
                                 hide_index=True, width="stretch", column_config={"Link": st.column_config.LinkColumn("Link", display_text="Open")})
                    st.caption("Press Find company locations again to retry. Scanned PDFs and sites that refuse downloads cannot be read.")
            located = [it for it in items if it.get("state")]
            if located:
                with st.expander("State-wise summary", expanded=True):
                    sm = pd.DataFrame([{"State": it["state"], "Source": it.get("exchange", "SEBI")} for it in located])
                    st.dataframe(sm.pivot_table(index="State", columns="Source", aggfunc="size", fill_value=0).assign(Total=lambda d: d.sum(axis=1))
                                 .sort_values("Total", ascending=False), width="stretch")
            labels = [f"{i['date']} | {i.get('exchange', 'SEBI')} {i['source']} | {i['title']}" + (f" | {i['state']}" if i.get("state") else "") for i in shown]
            pick = st.selectbox("Use as format reference (downloads the offer document PDF and compares it)", ["-"] + labels)
            if pick != "-" and st.button("Download and compare"):
                it = shown[labels.index(pick)]
                try:
                    p = os.path.join(ss.workdir, "reference_offer_document.pdf")
                    with st.spinner("Downloading (offer documents are large)..."):
                        link = regulatory.get_offer_pdf(it["url"], p, it.get("pdf"))
                    ss.regcheck["reference"] = {"name": it["title"] + (f" ({it['state']})" if it.get("state") else ""),
                                                "type": it["source"], "url": link, "date": str(it["date"]),
                                                "analysis": regulatory.analyse_offer_document(p)}
                except Exception as e:
                    st.warning(f"Download failed ({e}). Download the PDF from the link and upload it below.")

    with st.expander("Source links (open manually if the automatic check cannot reach them)"):
        for k, v in {**regulatory.SOURCES, **regulatory.EXCHANGE_SOURCES, **regulatory.MANUAL_LINKS}.items():
            st.markdown(f"- [{k}]({v})")

    st.subheader("Compare with a DRHP / RHP")
    ref = st.file_uploader("Upload a DRHP / RHP (PDF, or Word)", type=["pdf", "docx", "doc"], key="refdoc")
    if ref is not None and ss.get("_refname") != ref.name:
        p = os.path.join(ss.workdir, "reference_" + ref.name)
        with open(p, "wb") as fh:
            fh.write(ref.getbuffer())
        with st.spinner("Reading the restated financial information section..."):
            pdf = p
            if not p.lower().endswith(".pdf"):
                pdf = exporters.docx_to_pdf(docreader.to_modern(p, ss.workdir), ss.workdir)
            ss.regcheck["reference"] = {"name": ref.name, "type": "Uploaded", "url": "", "date": dt.date.today().isoformat(),
                                        "analysis": regulatory.analyse_offer_document(pdf)}
        ss._refname = ref.name
    rf = ss.regcheck.get("reference")
    if rf:
        a = rf["analysis"]
        st.write(f"**Reference:** {rf['name']} ({rf['type']})" + (f" - restated section found on pages {a['pages'][0]}-{a['pages'][1]} of {a['total_pages']}" if a["pages"] else " - restated section not found"))
        st.dataframe(pd.DataFrame(a["coverage"]), hide_index=True, width="stretch")
        if a["not_in_app"]:
            with st.expander(f"Headings in the reference not matched to this app's annexures ({len(a['not_in_app'])}) - consider adding"):
                for h in a["not_in_app"]:
                    st.write("- " + h)
        with st.expander("All headings found in the reference restated section"):
            for h in a["headings"]:
                st.write("- " + h)

    st.subheader("ICDR checklist followed by the RFS layout")
    st.dataframe(pd.DataFrame(regulatory.CHECKLIST, columns=["Requirement", "Source"]), hide_index=True, width="stretch")

# ---------------------------------------------------------------- supporting documents
with docs_tab:
    st.write("Upload the supporting documents in any format (Excel, Word, PowerPoint, PDF, .doc / .xls / .ppt): audited financial "
             "statements, sanction letters, actuarial reports, ageing, tax computations, related party details, etc. "
             "Text is extracted so you can search them while completing the inputs; the list is recorded in the output workbook.")
    ups = st.file_uploader("Supporting documents", type=FILE_TYPES, accept_multiple_files=True, key="docs_upload")
    for f in ups or []:
        if f.name in ss.docs:
            continue
        with st.spinner(f"Reading {f.name}..."):
            err = add_doc(f)
        if err:
            st.warning(f"{f.name}: text could not be extracted ({err}). It is still listed.")
    if ss.docs:
        cats = ["", "Audited financial statements", "Trial balance", "Sanction letter / loan statement", "Actuarial report",
                "Ageing", "Tax computation / ITR / 3CD", "Related party", "Fixed asset register", "Contingent liabilities",
                "Company records (MOA / AOA / ROC)", "Other"]
        df = pd.DataFrame([{"Document": k, "Type": v["type"], "Category": v.get("category", ""),
                            "Characters read": len(v["text"])} for k, v in ss.docs.items()])
        ed = st.data_editor(df, hide_index=True, width="stretch", key="docsed", disabled=["Document", "Type", "Characters read"],
                            column_config={"Category": st.column_config.SelectboxColumn(options=cats)})
        for _, rw in ed.iterrows():
            ss.docs[rw["Document"]]["category"] = rw["Category"] or ""
        q = st.text_input("Search in documents (e.g. closing stock, gratuity, sanction, contingent)")
        if q:
            hits = []
            for k, v in ss.docs.items():
                for line in v["text"].splitlines():
                    if q.lower() in line.lower():
                        hits.append({"Document": k, "Line": line.strip()[:300]})
            st.write(f"{len(hits)} line(s) found")
            st.dataframe(pd.DataFrame(hits[:500]), hide_index=True, width="stretch")
        pick = st.selectbox("Preview document text", ["-"] + list(ss.docs))
        if pick != "-":
            st.text_area("Extracted text", ss.docs[pick]["text"][:200000], height=300)

# ---------------------------------------------------------------- 3. mapping
with tabs[2]:
    if not ready:
        st.info("Upload all the trial balances first.")
    else:
        st.write("Each ledger is mapped to a note line (Masters). The 'if Dr' / 'if Cr' column is used depending on the balance, "
                 "so debit creditors and credit debtors are reclassified automatically. Edit any row, then press Save mapping.")
        rows = {}
        for slot, _ in SLOTS:
            for x in ss.tb[slot]:
                if not x["inc"]:
                    continue
                k = mapping.key(x)
                if k not in rows:
                    d, c = mapping.suggest(x["primary"], x["parent"], x["name"])
                    ov = ss.overrides.get(k)
                    if ov:
                        d, c = ov["dr"], ov["cr"]
                    rows[k] = {"key": k, "Primary group": x["primary"], "Group": x["parent"], "Ledger": x["name"],
                               **{s_: 0.0 for s_, _ in SLOTS}, "Master if Dr": d, "Master if Cr": c}
                rows[k][slot] = round(x["v"][3], 2)
        df = pd.DataFrame(list(rows.values()))
        extra = sorted({v for v in list(df["Master if Dr"]) + list(df["Master if Cr"]) if str(v).startswith(mapping.DIRECTOR_PREFIX)})
        options = mapping.all_labels(extra)
        f1, f2 = st.columns([2, 1])
        prim = f1.multiselect("Filter primary group", sorted(df["Primary group"].unique()))
        search = f2.text_input("Search ledger")
        view = df
        if prim: view = view[view["Primary group"].isin(prim)]
        if search: view = view[view["Ledger"].str.contains(search, case=False)]
        edited = st.data_editor(
            view, width="stretch", height=520, hide_index=True, key="mapedit",
            disabled=["key", "Primary group", "Group", "Ledger"] + [s_ for s_, _ in SLOTS],
            column_config={"key": None,
                           "Master if Dr": st.column_config.SelectboxColumn(options=options, width="large"),
                           "Master if Cr": st.column_config.SelectboxColumn(options=options, width="large"),
                           **{s_: st.column_config.NumberColumn(plabel(s_) + " closing (Dr+/Cr-)", format="%.2f") for s_, _ in SLOTS}})
        b1, b2, b3 = st.columns(3)
        if b1.button("Save mapping", type="primary"):
            for _, rw in edited.iterrows():
                ss.overrides[rw["key"]] = {"dr": rw["Master if Dr"], "cr": rw["Master if Cr"]}
            st.success(f"Saved {len(edited)} rows.")
        new_dir = b2.text_input("Add director loan master (name)")
        if new_dir:
            b2.caption(f"Use '{mapping.DIRECTOR_PREFIX}{new_dir}' in the Master columns")
        b3.download_button("Download mapping (.json)", json.dumps(ss.overrides, indent=1), "mapping.json")
        mup = b3.file_uploader("Load mapping (.json)", type="json", key="mapup")
        if mup is not None and ss.get("_maploaded") != mup.name:
            ss.overrides.update(json.load(mup)); ss._maploaded = mup.name
            st.rerun()

# ---------------------------------------------------------------- 4. shareholders & ageing
with tabs[3]:
    st.subheader("Register of members (number of shares)")
    SHK = {"FYS": "sS", "FY26": "s26", "FY25": "s25", "FY24": "s24"}
    sh_cols = {"name": "Name", **{SHK[s_]: plabel(s_) + " end" for s_, _ in SLOTS}, "promoter": "Promoter"}
    sh_df = pd.DataFrame(P("shareholders", [{"name": "", "s26": 0, "s25": 0, "s24": 0, "promoter": False}]))
    for k_ in sh_cols:
        if k_ not in sh_df.columns:
            sh_df[k_] = 0
    sh_df = sh_df[list(sh_cols)]
    sh_df = st.data_editor(sh_df, num_rows="dynamic", width="stretch", key="sh",
                           column_config={k: st.column_config.Column(v) for k, v in sh_cols.items()} | {"promoter": st.column_config.CheckboxColumn("Promoter")})
    ss.project["shareholders"] = sh_df.fillna(0).to_dict("records")

    def ageing_editor(title, key, buckets, rows_lbl):
        st.subheader(title + " (Rs)")
        store = P(key, {})
        out = {}
        for dk in DKEYS:
            base = store.get(dk, [])
            data = [[(base[j][i] if j < len(base) and i < len(base[j]) else 0) for i in range(len(buckets))] for j in range(4)]
            df_ = pd.DataFrame(data, columns=buckets, index=rows_lbl)
            st.caption(f"As at {years[dk]:%B %d, %Y}")
            e = st.data_editor(df_, width="stretch", key=f"{key}_{dk}")
            out[dk] = e.fillna(0).values.tolist()
        ss.project[key] = out

    a1, a2 = st.columns(2)
    with a1:
        ageing_editor("Trade receivables ageing", "ageing_dr", ["< 6 months", "6 m - 1 yr", "1-2 yrs", "2-3 yrs", "> 3 yrs"],
                      ["Undisputed - good", "Undisputed - doubtful", "Disputed - good", "Disputed - doubtful"])
    with a2:
        ageing_editor("Trade payables ageing", "ageing_cr", ["< 1 yr", "1-2 yrs", "2-3 yrs", "> 3 yrs"],
                      ["MSME", "Others", "Disputed - MSME", "Disputed - Others"])

# ---------------------------------------------------------------- 5. adjustments
with tabs[4]:
    st.write("Enter restatement adjustments as journal entries in Rs. lakhs. Lines with the same Entry no. form one entry "
             "(Dr total must equal Cr total). Balance sheet masters: date the entry on the balance sheet date it first affects "
             f"(opening adjustments: {years['d24'].replace(year=years['d24'].year - 1):%d-%m-%Y}). P&L masters: any date in the year.")
    extra = sorted({o.get("cr") for o in ss.overrides.values() if str(o.get("cr", "")).startswith(mapping.DIRECTOR_PREFIX)})
    options = [o for o in mapping.all_labels(extra) if o != mapping.EXCLUDE]
    adj_rows = P("adj_rows", [{"entry": 1, "date": years["d24"].isoformat(), "master": "Reserves and surplus", "particulars": "", "dr": 0.0, "cr": 0.0, "narration": ""}])
    adf = pd.DataFrame(adj_rows)
    adf["date"] = pd.to_datetime(adf["date"]).dt.date
    adf = st.data_editor(adf, num_rows="dynamic", width="stretch", key="adj", column_config={
        "entry": st.column_config.NumberColumn("Entry no.", step=1),
        "date": st.column_config.DateColumn("Date", format="DD-MM-YYYY"),
        "master": st.column_config.SelectboxColumn("Masters (note line)", options=options, width="large"),
        "particulars": st.column_config.TextColumn("Particulars (optional)"),
        "dr": st.column_config.NumberColumn("Dr (lakhs)", format="%.5f"), "cr": st.column_config.NumberColumn("Cr (lakhs)", format="%.5f"),
        "narration": st.column_config.TextColumn("Narration (first line of entry)", width="large")})
    adf = adf.dropna(subset=["master"])
    ss.project["adj_rows"] = [{**r, "date": (r["date"].isoformat() if hasattr(r["date"], "isoformat") else str(r["date"]))}
                              for r in adf.fillna({"dr": 0, "cr": 0, "particulars": "", "narration": ""}).to_dict("records")]
    entries = []
    for eno, g in adf.groupby("entry", sort=True):
        legs = [{"master": r["master"], "particulars": r["particulars"] or r["master"], "dr": float(r["dr"] or 0), "cr": float(r["cr"] or 0)}
                for _, r in g.iterrows() if (r["dr"] or r["cr"])]
        if not legs:
            continue
        drt, crt = sum(l["dr"] for l in legs), sum(l["cr"] for l in legs)
        if abs(drt - crt) > 0.00001:
            st.error(f"Entry {eno}: Dr {drt:.5f} not equal to Cr {crt:.5f}")
        narr = next((n for n in g["narration"] if isinstance(n, str) and n.strip()), "")
        entries.append({"date": g["date"].iloc[0], "legs": legs, "narration": narr})
    ss.entries = entries
    st.caption(f"{len(entries)} balanced entr{'y' if len(entries) == 1 else 'ies'} will be posted to 'Adj Entry RFS'.")

# ---------------------------------------------------------------- 7. restatement inputs
RST = ss.project.setdefault("restatement", {})
CHRON_SLOTS = [s_ for s_, _ in reversed(SLOTS)]                  # earliest first
OPEN_KEY = "d23"
years_open = years["d24"].replace(year=years["d24"].year - 1)


def _per_table(key, label, keys, lab_fn):
    """One-column editable table (Rs) keyed by date keys; stored in RST[key]."""
    store = RST.setdefault(key, {})
    df_ = pd.DataFrame({"Period / date": [lab_fn(k) for k in keys], label: [store.get(k) for k in keys]})
    e = st.data_editor(df_, hide_index=True, width="stretch", key="rst_" + key, disabled=["Period / date"],
                       column_config={label: st.column_config.NumberColumn(label, format="%.2f")})
    RST[key] = {k: (None if pd.isna(v) else float(v)) for k, v in zip(keys, e[label])}


def date_lab(k):
    return f"As at {years_open:%d-%b-%Y} (opening)" if k == OPEN_KEY else (f"{'Period' if k == 'dS' else 'Year'} ended {years[k]:%d-%b-%Y}")


def key_for_date(d):
    """Date key whose date equals d (period end / opening date)."""
    if d == years_open:
        return OPEN_KEY
    return next((k for k in DKEYS if years[k] == d), None)


with rst_tab:
    st.write("Inputs for the standard restatement adjustments. Each section posts formula-linked entries to 'Adj Entry RFS' (with deferred tax "
             "where applicable) and adds its working sheet to the workbook. Values can be read from the supporting documents and edited here; "
             "leave a section empty to skip it.")
    per_keys = [YKEY[s_] for s_ in CHRON_SLOTS]

    with st.expander("1. Fixed asset register - depreciation restated as per Schedule II (WDV)", expanded=bool(RST.get("fa_assets") or RST.get("fa_map_path"))):
        RST["fa_mode"] = "mapped" if st.radio("Register", ["Standard template - the app computes Schedule II WDV for each asset",
                                                           "Client's own register - use its WDV / depreciation columns"],
                                              index=1 if RST.get("fa_mode") == "mapped" else 0, key="fa_mode") .startswith("Client") else "standard"
    if RST.get("fa_mode") == "mapped":
      with st.expander("Client's register - column mapping", expanded=True):
        fum = st.file_uploader("Client's fixed asset register (Excel)", type=["xlsx", "xlsm"], key="fam_up")
        if fum is not None and ss.get("_fam_loaded") != fum.name + str(fum.size):
            pth = os.path.join(ss.workdir, "far_" + fum.name)
            with open(pth, "wb") as fh:
                fh.write(fum.getbuffer())
            RST["fa_map_path"] = pth
            ss._fam_loaded = fum.name + str(fum.size)
        if RST.get("fa_map_path") and os.path.exists(RST["fa_map_path"]):
            st.caption(f"Register: {os.path.basename(RST['fa_map_path'])} - copied into the 'FA Register' sheet as it is; normalised columns are added.")
            ui_book.fa_mapped_ui(ss, RST, SLOTS, plabel, RST["fa_map_path"])
            _per_table("fa_direct", "PPE through reserves in books (Rs)", per_keys, date_lab)
        elif RST.get("fa_map_path"):
            st.warning("Upload the register again (files are not kept between sessions).")
    with st.expander("Standard register - asset list", expanded=False) if RST.get("fa_mode") != "mapped" else st.container():
      if RST.get("fa_mode") != "mapped":
        fu = st.file_uploader("Fixed asset register (Excel, CSV, PDF, Word)", type=FILE_TYPES, key="fa_up")
        cand = [k for k in ss.docs if re.search(r"fixed asset|fa[ _-]?(details|register)|asset register|\bfar\b", k, re.I)]
        pick = st.selectbox("...or read it from the supporting documents", ["-"] + cand, key="fa_pick")
        src_path = None
        if fu is not None and ss.get("_fa_loaded") != fu.name + str(fu.size):
            src_path = os.path.join(ss.workdir, "fa_" + fu.name)
            with open(src_path, "wb") as fh:
                fh.write(fu.getbuffer())
            ss._fa_loaded = fu.name + str(fu.size)
        elif pick != "-" and ss.get("_fa_loaded") != pick:
            src_path = ss.docs[pick]["path"]; ss._fa_loaded = pick
        if src_path:
            with st.spinner("Reading the register..."):
                try:
                    RST["fa_assets"] = restate.read_fa_register(src_path, ss.workdir)
                    st.success(f"{len(RST['fa_assets'])} assets read.")
                except Exception as e:
                    st.error(f"Could not read the register: {e}")
        assets = RST.get("fa_assets") or []
        if assets:
            fdf = pd.DataFrame(assets)
            st.caption("Class mapping - register classification to the PPE note line (applies to all assets of that classification):")
            cm = fdf.groupby("source_class", dropna=False).agg(class_=("class", "first"), assets=("cost", "size"), cost=("cost", "sum")).reset_index()
            cm.columns = ["Register classification", "Note line", "Assets", "Cost (Rs)"]
            cm_e = st.data_editor(cm, hide_index=True, width="stretch", key="fa_cm", disabled=["Register classification", "Assets", "Cost (Rs)"],
                                  column_config={"Note line": st.column_config.SelectboxColumn(options=restate.FA_CLASSES, required=True),
                                                 "Cost (Rs)": st.column_config.NumberColumn(format="%.2f")})
            cmap = dict(zip(cm_e["Register classification"], cm_e["Note line"]))
            fdf["class"] = [cmap.get(sc, c) for sc, c in zip(fdf["source_class"], fdf["class"])]
            for c_ in ("date", "del_date"):
                fdf[c_] = pd.to_datetime(fdf[c_], errors="coerce").dt.date
            st.caption("Assets - enter the date of deletion and sale value for assets sold / discarded (the register often has no such column):")
            fe = st.data_editor(fdf, num_rows="dynamic", width="stretch", height=320, key="fa_grid", column_config={
                "name": "Asset", "source_class": "Register classification", "class": st.column_config.SelectboxColumn("Note line", options=restate.FA_CLASSES),
                "date": st.column_config.DateColumn("Date put to use", format="DD-MM-YYYY"), "cost": st.column_config.NumberColumn("Cost (Rs)", format="%.2f"),
                "rate": st.column_config.NumberColumn("WDV rate", format="%.4f", help="Blank: from useful life (5% residual value)"),
                "life": st.column_config.NumberColumn("Useful life (yrs)"),
                "del_date": st.column_config.DateColumn("Date of deletion", format="DD-MM-YYYY"), "sale": st.column_config.NumberColumn("Sale value (Rs)", format="%.2f")})
            RST["fa_assets"] = [{**r, "date": r["date"].isoformat() if hasattr(r["date"], "isoformat") else None,
                                 "del_date": r["del_date"].isoformat() if hasattr(r.get("del_date"), "isoformat") else None}
                                for r in fe.where(pd.notna(fe), None).to_dict("records") if r.get("cost")]
            st.caption(f"{len(RST['fa_assets'])} assets, cost Rs {sum(float(a['cost'] or 0) for a in RST['fa_assets']):,.2f}. "
                       "Assets without a rate use 1 - 0.05^(1/useful life); without both, the Schedule II life of the class.")
            if ready:
                book_cls = sorted({m[0] for s_, _ in SLOTS for x in ss.tb[s_] if x["inc"] for m in [mapping.master_for(x, 3, ss.overrides)]
                                   if m and m[0] in restate.FA_CLASSES})
                missing = [c_ for c_ in book_cls if c_ not in {a["class"] for a in RST["fa_assets"]}]
                if missing:
                    st.warning("Note lines with book balances but no asset in the register: " + ", ".join(missing) +
                               ". They are left as per books (not restated). Check the class mapping or the TB mapping.")
            st.caption("Book entries that changed PPE directly through reserves (e.g. an alignment with the register passed in the books), Rs - "
                       "increase in PPE positive:")
            _per_table("fa_direct", "PPE through reserves in books (Rs)", per_keys, date_lab)
            if st.button("Clear the register", key="fa_clear"):
                RST["fa_assets"] = []; ss._fa_loaded = None; st.rerun()

    with st.expander("2. Gratuity - provision restated to the actuarial valuation", expanded=bool((RST.get("dbo") or {}) and any((RST.get("dbo") or {}).values()))):
        if st.button("Read from actuarial reports in supporting documents", key="g_read"):
            found = 0
            for k_, d_ in ss.docs.items():
                for iso, amt in restate.read_gratuity(d_.get("text", "")).items():
                    kk = key_for_date(dt.date.fromisoformat(iso))
                    if kk:
                        RST.setdefault("dbo", {})[kk] = amt; found += 1
            (st.success if found else st.info)(f"{found} valuation(s) read." if found else "No actuarial valuation found - enter the figures below.")
        st.caption("Defined benefit obligation as per the actuarial valuation at each date (Rs). All dates are needed, including the opening date.")
        _per_table("dbo", "Obligation (Rs)", [OPEN_KEY] + per_keys, date_lab)
        st.caption("Gratuity charged directly to reserves in the books during the period (Rs) - reclassified to the Statement of Profit and Loss:")
        _per_table("gdir", "Charged to reserves in books (Rs)", per_keys, date_lab)

    with st.expander("3. MSME interest - Section 16 of the MSMED Act", expanded=bool(any((RST.get("msme") or {}).values()))):
        if st.button("Read from MSME interest computations in supporting documents", key="m_read"):
            found = 0
            for k_, d_ in ss.docs.items():
                if re.search(r"msme|micro|interest", k_, re.I) and d_.get("path") and d_["path"].lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
                    amt, end = restate.read_msme_interest(d_["path"], ss.workdir)
                    kk = key_for_date(end) if end else None
                    if amt is not None and kk:
                        RST.setdefault("msme", {})[kk] = amt; found += 1
            (st.success if found else st.info)(f"{found} computation(s) read." if found else "No MSME computation found - enter the figures below.")
        st.caption("Interest on delayed payments to micro and small enterprises for each period (Rs). The entry posts the difference from the "
                   "interest already booked. Not deductible (Section 23), so no deferred tax.")
        _per_table("msme", "MSME interest (Rs)", per_keys, date_lab)

    with st.expander("4. Current tax - as per return of income / computation", expanded=bool(any((RST.get("tax") or {}).values()))):
        if st.button("Read from ITR acknowledgements in supporting documents", key="t_read"):
            found = 0
            for k_, d_ in ss.docs.items():
                if re.search(r"itr|return|acknow|computation", k_, re.I):
                    amt, end = restate.read_tax(d_.get("text", ""))
                    kk = key_for_date(end) if end else None
                    if amt is not None and kk:
                        RST.setdefault("tax", {})[kk] = amt; found += 1
            (st.success if found else st.info)(f"{found} return(s) read." if found else "No return found - enter the figures below.")
        st.caption("Current tax as per the return / computation (Rs); blank = no adjustment. The entry posts the difference from the tax provided in the books.")
        _per_table("tax", "Tax as per return (Rs)", [k for k in per_keys if k != "dS"], date_lab)
        if stub_end:
            RST["tax_stub_computed"] = st.checkbox("Stub period: provide current tax on the restated taxable income ('Tax Computation' sheet)",
                                                   value=bool(RST.get("tax_stub_computed")), key="t_stub")



def tb_title(slot):
    """first lines of the (normalised) trial balance: company, 'Trial Balance', period"""
    try:
        import openpyxl
        ws_ = openpyxl.load_workbook(ss.tbfile[slot], read_only=True).worksheets[0]
        vals = [str(r[0]) for r in ws_.iter_rows(min_row=1, max_row=3, max_col=1, values_only=True) if r and r[0]]
        return (vals + ["", "", ""])[:3]
    except Exception:
        return [ss.project.get("name", ""), "Trial Balance", ""]


def tb_ledgers():
    """TB ledgers by use, for the notes inputs (borrowings, other income)"""
    out = {"borrow": set(), "oi": []}
    if not ready:
        return {"borrow": [], "oi": []}
    BORR = {"Term loans from banks and NBFCs (secured)", "Overdraft facilities from banks (secured)", "Bill discounting and trade finance loans (secured)",
            "Purchase bill finance and channel finance (unsecured)"}
    seen = set()
    for s_, _ in SLOTS:
        for x in ss.tb[s_]:
            if not x["inc"]:
                continue
            m = mapping.master_for(x, 3, ss.overrides)
            if not m:
                continue
            if m[0] in BORR:
                out["borrow"].add(x["name"])
            if m[0] in ("Interest income", "Other non-operating income") and (m[0], x["name"]) not in seen:
                seen.add((m[0], x["name"]))
                out["oi"].append((m[0], x["name"]))
    return {"borrow": sorted(out["borrow"]), "oi": out["oi"]}


with notes_tab:
    if st.button("Read borrowings from sanction letters in supporting documents", key="b_read2"):
        got = [restate.read_sanction(d_.get("text", ""), k_) for k_, d_ in ss.docs.items() if re.search(r"sanction", k_ + " " + d_.get("text", "")[:2000], re.I)]
        NB = ss.project.setdefault("notes", {}).setdefault("borrowings", [])
        for g_ in got:
            NB.append({"lender": g_.get("lender"), "nature": g_.get("facility"), "limit_rs": g_.get("sanctioned"), "roi": g_.get("rate"),
                       "tenor": g_.get("tenure"), "security": g_.get("security"), "guarantee": bool(g_.get("guarantee")), "ledgers": g_.get("ledger") or ""})
        (st.success if got else st.info)(f"{len(got)} sanction letter(s) read." if got else "No sanction letter found - add rows below.")
    ui_book.notes_tab(ss, P, SLOTS, plabel, years, ready, tb_ledgers())

with texts_tab:
    ui_book.texts_tab(ss)


def restatement_cfg():
    """Restatement inputs from the tab -> builder cfg (dates as date objects, tax keyed by period)."""
    out = {}
    if RST.get("fa_assets"):
        def d_(v):
            return dt.date.fromisoformat(v) if isinstance(v, str) and v else None
        out["fa"] = {"assets": [{**a, "date": d_(a.get("date")), "del_date": d_(a.get("del_date"))} for a in RST["fa_assets"] if d_(a.get("date"))],
                     "direct": {k: v for k, v in (RST.get("fa_direct") or {}).items() if v}}
    dbo = RST.get("dbo") or {}
    need = [OPEN_KEY] + [YKEY[s_] for s_ in CHRON_SLOTS]
    if all(dbo.get(k) not in (None, "") for k in need):
        out["gratuity"] = {"dbo": {k: dbo[k] for k in need}, "direct": {k: v for k, v in (RST.get("gdir") or {}).items() if v}}
    if any((RST.get("msme") or {}).values()):
        out["msme"] = {k: v for k, v in RST["msme"].items() if v not in (None, "")}
    tx = {s_: (RST.get("tax") or {}).get(YKEY[s_]) for s_ in CHRON_SLOTS if s_ != "FYS"}
    tx = {k: v for k, v in tx.items() if v not in (None, "")}
    if stub_end and RST.get("tax_stub_computed"):
        tx["FYS"] = "computed"
    if tx:
        out["tax"] = tx
    if RST.get("borrowings"):
        out["borrowings"] = RST["borrowings"]
    return out


# ---------------------------------------------------------------- 6. generate
with tabs[5]:
    st.download_button("Save project (.json) for later", json.dumps({"project": ss.project, "overrides": ss.overrides, "regcheck": ss.regcheck,
                                   "documents": {k: {kk: vv for kk, vv in v.items() if kk not in ("text", "path")} for k, v in ss.docs.items()}},
                                  indent=1, default=str),
                       f"{(name or 'rfs').replace(' ', '_')}_project.json")
    problems = []
    if not ready: problems.append("Upload all the trial balances" + (" (including the stub period)." if stub_end else "."))
    if not name or not cin: problems.append("Company name and CIN are required.")
    if ready:
        known = set(mapping.MASTER_INFO) | {mapping.EXCLUDE} | {c["name"] for c in mapping.CUSTOM}
        bad = {v for o in ss.overrides.values() for v in (o["dr"], o["cr"]) if v not in known and not str(v).startswith(mapping.DIRECTOR_PREFIX)}
        if bad: problems.append("Unknown masters in mapping: " + ", ".join(sorted(bad)))
    for p_ in problems:
        st.warning(p_)
    FORMATS = {"Excel (.xlsx)": "xlsx", "Word (.docx)": "docx", "Word 97-2003 (.doc)": "doc", "PowerPoint (.pptx)": "pptx", "HTML (.html)": "html"}
    engine = exporters.recalc_engine()
    chosen = st.multiselect("Output formats", list(FORMATS), default=["Excel (.xlsx)"])
    if engine:
        st.caption(f"Word / PowerPoint / HTML are produced from the calculated figures (recalculated with {engine}).")
    elif any(FORMATS[c] != "xlsx" for c in chosen):
        problems.append("Word / PowerPoint / HTML need LibreOffice installed (or Microsoft Excel with the pywin32 package on Windows).")
        st.warning(problems[-1])
    if st.button("Generate RFS", type="primary", disabled=bool(problems) or not chosen):
        cfg = dict(ss.project)
        cfg["years"] = {k: v for k, v in years.items() if k != "dS"}
        if stub_end:
            cfg["stub"] = {"end": stub_end, "closing_stock": stub_stock}
        cfg["restatement"] = restatement_cfg()
        cfg["adjustments"] = getattr(ss, "entries", [])
        cfg["directors"] = [d for d in cfg.get("directors", []) if d.get("name")]
        cfg["shareholders"] = [h for h in cfg.get("shareholders", []) if h.get("name")]
        run_dir = tempfile.mkdtemp(dir=ss.workdir)
        upto = f"{stub_end:%b-%y}" if stub_end else fy_label(years['d26'])[3:]
        out = os.path.join(run_dir, f"{(name or 'Company').replace(' ', '_')}_RFS_{fy_label(years['d24'])[3:]}_to_{upto}.xlsx")
        with st.spinner("Building workbook..."):
            book.build(ui_book.make_input(ss, P, SLOTS, years, stub_end, tb_title), out)
            ss.text_defaults = book.LAST_TEXTS
            exporters.add_reference_sheet(out, ss.regcheck, ss.docs,
                                          {plabel(s_): ss.get("tbsrc_" + s_, os.path.basename(ss.tbfile.get(s_, ""))) for s_, _ in SLOTS})
        with st.spinner("Preparing output files..."):
            try:
                ss.outputs = exporters.export_all(out, [FORMATS[c] for c in chosen], os.path.join(run_dir, "out"), name)
            except Exception as e:
                st.error(f"Export failed: {e}")
                ss.outputs = {"xlsx": out} if "Excel (.xlsx)" in chosen else {}
        calc = ss.outputs.get("_calc")
        ss.preview = None
        if calc:
            import openpyxl
            v = openpyxl.load_workbook(calc, data_only=True)
            bs, pl = v["BS_1"], v["PL_2"]
            pc = "DEFG"[:len(SLOTS)]
            rows_ = []
            for r in range(9, 60):
                lab = bs[f"B{r}"].value
                if lab and isinstance(bs[f"D{r}"].value, (int, float)) and "Total" in str(lab):
                    rows_.append([lab] + [bs[f"{c}{r}"].value for c in pc])
            for r in range(8, 32):
                lab = pl[f"B{r}"].value
                if lab and isinstance(pl[f"D{r}"].value, (int, float)) and any(k in str(lab) for k in ("Revenue", "Profit", "Total")):
                    rows_.append([lab] + [pl[f"{c}{r}"].value for c in pc])
            ss.preview = (rows_, [bs[f"{c}1"].value for c in pc], [plabel(s_) for s_, _ in SLOTS])

    if ss.get("preview"):
        rows_, diff, heads_ = ss.preview
        cols_ = ["Line (Rs. lakhs)"] + heads_
        st.dataframe(pd.DataFrame(rows_, columns=cols_), width="stretch", hide_index=True,
                     column_config={c: st.column_config.NumberColumn(format="%.2f") for c in cols_[1:]})
        (st.success if all(abs(d or 0) < 0.01 for d in diff) else st.error)(f"Balance sheet difference: {diff}")
    if ss.get("outputs") and any(not k.startswith("_") for k in ss.outputs):
        st.subheader("Downloads")
        mime = {"xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "doc": "application/msword",
                "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation", "html": "text/html"}
        label = {"xlsx": "Excel workbook", "docx": "Word document", "doc": "Word 97-2003 document", "pptx": "PowerPoint summary", "html": "HTML report"}
        files = {k: v for k, v in ss.outputs.items() if not k.startswith("_")}
        cols = st.columns(len(files))
        for col, (fmt, path) in zip(cols, files.items()):
            with open(path, "rb") as fh:
                col.download_button(f"Download {label[fmt]}", fh.read(), os.path.basename(path), mime=mime[fmt], key="dl_" + fmt)
        st.caption("Excel: every figure is a formula linked to the TB sheets - review the yellow input cells and grey check cells. "
                   "Word / HTML carry all annexures and notes; PowerPoint is a summary deck (statements, adjustments, ratios, capitalisation).")
