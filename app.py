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

from rfs import tally, builder, mapping, exporters, docreader, regulatory, classify, extract


def _reload_if_updated():
    """After a git pull while the app is running, Streamlit re-runs app.py but keeps the old rfs modules in memory:
    reload any module whose file changed since it was loaded."""
    import importlib
    for mod in (classify, mapping, tally, docreader, regulatory, extract, builder, exporters):     # dependencies first
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
SLOTS = [("FY26", "Latest year"), ("FY25", "Previous year"), ("FY24", "Earliest year")]

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
    st.markdown("**Steps**\n1. ICDR & DRHP check\n2. Company\n3. Trial balances\n4. Supporting documents\n5. Mapping\n6. Shareholders & ageing\n7. Adjustments\n8. Generate")
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
                "6. Shareholders & ageing", "7. Adjustments", "8. Generate"])
reg_tab, docs_tab = TABS[0], TABS[3]
tabs = [TABS[1], TABS[2], TABS[4], TABS[5], TABS[6], TABS[7]]

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
        background = st.text_area("Background (incorporation, conversion, registered office)", P("background", ""), height=120)
        objects = st.text_area("Main objects of the Company", P("objects", ""), height=100)
    with c2:
        st.subheader("RFS period and inputs")
        latest = st.date_input("Latest year end", dt.date.fromisoformat(P("latest", "2026-03-31")))
        st.caption(f"Years covered: {fy_label(latest)}, {fy_label(latest.replace(year=latest.year - 1))}, {fy_label(latest.replace(year=latest.year - 2))}")
        authorised = st.number_input("Authorised capital (Rs)", value=float(P("authorised_capital", 0.0)), step=100000.0)
        fv = st.number_input("Face value per share (Rs)", value=float(P("face_value", 10.0)))
        bonus = st.number_input("Bonus / split factor after latest year end (1 = none)", value=float(P("bonus_factor", 1.0)))
        tax_rate = st.number_input("Tax rate (e.g. 0.25168 for Sec 115BAA)", value=float(P("tax_rate", 0.25168)), format="%.5f")
        closing_stock = st.number_input("Closing stock at latest year end (Rs) - not in Tally TB", value=float(P("closing_stock", 0.0)))
        st.subheader("Auditor (peer reviewed)")
        aud = P("auditor", {})
        a1, a2 = st.columns(2)
        auditor = {
            "firm": a1.text_input("Firm name", aud.get("firm", "")), "frn": a2.text_input("FRN", aud.get("frn", "")),
            "peer": a1.text_input("Peer review certificate no.", aud.get("peer", "")), "partner": a2.text_input("Partner", aud.get("partner", "")),
            "mno": a1.text_input("Membership no.", aud.get("mno", "")), "udin": a2.text_input("UDIN", aud.get("udin", "")),
            "sign_date": a1.text_input("Date of signing", aud.get("sign_date", "")), "place": a2.text_input("Place", aud.get("place", "")),
        }
    st.subheader("Directors / signatories (first two sign the statements)")
    dir_df = st.data_editor(pd.DataFrame(P("directors", [{"din": "", "name": "", "designation": "Director", "category": "Promoter", "signatory": "Yes"}])),
                            num_rows="dynamic", width="stretch", key=f"dirs{ss.get('_fill_n', 0)}")
    ss.project.update({"name": name, "former": former, "cin": cin, "address": address, "segment": segment, "background": background,
                       "objects": objects, "latest": latest.isoformat(), "authorised_capital": authorised, "face_value": fv,
                       "bonus_factor": bonus, "tax_rate": tax_rate, "closing_stock": closing_stock, "auditor": auditor,
                       "directors": dir_df.fillna("").to_dict("records")})

years = {"d26": latest, "d25": latest.replace(year=latest.year - 1), "d24": latest.replace(year=latest.year - 2)}
YKEY = {"FY26": "d26", "FY25": "d25", "FY24": "d24"}


def load_tb(slot, f):
    """Save an uploaded trial balance for a year slot and read it (raises if it cannot be read)."""
    ext = os.path.splitext(f.name)[1].lower()
    tdir = os.path.join(ss.workdir, "tb_" + slot)
    os.makedirs(tdir, exist_ok=True)
    path = os.path.join(tdir, f"TB_{slot}{ext}")
    with open(path, "wb") as fh:
        fh.write(f.getbuffer())
    ss.tb[slot], ss.tbfile[slot] = tally.read_any_tb(path, tdir)
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
    for slot, _ in SLOTS:
        y = years[YKEY[slot]].year
        pats = [f"{y - 1}-{str(y)[2:]}", f"{str(y - 1)[2:]}-{str(y)[2:]}", f"{y - 1}-{y}", f"fy-?{str(y)[2:]}(?!\\d)", f"mar[a-z]*-?{y}"]
        if any(re.search(r"(?<!\d)" + p_, t) for p_ in pats):
            return slot
    return None


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
    slot = year_slot(up.name) or year_slot(heading_text(up))
    named_tb = bool(re.search(r"trial|\btb\b|tb[-_ ]", up.name, re.I))
    if (slot or named_tb) and up.name.lower().rsplit(".", 1)[-1] in FILE_TYPES:
        slot = slot or next((s_ for s_, _ in SLOTS if s_ not in ss.tb), None)
        if slot:
            before = {k: ss.get(k) for k in ("tbname_" + slot, "tbsrc_" + slot)}, ss.tb.get(slot), ss.tbfile.get(slot)
            with st.spinner(f"Reading {up.name} as trial balance..."):
                try:
                    load_tb(slot, up)
                    op, drs, crs, cl = tally.tb_check(ss.tb[slot])
                    ok = abs(op) < 1 and abs(cl) < 1 and abs(drs - crs) < 1 and sum(1 for x in ss.tb[slot] if x["inc"]) >= 10
                    if named_tb or ok:
                        quick_msg.success(f"{up.name} -> trial balance {fy_label(years[YKEY[slot]])}")
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
    cols = st.columns(3)
    for (slot, lbl), col in zip(SLOTS, cols):
        with col:
            f = st.file_uploader(f"{lbl} - {fy_label(years[YKEY[slot]])}", type=FILE_TYPES, key="tb_" + slot)
            if f is not None and ss.get("tbname_" + slot) != f.name + str(f.size):
                with st.spinner("Reading trial balance..."):
                    try:
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
        st.info("Upload all three trial balances first.")
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
                               "FY26": 0.0, "FY25": 0.0, "FY24": 0.0, "Master if Dr": d, "Master if Cr": c}
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
            disabled=["key", "Primary group", "Group", "Ledger", "FY26", "FY25", "FY24"],
            column_config={"key": None,
                           "Master if Dr": st.column_config.SelectboxColumn(options=options, width="large"),
                           "Master if Cr": st.column_config.SelectboxColumn(options=options, width="large"),
                           "FY26": st.column_config.NumberColumn(fy_label(years["d26"]) + " closing (Dr+/Cr-)", format="%.2f"),
                           "FY25": st.column_config.NumberColumn(fy_label(years["d25"]) + " closing", format="%.2f"),
                           "FY24": st.column_config.NumberColumn(fy_label(years["d24"]) + " closing", format="%.2f")})
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
    sh_cols = {"name": "Name", "s26": fy_label(years["d26"]) + " end", "s25": fy_label(years["d25"]) + " end", "s24": fy_label(years["d24"]) + " end", "promoter": "Promoter"}
    sh_df = pd.DataFrame(P("shareholders", [{"name": "", "s26": 0, "s25": 0, "s24": 0, "promoter": False}]))
    sh_df = st.data_editor(sh_df, num_rows="dynamic", width="stretch", key="sh",
                           column_config={k: st.column_config.Column(v) for k, v in sh_cols.items()} | {"promoter": st.column_config.CheckboxColumn("Promoter")})
    ss.project["shareholders"] = sh_df.fillna(0).to_dict("records")

    def ageing_editor(title, key, buckets, rows_lbl):
        st.subheader(title + " (Rs)")
        store = P(key, {})
        out = {}
        for dk in ["d26", "d25", "d24"]:
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

# ---------------------------------------------------------------- 6. generate
with tabs[5]:
    st.download_button("Save project (.json) for later", json.dumps({"project": ss.project, "overrides": ss.overrides, "regcheck": ss.regcheck,
                                   "documents": {k: {kk: vv for kk, vv in v.items() if kk not in ("text", "path")} for k, v in ss.docs.items()}},
                                  indent=1, default=str),
                       f"{(name or 'rfs').replace(' ', '_')}_project.json")
    problems = []
    if not ready: problems.append("Upload all three trial balances.")
    if not name or not cin: problems.append("Company name and CIN are required.")
    if ready:
        known = set(mapping.MASTER_INFO) | {mapping.EXCLUDE}
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
        cfg["years"] = years
        cfg["adjustments"] = getattr(ss, "entries", [])
        cfg["directors"] = [d for d in cfg.get("directors", []) if d.get("name")]
        cfg["shareholders"] = [h for h in cfg.get("shareholders", []) if h.get("name")]
        run_dir = tempfile.mkdtemp(dir=ss.workdir)
        out = os.path.join(run_dir, f"{(name or 'Company').replace(' ', '_')}_RFS_{fy_label(years['d24'])[3:]}_to_{fy_label(years['d26'])[3:]}.xlsx")
        with st.spinner("Building workbook..."):
            builder.build(cfg, ss.tb, ss.tbfile, out, ss.overrides)
            exporters.add_reference_sheet(out, ss.regcheck, ss.docs,
                                          {fy_label(years[YKEY[s_]]): ss.get("tbsrc_" + s_, os.path.basename(ss.tbfile.get(s_, ""))) for s_, _ in SLOTS})
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
            rows_ = []
            for r in range(9, 60):
                lab = bs[f"B{r}"].value
                if lab and isinstance(bs[f"D{r}"].value, (int, float)) and "Total" in str(lab):
                    rows_.append([lab, bs[f"D{r}"].value, bs[f"E{r}"].value, bs[f"F{r}"].value])
            for r in range(8, 32):
                lab = pl[f"B{r}"].value
                if lab and isinstance(pl[f"D{r}"].value, (int, float)) and any(k in str(lab) for k in ("Revenue", "Profit", "Total")):
                    rows_.append([lab, pl[f"D{r}"].value, pl[f"E{r}"].value, pl[f"F{r}"].value])
            ss.preview = (rows_, [bs[f"{c}1"].value for c in "DEF"])

    if ss.get("preview"):
        rows_, diff = ss.preview
        cols_ = ["Line (Rs. lakhs)", fy_label(years["d26"]), fy_label(years["d25"]), fy_label(years["d24"])]
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
