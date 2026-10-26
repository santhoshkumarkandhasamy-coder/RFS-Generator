"""RFS Generator - Streamlit app.

Upload the three Tally trial balances, fill in company details, review the ledger mapping,
enter restatement adjustments, and download the Restated Financial Statements workbook.
Run:  streamlit run app.py
"""
import datetime as dt
import json
import os
import tempfile

import pandas as pd
import streamlit as st

from rfs import tally, builder, mapping, exporters

st.set_page_config(page_title="RFS Generator", layout="wide")
SLOTS = [("FY26", "Latest year"), ("FY25", "Previous year"), ("FY24", "Earliest year")]

ss = st.session_state
ss.setdefault("tb", {})          # slot -> parsed rows
ss.setdefault("tbfile", {})      # slot -> saved path
ss.setdefault("workdir", tempfile.mkdtemp(prefix="rfs_"))
ss.setdefault("overrides", {})
ss.setdefault("project", {})


def fy_label(d):
    return f"FY {d.year - 1}-{str(d.year)[2:]}"


def P(key, default=None):
    return ss.project.get(key, default)


# ---------------------------------------------------------------- sidebar: project file
with st.sidebar:
    st.header("RFS Generator")
    st.caption("SEBI ICDR - Indian GAAP - Schedule III Division I")
    up = st.file_uploader("Reload a saved project (optional, .json)", type="json", key="projup",
                          help="Only for re-opening a project saved from the Generate tab. Trial balances are uploaded in tab 2.")
    if up is not None and ss.get("_loaded") != up.name:
        proj = json.load(up)
        ss.project = proj.get("project", {})
        ss.overrides = proj.get("overrides", {})
        ss._loaded = up.name
        st.success("Project loaded. Upload the trial balances again.")
    st.markdown("**Steps**\n1. Company\n2. Trial balances\n3. Mapping\n4. Shareholders & ageing\n5. Adjustments\n6. Generate")

tabs = st.tabs(["1. Company", "2. Trial balances", "3. Mapping review", "4. Shareholders & ageing", "5. Adjustments", "6. Generate"])

# ---------------------------------------------------------------- 1. company
with tabs[0]:
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
                            num_rows="dynamic", width="stretch", key="dirs")
    ss.project.update({"name": name, "former": former, "cin": cin, "address": address, "segment": segment, "background": background,
                       "objects": objects, "latest": latest.isoformat(), "authorised_capital": authorised, "face_value": fv,
                       "bonus_factor": bonus, "tax_rate": tax_rate, "closing_stock": closing_stock, "auditor": auditor,
                       "directors": dir_df.fillna("").to_dict("records")})

years = {"d26": latest, "d25": latest.replace(year=latest.year - 1), "d24": latest.replace(year=latest.year - 2)}
YKEY = {"FY26": "d26", "FY25": "d25", "FY24": "d24"}

# ---------------------------------------------------------------- 2. trial balances
with tabs[1]:
    st.write("Upload the Tally trial balance exports (Excel, Detailed: opening, debit, credit, closing with Dr/Cr). "
             "Ledgers must be in italics (default Tally export).")
    cols = st.columns(3)
    for (slot, lbl), col in zip(SLOTS, cols):
        with col:
            f = st.file_uploader(f"{lbl} - {fy_label(years[YKEY[slot]])}", type=["xlsx"], key="tb_" + slot)
            if f is not None and ss.get("tbname_" + slot) != f.name + str(f.size):
                path = os.path.join(ss.workdir, f"TB_{slot}.xlsx")
                with open(path, "wb") as fh:
                    fh.write(f.getbuffer())
                with st.spinner("Reading trial balance..."):
                    ss.tb[slot] = tally.read_tb(path)
                ss.tbfile[slot] = path
                ss["tbname_" + slot] = f.name + str(f.size)
            if slot in ss.tb:
                op, drs, crs, cl = tally.tb_check(ss.tb[slot])
                n = sum(1 for x in ss.tb[slot] if x["inc"])
                ok = abs(op) < 1 and abs(cl) < 1 and abs(drs - crs) < 1
                (st.success if ok else st.error)(f"{n} ledgers | Dr {drs:,.2f} | Cr {crs:,.2f} | closing net {cl:,.2f}")

ready = all(s in ss.tb for s, _ in SLOTS)

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
    st.download_button("Save project (.json) for later", json.dumps({"project": ss.project, "overrides": ss.overrides}, indent=1, default=str),
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
