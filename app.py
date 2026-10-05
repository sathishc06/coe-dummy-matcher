"""COE Dummy Number Matcher - Streamlit application.  Run:  streamlit run app.py"""
import os
import shutil
import tempfile
from collections import Counter

import pandas as pd
import streamlit as st

from coe_matcher import audit as audit_mod
from coe_matcher.corrections import LOG_FIELDS, Override, candidates_for, log_to_csv
from coe_matcher.gate import REL_BLOCKED, REL_COMPLETE, REL_PARTIAL
from coe_matcher.models import (AMBIGUOUS, EXCLUDED, NOT_FOUND, PROC_ERROR, VERIFIED, VERIFIED_ALT)
from coe_matcher.pipeline import reprocess, run_reconciliation

st.set_page_config(page_title="COE Dummy Number Matcher", layout="wide")
PAGES = ["1. Upload", "2. File inventory", "3. Subject mapping", "4. Reconciliation progress", "5. Verified matches",
         "6. Unmatched records", "7. Conflict resolution", "8. Workbook verification", "9. Final download"]
BADGE = {VERIFIED: "🟢 Verified", VERIFIED_ALT: "🟡 Verified alternate source", AMBIGUOUS: "🟠 Ambiguous",
         NOT_FOUND: "🔴 Not found", EXCLUDED: "⚪ Excluded", PROC_ERROR: "⛔ Processing error"}
ADMIN_PW = os.environ.get("COE_ADMIN_PASSWORD", "")      # empty -> manual correction disabled


def S():
    return st.session_state


def save_upload(f, dest):
    with open(dest, "wb") as out:
        shutil.copyfileobj(f, out)
    return dest


def rows_df(run):
    return pd.DataFrame([dict(zip(audit_mod.ROW_COLS, audit_mod.row_values(r))) for r in run.results])


def filters(df):
    c = st.columns(5)
    pick = lambda col, label: c[["Dept", "Subject code(s)", "Sem", "Exam date", "Session"].index(col)].multiselect(
        label, sorted(x for x in df[col].dropna().unique() if x != ""))
    sel = {k: pick(k, k) for k in ["Dept", "Subject code(s)", "Sem", "Exam date", "Session"]}
    for k, v in sel.items():
        if v:
            df = df[df[k].isin(v)]
    return df


def require_run():
    if "run" not in S():
        st.info("Upload files and run the reconciliation first (page 1).")
        st.stop()
    return S().run


page = st.sidebar.radio("Navigation", PAGES)
st.sidebar.caption("Statuses: " + " · ".join(BADGE.values()))

# ------------------------------------------------------------------ 1 upload
if page == PAGES[0]:
    st.title("COE Dummy Number Matcher")
    st.write("Upload the three inputs. The date/session index is auto-detected if it is inside the Daywise ZIP.")
    vz = st.file_uploader("Valuation ZIP", type=["zip"])
    dz = st.file_uploader("Daywise ZIP", type=["zip"])
    ix = st.file_uploader("Official Daywise date/session index (.xlsx) - optional if inside Daywise ZIP", type=["xlsx"])
    excl = st.text_input("Subject codes to EXCLUDE (comma separated)", "225PST02")
    procs = st.slider("Worker processes", 1, 8, 4)
    if st.button("Validate and run reconciliation", type="primary", disabled=not (vz and dz)):
        wd = tempfile.mkdtemp(prefix="coe_")
        paths = dict(v=save_upload(vz, os.path.join(wd, "valuation.zip")), d=save_upload(dz, os.path.join(wd, "daywise.zip")),
                     i=save_upload(ix, os.path.join(wd, "index.xlsx")) if ix else "")
        bar, msg = st.progress(0.0), st.empty()
        try:
            run = run_reconciliation(paths["v"], paths["d"], paths["i"], wd, [c.strip().upper() for c in excl.split(",") if c.strip()],
                                     processes=procs, progress=lambda m, f: (bar.progress(f), msg.write(m)))
            S().run, S().paths, S().procs, S().overrides, S().subject_overrides = run, paths, procs, [], {}
            if not run.val_report.ok or not run.day_report.ok:
                st.error(f"Archive problem: {run.val_report.error or run.day_report.error}")
            else:
                st.success("Reconciliation finished - use the pages on the left.")
        except Exception as e:           # never crash the UI
            st.error(f"Processing error: {e!r}")

# ------------------------------------------------------------------ 2 inventory
elif page == PAGES[1]:
    run = require_run()
    st.header("File inventory")
    c = st.columns(5)
    c[0].metric("Valuation workbooks", len({u.wb_rel for u in run.units} | set(run.wb_problems)))
    c[1].metric("Daywise Excel files", run.catalog.n_xlsx)
    c[2].metric("Daywise CSV files", run.catalog.n_csv)
    c[3].metric("Index entries", len(run.index.entries) if run.index else 0)
    c[4].metric("Lock files skipped", len(run.lock_files))
    st.write("Official index:", run.index_path or "**NOT FOUND**")
    for r in (run.val_report, run.day_report):
        if r.skipped:
            st.warning(f"{r.zip_name}: {len(r.skipped)} unsafe/corrupt member(s) skipped")
            st.write(r.skipped[:50])
    bad = [f for f in run.catalog.files if f.status not in ("ok",)]
    if bad:
        st.warning("Source files not usable (reported, not fatal)")
        st.dataframe(pd.DataFrame([f.__dict__ for f in bad]))
    if run.wb_problems:
        st.warning("Valuation workbooks with problems")
        st.dataframe(pd.DataFrame(sorted(run.wb_problems.items()), columns=["Workbook", "Problem"]))
    if run.index and run.index.warnings:
        with st.expander(f"Index warnings ({len(run.index.warnings)})"):
            st.write(run.index.warnings)
    st.subheader("Daywise source files")
    st.dataframe(pd.DataFrame([f.__dict__ for f in run.catalog.files]), height=300)

# ------------------------------------------------------------------ 3 subjects
elif page == PAGES[2]:
    run = require_run()
    st.header("Subject mapping")
    df = pd.DataFrame([dict(unit_key=s.unit_key, workbook=s.wb_rel, sheet=s.sheet, codes=s.codes, name=s.subject_name,
                            identification=s.id_status, note=s.id_note, crosswalk=s.crosswalk.status, rows=s.n_rows)
                       for s in run.summaries])
    only_issues = st.checkbox("Show only identification/crosswalk issues", True)
    if only_issues:
        df = df[~df.identification.isin(["CONSISTENT", "FILENAME_ONLY", "CONTENT_ONLY"]) | ~df.crosswalk.isin(["EXACT", "COMPOSITE"])]
        df = df[df.rows > 0]
    st.dataframe(df, height=400)
    dup = {}
    for e in (run.index.entries if run.index else []):
        for t in e.tokens:
            dup.setdefault(t, set()).add(e.name)
    d = {k: v for k, v in dup.items() if len(v) > 1}
    if d:
        st.warning("Duplicate subject codes with different subject names in the index: " + str(d))
    st.subheader("Manual subject mapping (authorised users)")
    pw = st.text_input("Authorisation password", type="password", key="pw3")
    if ADMIN_PW and pw == ADMIN_PW:
        uk = st.selectbox("Worksheet", list(df.unit_key) if len(df) else [])
        codes = st.text_input("Official subject code(s), comma separated")
        if st.button("Apply and re-run matching") and uk and codes:
            S().subject_overrides[uk] = [c.strip() for c in codes.split(",")]
            S().run = reprocess(run, S().procs, overrides=S().overrides, subject_overrides=S().subject_overrides)
            st.success("Re-processed with manual subject mapping (flagged as alternate-source).")
    else:
        st.caption("Set COE_ADMIN_PASSWORD on the server to enable manual mapping.")

# ------------------------------------------------------------------ 4 progress
elif page == PAGES[3]:
    run = require_run()
    st.header("Reconciliation progress")
    n = audit_mod.summary_numbers(run)
    c = st.columns(5)
    for col, (k, lab) in zip(c, [("rows", "Rows"), ("verified", "Verified"), ("alt", "Alt source"), ("not_found", "Not found"), ("ambiguous", "Ambiguous")]):
        col.metric(lab, f"{n[k]:,}")
    st.write("Stage timings (s):", run.timings)
    rows = Counter((r.dept or "(unknown)", r.status) for r in run.results)
    st.bar_chart(pd.DataFrame([{"dept": d, "status": s, "rows": v} for (d, s), v in rows.items()]).pivot_table(
        index="dept", columns="status", values="rows", fill_value=0))

# ------------------------------------------------------------------ 5/6 dashboards
elif page in (PAGES[4], PAGES[5]):
    run = require_run()
    df = rows_df(run)
    if page == PAGES[4]:
        st.header("Verified matches")
        df = df[df.Status.isin([VERIFIED, VERIFIED_ALT])]
    else:
        st.header("Unmatched and ambiguous records")
        df = df[df.Status.isin([NOT_FOUND, AMBIGUOUS])]
    df = filters(df)
    df.insert(0, "Indicator", df.Status.map(BADGE))
    st.write(f"{len(df):,} rows")
    st.dataframe(df.head(5000), height=500)
    st.download_button("Download filtered CSV", df.to_csv(index=False).encode("utf-8-sig"), "filtered_rows.csv")
    if page == PAGES[5] and len(df):
        st.subheader("Reasons")
        st.dataframe(df["Reason / note"].str.slice(0, 120).value_counts().head(20))

# ------------------------------------------------------------------ 7 conflicts
elif page == PAGES[6]:
    run = require_run()
    st.header("Conflict resolution (manual corrections)")
    st.caption("Original source files are never modified. Every correction is logged with previous value, new value, "
               "reason, evidence, timestamp and approval note.")
    pw = st.text_input("Authorisation password", type="password", key="pw7")
    if not (ADMIN_PW and pw == ADMIN_PW):
        st.info("Authorised users only (server env COE_ADMIN_PASSWORD).")
        st.stop()
    user = st.text_input("Your name")
    unresolved = [r for r in run.results if r.status in (AMBIGUOUS, NOT_FOUND)]
    st.write(f"{len(unresolved):,} unresolved rows")
    wbsel = st.selectbox("Workbook", sorted({r.wb_rel for r in unresolved}))
    rr_opts = [r for r in unresolved if r.wb_rel == wbsel][:500]
    pick = st.selectbox("Row", rr_opts, format_func=lambda r: f"{r.sheet} row {r.sheet_row} dummy {r.dummy_key} [{r.status}]")
    if pick:
        st.write("Current reason:", pick.reason)
        cands = candidates_for(run, pick)
        if not cands:
            st.error("No source record carries this dummy number - nothing can be selected.")
        else:
            cd = pd.DataFrame(cands)
            st.dataframe(cd)
            i = st.selectbox("Select source record", range(len(cands)),
                             format_func=lambda i: f"{cands[i]['source_file']} #{cands[i]['row']} -> {cands[i]['reg']} ({cands[i]['subject_codes']})")
            reason, ev, note = st.text_area("Reason"), st.text_area("Source evidence"), st.text_area("Approval note")
            if st.button("Record correction and re-run"):
                if not (reason.strip() and ev.strip() and note.strip() and user.strip()):
                    st.error("Reason, evidence, approval note and your name are all required.")
                else:
                    c = cands[i]
                    S().overrides.append(Override(pick.unit_key, pick.sheet_row, c["source_file"], c["row"], reason, ev, note, user))
                    S().run = reprocess(run, S().procs, overrides=S().overrides, subject_overrides=S().subject_overrides)
                    st.success("Correction applied (status: Verified alternate source, flagged MANUAL_OVERRIDE).")
    if getattr(run, "manual_log", None):
        st.subheader("Correction log")
        st.dataframe(pd.DataFrame(run.manual_log)[LOG_FIELDS])
        st.download_button("Download correction log", log_to_csv(run.manual_log).encode("utf-8-sig"), "manual_corrections_log.csv")

# ------------------------------------------------------------------ 8 verification
elif page == PAGES[7]:
    run = require_run()
    st.header("Workbook verification")
    df = pd.DataFrame([dict(workbook=rel, release=st_, reason=why, written=rel in run.written,
                            verification=("PASS" if run.verification[rel]["passed"] else "FAIL") if rel in run.verification else "n/a")
                       for rel, (st_, why) in sorted(run.wb_status.items())])
    f = st.multiselect("Release status", sorted(df.release.unique()), default=[x for x in df.release.unique() if x in (REL_BLOCKED, REL_PARTIAL)])
    st.dataframe(df[df.release.isin(f)] if f else df, height=450)
    rel = st.selectbox("Inspect checks for", sorted(run.verification))
    if rel:
        st.dataframe(pd.DataFrame(run.verification[rel]["checks"], columns=["check", "passed", "detail"]))

# ------------------------------------------------------------------ 9 download
elif page == PAGES[8]:
    run = require_run()
    st.header("Final download")
    gate = run.gate
    if gate["passed"]:
        st.success("RELEASE GATE PASSED")
    else:
        st.error(f"NOT APPROVED FOR OFFICIAL USE - release gate failed ({len(gate['problems'])} workbook(s) with unresolved issues). "
                 "The official ZIP is not generated. Audit-only downloads below are labelled accordingly.")
        st.dataframe(pd.DataFrame(gate["problems"], columns=["workbook", "status", "reason"]).head(300))
    o = run.outputs
    def dl(label, key, mime):
        if key in o and os.path.exists(o[key]):
            st.download_button(label, open(o[key], "rb").read(), os.path.basename(o[key]), mime)
    if gate["passed"]:
        dl("⬇ Official corrected workbooks ZIP", "official_zip", "application/zip")
    else:
        st.warning("Official ZIP blocked by the strict release gate.")
    dl("Audit workbook (FINAL_RECONCILIATION_AUDIT.xlsx)", "audit_xlsx", "application/vnd.ms-excel")
    dl("All student rows CSV", "csv", "text/csv")
    dl("Exception report CSV", "exceptions_csv", "text/csv")
    dl("Verification report", "report", "text/markdown")
    dl("Fully-reconciled workbooks only (audit use" + ("" if gate["passed"] else " - NOT APPROVED FOR OFFICIAL USE") + ")", "complete_zip", "application/zip")
    dl("Partial workbooks (NOT APPROVED FOR OFFICIAL USE)", "partial_zip", "application/zip")
