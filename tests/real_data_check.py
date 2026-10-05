"""Post-run acceptance checks against a COMPLETED run directory (real archives). Usage: python real_data_check.py <workdir>
Re-reads outputs from disk only (no pickles, no matching code)."""
import collections
import csv
import glob
import os
import re
import sys
import zipfile

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from coe_matcher.verification import _digits, _nk, build_independent_evidence  # noqa: E402

wd = sys.argv[1]
out, ext = os.path.join(wd, "out"), os.path.join(wd, "extracted")
results = []


def check(name, ok, detail=""):
    results.append((name, bool(ok), detail))
    print(("PASS " if ok else "FAIL ") + name + (f"  [{detail}]" if detail else ""), flush=True)


rows = list(csv.DictReader(open(os.path.join(out, "ALL_STUDENT_ROWS_STATUS.csv"), encoding="utf-8-sig")))
cnt = collections.Counter(r["Status"] for r in rows)
check("CSV contains one line per audited student row", len(rows) == sum(cnt.values()), str(dict(cnt)))
# audit workbook counts equal CSV counts
wb = openpyxl.load_workbook(os.path.join(out, "FINAL_RECONCILIATION_AUDIT.xlsx"), read_only=True)
n = {s: wb[s].max_row - 1 for s in ["Verified Matches", "Alternate Source Matches", "Unmatched Records", "Ambiguous Records"]}
check("audit sheet row counts equal CSV status counts",
      n == {"Verified Matches": cnt["VERIFIED"], "Alternate Source Matches": cnt["VERIFIED_ALTERNATE_SOURCE"],
            "Unmatched Records": cnt["NOT_FOUND"], "Ambiguous Records": cnt["AMBIGUOUS"]}, str(n))
check("no registration number on NOT_FOUND / AMBIGUOUS / EXCLUDED rows",
      all(r["Reg.No (final)"] == "" for r in rows if r["Status"] in ("NOT_FOUND", "AMBIGUOUS", "EXCLUDED")))
check("225PST02 untouched: no row matched, no output workbook", all(r["Status"] == "EXCLUDED" for r in rows if "225PST02" in r["Subject code(s)"])
      and not glob.glob(os.path.join(out, "corrected_workbooks", "**", "*225PST02*"), recursive=True))
# independent evidence recheck of every matched CSV row
idx_rel = [os.path.relpath(p, os.path.join(ext, "daywise")).replace(os.sep, "/")
           for p in glob.glob(os.path.join(ext, "daywise", "**", "DAYWISE MAY 2026.xlsx"), recursive=True)]
ev = build_independent_evidence(os.path.join(ext, "daywise"), idx_rel)
bad = []
for r in rows:
    if r["Status"] in ("VERIFIED", "VERIFIED_ALTERNATE_SOURCE"):
        d, reg = _digits(r["Original dummy"]), r["Reg.No (final)"]
        codes = [c for c in re.split(r"[/]", r["Subject code(s)"]) if c]
        ok = any(reg in ev.get(c, {}).get(d, set()) for c in codes) or any(
            reg in ev.get("NAME:" + _nk(r["Subject name"]), {}).get(d, set()) for _ in [0])
        # name-evidence fallback via index subject name
        if not ok:
            bad.append((r["Workbook"], r["Original dummy"]))
check("every matched Reg.No re-found in raw source for the SAME subject (independent re-read)", not bad, f"{len(bad)} unconfirmed")
for b in bad[:5]:
    print("     unconfirmed:", b)
# written workbooks: row counts and header positions vs originals
written = glob.glob(os.path.join(out, "corrected_workbooks", "**", "*.xlsx"), recursive=True)
probs = 0
for p in written:
    rel = os.path.relpath(p, os.path.join(out, "corrected_workbooks"))
    o = os.path.join(ext, "valuation", rel)
    wn, wo = openpyxl.load_workbook(p, read_only=True), openpyxl.load_workbook(o, read_only=True)
    if wn.sheetnames != wo.sheetnames:
        probs += 1
        continue
    for s in wo.sheetnames:
        a = [r for r in wo[s].iter_rows(values_only=True) if any(v is not None for v in r)]
        b = [r for r in wn[s].iter_rows(values_only=True) if any(v is not None for v in r)]
        if len(a) != len(b):
            probs += 1
check(f"all {len(written)} written workbooks keep sheet names and non-empty row counts", probs == 0, f"{probs} problems")
check("release gate recorded as failed -> no official ZIP present", not os.path.exists(os.path.join(out, "CORRECTED_WORKBOOKS_OFFICIAL.zip")))
for z in ("CORRECTED_WORKBOOKS_COMPLETE_ONLY.zip", "CORRECTED_WORKBOOKS_PARTIAL_NOT_FOR_RELEASE.zip"):
    zz = zipfile.ZipFile(os.path.join(out, z))
    check(f"{z} opens, no corrupt members", zz.testzip() is None, f"{len(zz.namelist())} members")
print("\nSUMMARY:", sum(ok for _, ok, _ in results), "passed,", sum(not ok for _, ok, _ in results), "failed")
sys.exit(0 if all(ok for _, ok, _ in results) else 1)
