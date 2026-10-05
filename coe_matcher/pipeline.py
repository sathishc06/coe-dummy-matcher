"""End-to-end orchestration: extract -> index -> scan -> catalog -> match -> write -> verify -> audit."""
import glob
import os
import shutil
import time
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from multiprocessing import Pool
from typing import Callable, Dict, List, Optional

from . import audit as audit_mod
from .archive import ArchiveReport, is_junk, is_lock_file, safe_extract
from .gate import (REL_BLOCKED, REL_COMPLETE, REL_EXCL, REL_NA, REL_NONE, REL_PARTIAL, evaluate_gate,
                   workbook_release_status)
from .index import IndexData, load_index
from .matching import EXCLUDED_DEFAULT, match_unit, unit_key
from .models import MATCHED_STATUSES, PROC_ERROR, RowResult, Unit
from .processing import process_workbook
from .sources import SourceCatalog, build_catalog
from .subjects import build_code_lookup
from .valuation import scan_workbook
from .verification import build_independent_evidence, verify_workbook

_G = {}   # worker globals (inherited on fork)


@dataclass
class Run:
    workdir: str
    val_root: str = ""
    day_root: str = ""
    index_path: str = ""
    val_report: Optional[ArchiveReport] = None
    day_report: Optional[ArchiveReport] = None
    index: Optional[IndexData] = None
    catalog: Optional[SourceCatalog] = None
    units: List[Unit] = field(default_factory=list)
    wb_problems: Dict[str, str] = field(default_factory=dict)     # rel -> problem (unreadable etc.)
    lock_files: List[str] = field(default_factory=list)
    results: List[RowResult] = field(default_factory=list)
    summaries: list = field(default_factory=list)
    wb_status: Dict[str, tuple] = field(default_factory=dict)
    verification: Dict[str, dict] = field(default_factory=dict)
    written: Dict[str, str] = field(default_factory=dict)         # rel -> output path
    write_msgs: Dict[str, str] = field(default_factory=dict)
    gate: dict = field(default_factory=dict)
    excluded_codes: tuple = EXCLUDED_DEFAULT
    outputs: Dict[str, str] = field(default_factory=dict)
    timings: Dict[str, float] = field(default_factory=dict)
    log: List[str] = field(default_factory=list)


def find_index_file(roots: List[str], explicit: str = ""):
    if explicit and os.path.exists(explicit):
        return explicit
    from .index import find_index_sheet
    import openpyxl
    cands = []
    for root in roots:
        for dp, _, fns in os.walk(root):
            for fn in fns:
                if fn.lower().endswith(".xlsx") and not fn.startswith("~$"):
                    if "daywise" in fn.lower() or "index" in fn.lower() or "schedule" in fn.lower() or "time" in fn.lower():
                        cands.append(os.path.join(dp, fn))
    for c in cands:
        try:
            wb = openpyxl.load_workbook(c, read_only=True)
            ws, h, hs = find_index_sheet(wb)
            if ws is not None:
                return c
        except Exception:
            pass
    return ""


def _scan_one(args):
    ap, rel = args
    if is_lock_file(rel):
        return rel, [], ["Office temporary lock file (~$) - not a workbook"]
    return (rel,) + scan_workbook(ap, rel)


def _write_one(args):
    rel, ap, out, units, rb = args
    ok, msg, info = process_workbook(ap, out, units, rb)
    return rel, ok, msg, {k: {"n_matched": v["n_matched"], "n_unmatched": v["n_unmatched"]} for k, v in info.items()}


def _init_verify_worker(evidence):
    # Streamlit Cloud can use a multiprocessing start method where module globals
    # are not inherited.  Initialise the worker explicitly so verification never
    # depends on _G["evidence"] already existing in the child process.
    _G["evidence"] = evidence


def _verify_one(args):
    rel, ap, out, sheets_info = args
    return rel, verify_workbook(ap, out, sheets_info, _G.get("evidence", {}))


def run_reconciliation(val_zip: str, day_zip: str, index_xlsx: str, workdir: str, excluded=EXCLUDED_DEFAULT,
                       processes: int = 4, progress: Callable[[str, float], None] = None, overrides=None,
                       subject_overrides=None) -> Run:
    P = progress or (lambda m, f: None)
    run = Run(workdir=workdir, excluded_codes=tuple(excluded))
    T = time.time()
    tick = lambda k: run.timings.__setitem__(k, round(time.time() - T, 1))
    ex = os.path.join(workdir, "extracted")
    P("Extracting valuation archive", 0.02)
    run.val_report = safe_extract(val_zip, os.path.join(ex, "valuation"))
    P("Extracting Daywise archive", 0.06)
    run.day_report = safe_extract(day_zip, os.path.join(ex, "daywise"))
    run.val_root, run.day_root = os.path.join(ex, "valuation"), os.path.join(ex, "daywise")
    if not (run.val_report.ok and run.day_report.ok):
        run.log.append("archive extraction failed")
        return run
    tick("extract")
    # --- index --------------------------------------------------------------------------
    run.index_path = find_index_file([run.day_root, run.val_root], index_xlsx)
    index_rel_skip = []
    if run.index_path:
        run.index = load_index(run.index_path)
        if run.index_path.startswith(run.day_root):
            index_rel_skip.append(os.path.relpath(run.index_path, run.day_root).replace(os.sep, "/"))
    # --- scan valuation workbooks --------------------------------------------------------
    P("Scanning valuation workbooks", 0.10)
    vfiles = []
    for dp, _, fns in os.walk(run.val_root):
        for fn in sorted(fns):
            ap = os.path.join(dp, fn)
            rel = os.path.relpath(ap, run.val_root).replace(os.sep, "/")
            if is_junk(rel):
                continue
            if fn.lower().endswith((".xlsx", ".xlsm")):
                vfiles.append((ap, rel))
    if processes > 1:
        with Pool(processes) as pool:
            scans = pool.map(_scan_one, vfiles, chunksize=4)
    else:
        scans = [_scan_one(a) for a in vfiles]
    for rel, units, problems in scans:
        if rel and is_lock_file(rel):
            run.lock_files.append(rel)
            continue
        if problems and not units:
            run.wb_problems[rel] = "; ".join(problems)
        run.units += units
    tick("scan")
    # --- source catalog -------------------------------------------------------------------
    P("Parsing Daywise sources", 0.30)
    run.catalog = build_catalog(run.day_root, run.index, skip_rel=index_rel_skip)
    tick("catalog")
    return reprocess(run, processes, P, overrides, subject_overrides, T)


def reprocess(run: "Run", processes: int = 4, progress=None, overrides=None, subject_overrides=None, T0=None) -> "Run":
    """Match -> write -> verify -> gate -> outputs. Re-runnable after manual corrections (no re-extraction)."""
    P = progress or (lambda m, f: None)
    T = T0 or time.time()
    tick = lambda k: run.timings.__setitem__(k, round(time.time() - T, 1))
    run.results, run.summaries, run.wb_status, run.verification, run.written, run.write_msgs = [], [], {}, {}, {}, {}
    run.manual_log = []
    index_rel_skip = []
    if run.index_path and run.index_path.startswith(run.day_root):
        index_rel_skip.append(os.path.relpath(run.index_path, run.day_root).replace(os.sep, "/"))
    out_root = os.path.join(run.workdir, "out")
    shutil.rmtree(out_root, ignore_errors=True)
    # --- match -----------------------------------------------------------------------------
    P("Matching dummy numbers (subject + session scoped)", 0.50)
    lookup = build_code_lookup(run.index)
    for u in run.units:
        if subject_overrides and unit_key(u) in subject_overrides:
            u.manual_codes = [c.strip().upper() for c in subject_overrides[unit_key(u)] if c.strip()]
        rows, summ = match_unit(u, run.index, run.catalog, run.excluded_codes, lookup)
        run.results += rows
        run.summaries.append(summ)
    if overrides:
        from .corrections import apply_overrides
        apply_overrides(run, overrides)
    tick("match")
    # --- write -----------------------------------------------------------------------------
    P("Writing corrected workbooks", 0.62)
    out_dir = os.path.join(run.workdir, "out", "corrected_workbooks")
    rb = defaultdict(list)
    for r in run.results:
        rb[r.unit_key].append(r)
    by_wb = defaultdict(list)
    for u in run.units:
        by_wb[u.wb_rel].append(u)
    jobs = []
    for rel, us in by_wb.items():
        if any(r.status in MATCHED_STATUSES for u in us for r in rb[unit_key(u)]):
            jobs.append((rel, us[0].wb_abs, os.path.join(out_dir, rel), us, {unit_key(u): rb[unit_key(u)] for u in us}))
    if processes > 1:
        with Pool(processes) as pool:
            wres = pool.map(_write_one, jobs, chunksize=2)
    else:
        wres = [_write_one(j) for j in jobs]
    sheet_written = {}
    for rel, ok, msg, info in wres:
        run.write_msgs[rel] = msg
        if ok:
            run.written[rel] = os.path.join(out_dir, rel)
            sheet_written[rel] = info
    tick("write")
    # --- independent verification ---------------------------------------------------------
    P("Independently verifying corrected workbooks", 0.78)
    _G["evidence"] = build_independent_evidence(run.day_root, index_rel_skip)
    vjobs = []
    for rel, outp in run.written.items():
        sheets_info = []
        for u in by_wb[rel]:
            if u.sheet in sheet_written[rel]:
                rows = rb[unit_key(u)]
                nm = {e.name for sm in run.summaries if sm.unit_key == unit_key(u) for e in sm.crosswalk.entries}
                sheets_info.append(dict(sheet=u.sheet, header_row=u.header_row, codes=list(u.codes), names=sorted(nm),
                                        expected_matched=sum(1 for r in rows if r.status in MATCHED_STATUSES),
                                        n_student_rows=len(u.student_rows)))
        vjobs.append((rel, by_wb[rel][0].wb_abs, outp, sheets_info))
    if processes > 1 and vjobs:
        # Explicitly initialise each verification worker with the evidence map.
        # Without this, Streamlit Cloud/Python multiprocessing may start a fresh
        # interpreter and _G["evidence"] is absent, producing KeyError('evidence').
        with Pool(processes, initializer=_init_verify_worker, initargs=(_G["evidence"],)) as pool:
            vres = pool.map(_verify_one, vjobs, chunksize=2)
    else:
        vres = [_verify_one(j) for j in vjobs]
    run.verification = dict(vres)
    for rel, v in run.verification.items():
        for r in run.results:
            pass
    tick("verify")
    # --- classify + gate ----------------------------------------------------------------------
    for rel, us in by_wb.items():
        urows = {unit_key(u): rb[unit_key(u)] for u in us}
        blocked = ""
        if rel in run.written or rel not in run.write_msgs:
            pass
        if rel in run.write_msgs and rel not in run.written:
            blocked = run.write_msgs[rel] if run.write_msgs[rel].startswith("blocked") or "cannot" in run.write_msgs[rel] else ""
        vok = run.verification.get(rel, {}).get("passed", False)
        st = workbook_release_status(urows, None, rel in run.written, vok, blocked)
        run.wb_status[rel] = st
    for rel, prob in run.wb_problems.items():
        run.wb_status[rel] = (REL_BLOCKED, prob)
    # propagate verification result into row records
    for r in run.results:
        v = run.verification.get(r.wb_rel)
        if v is not None:
            r.verification = "PASS" if v["passed"] else "FAIL"
        elif r.status == "EXCLUDED":
            r.verification = "N/A (excluded)"
        else:
            r.verification = "N/A (no output workbook)"
    run.gate = evaluate_gate({k: v for k, v in run.wb_status.items() if v[0] != REL_EXCL and v[0] != REL_NA})
    # --- outputs -----------------------------------------------------------------------------------
    P("Writing audit, CSV and release packages", 0.92)
    run.outputs = audit_mod.write_all_outputs(run, os.path.join(run.workdir, "out"))
    tick("outputs")
    P("Done", 1.0)
    return run
