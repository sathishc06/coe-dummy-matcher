"""INDEPENDENT verification of written workbooks.

Deliberately does NOT reuse the matching/processing functions: it re-reads the original and corrected files with
openpyxl, rebuilds the (dummy -> registration) evidence straight from the raw Daywise files with its own small
reader, and evaluates the marks formulas with its own evaluator.
"""
import csv
import io
import os
import re
from collections import Counter, defaultdict
from typing import Dict, List

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

CODE_RE = re.compile(r"\d{3}[A-Z]{3}\d{2}")


def _nk(v):
    t = str(v or "").upper().replace("&", " AND ")
    return re.sub(r"[^A-Z0-9]", "", t)


def _digits(v):
    if v is None:
        return None
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"[^0-9]", "", s)
    return s.lstrip("0") or None if s else None


def build_independent_evidence(day_root: str, skip_rel=()):
    """code -> dummy_digits -> set(exact reg strings). Reads raw files itself (csv module / openpyxl)."""
    ev = defaultdict(lambda: defaultdict(set))
    skip = set(skip_rel)
    for dp, _, fns in os.walk(day_root):
        for fn in fns:
            rel = os.path.relpath(os.path.join(dp, fn), day_root).replace(os.sep, "/")
            if fn.startswith("~$") or rel in skip or not fn.lower().endswith((".csv", ".xlsx")):
                continue
            ap = os.path.join(dp, fn)
            fcodes = set(CODE_RE.findall(os.path.splitext(fn)[0].upper()))
            try:
                if fn.lower().endswith(".csv"):
                    raw = open(ap, "rb").read()
                    try:
                        txt = raw.decode("utf-8-sig")
                    except UnicodeDecodeError:
                        txt = raw.decode("latin-1")
                    table = list(csv.reader(io.StringIO(txt)))
                else:
                    wb = openpyxl.load_workbook(ap, read_only=True, data_only=True)
                    table = []
                    for ws in wb.worksheets:
                        table += [list(r) for r in ws.iter_rows(values_only=True)]
            except Exception:
                continue
            hdr, last = None, None
            for row in table:
                low = [re.sub(r"[^a-z]", "", str(c or "").lower()) for c in row]
                if "dummyid" in low or "dummyno" in low:
                    hdr = low
                    di = low.index("dummyid") if "dummyid" in low else low.index("dummyno")
                    ri = next((i for i, h in enumerate(low) if h in ("registrationnumber", "regno")), None)
                    ci = next((i for i, h in enumerate(low) if h == "subjectcode"), None)
                    last = None
                    continue
                if hdr is None:
                    # headerless layout [dummy, reg, code, name]
                    if row and len(row) > 2 and _digits(row[0]) and len(str(row[0])) >= 6:
                        hdr, di, ri, ci = [], 0, 1, 2
                    else:
                        continue
                if ri is None or di >= len(row) or ri >= len(row):
                    continue
                d, reg = _digits(row[di]), row[ri]
                if not d or reg in (None, ""):
                    continue
                codes = set()
                if ci is not None and ci < len(row) and str(row[ci] or "").strip():
                    last = str(row[ci])
                if ci is not None and last:
                    codes = set(CODE_RE.findall(last.upper()))
                if not codes:
                    codes = fcodes
                for c in codes:
                    ev[c][d].add(str(reg))
                if not codes:   # no code anywhere: keep under the source's own subject label, for exact-name corroboration
                    for lab in (row[ci] if ci is not None and ci < len(row) else None, row[ci - 1] if ci else None):
                        k = _nk(lab)
                        if k:
                            ev["NAME:" + k][d].add(str(reg))
    return ev


# --------------------------------------------------------------------------------------
_TOK = re.compile(r"^[A-Z]{1,3}\d+$")


def _cells(rng, ws_vals, row):
    a, b = rng.split(":")
    ma, mb = re.match(r"([A-Z]+)(\d+)", a), re.match(r"([A-Z]+)(\d+)", b)
    c1, c2 = column_index_from_string(ma.group(1)), column_index_from_string(mb.group(1))
    return [ws_vals(int(ma.group(2)), c) for c in range(c1, c2 + 1)] if ma.group(2) == mb.group(2) else None


def eval_marks(formula, get):
    """Evaluate the simple marks formulas used in these workbooks. get(row, col)->value. None if unsupported."""
    f = formula.lstrip("=").replace(" ", "")
    num = lambda v: v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0

    def ref(tok):
        m = re.match(r"^([A-Z]+)(\d+)$", tok)
        return num(get(int(m.group(2)), column_index_from_string(m.group(1)))) if m else None

    m = re.fullmatch(r"SUM\((.*)\)", f)
    if m:
        total = 0
        for part in m.group(1).split(","):
            if ":" in part:
                cells = _cells(part, get, None)
                if cells is None:
                    return None
                total += sum(num(v) for v in cells)
            else:
                v = ref(part)
                if v is None:
                    return None
                total += v
        return total
    neg = f.startswith("-")
    body = f[1:] if neg else f
    toks = body.split("+")
    if all(re.match(r"^[A-Z]+\d+$", t) for t in toks):
        s = sum(ref(t) for t in toks)
        return -s if neg else s
    return None


def _grid(ws):
    return {(c.row, c.column): c.value for row in ws.iter_rows() for c in row if c.value is not None}


def verify_workbook(orig_path, out_path, sheets_info, evidence):
    """sheets_info: list of dicts {sheet, header_row, codes, expected_matched(int), n_student_rows}.
    Returns dict(passed: bool, checks: list[(name, ok, detail)], sheet_details: list)."""
    checks = []
    add = lambda n, ok, d="": checks.append((n, bool(ok), d))
    try:
        wo, wn = openpyxl.load_workbook(orig_path), openpyxl.load_workbook(out_path)
        wov = openpyxl.load_workbook(orig_path, data_only=True)
    except Exception as e:
        return dict(passed=False, checks=[("reopen", False, repr(e))])
    add("reopens cleanly", True)
    add("sheet names/order preserved", wo.sheetnames == wn.sheetnames, f"{wo.sheetnames} vs {wn.sheetnames}")
    processed = {s["sheet"]: s for s in sheets_info}
    for name in wo.sheetnames:
        so, sn, sv = wo[name], wn[name], wov[name]
        if name not in processed:
            same = _grid(so) == _grid(sn)
            add(f"[{name}] untouched sheet identical", same)
            continue
        info = processed[name]
        h = info["header_row"]
        G0, G1, V0 = _grid(so), _grid(sn), _grid(sv)
        add(f"[{name}] headers Reg.No/Reg Edit adjacent to Dummy No.",
            [G1.get((h, 1)), G1.get((h, 2)), G1.get((h, 3))] == [G0.get((h, 1)), "Reg.No", "Reg Edit"] and
            G1.get((h, 4)) == G0.get((h, 2)), f"row {h}: {[G1.get((h, c)) for c in range(1, 5)]}")
        maxc0, maxr0 = so.max_column, so.max_row
        add(f"[{name}] row/column extent", sn.max_row == maxr0 and sn.max_column == maxc0 + 2,
            f"rows {maxr0}->{sn.max_row}, cols {maxc0}->{sn.max_column}")
        # student rows
        def students(G, V):
            out = []
            for r in range(h + 1, maxr0 + 1):
                if G.get((r, 1)) not in (None, ""):
                    out.append(r)
            return out
        r0, r1 = students(G0, V0), students(G1, V1 := G1)
        add(f"[{name}] student row count preserved", len(r0) == len(r1) == info["n_student_rows"], f"{len(r0)} vs {len(r1)}")
        # multiset of (dummy, question data) -> no marks moved to another student
        def sig0(r):
            return (str(G0.get((r, 1))), tuple(G0.get((r, c)) for c in range(3, maxc0 + 1) if not str(G0.get((r, c), "")).startswith("=")))

        def sig1(r):
            return (str(G1.get((r, 1))), tuple(G1.get((r, c)) for c in range(5, maxc0 + 3) if not str(G1.get((r, c), "")).startswith("=")))
        c0, c1 = Counter(sig0(r) for r in r0), Counter(sig1(r) for r in r1)
        add(f"[{name}] each dummy keeps its own entered marks (multiset equal)", c0 == c1)
        # no duplicated / deleted rows
        add(f"[{name}] dummy multiset preserved", Counter(str(G0[(r, 1)]) for r in r0) == Counter(str(G1[(r, 1)]) for r in r1))
        # formulas: independent evaluation vs original cached values
        orig_by_sig = defaultdict(list)
        for r in r0:
            f = G0.get((r, 2))
            if isinstance(f, str) and f.startswith("="):
                cached = V0.get((r, 2))
                if cached is None:
                    cached = eval_marks(f, lambda rr, cc, G=G0: G.get((rr, cc)))
                orig_by_sig[sig0(r)].append(cached)
        bad = unev = 0
        used = defaultdict(int)
        for r in r1:
            f = G1.get((r, 4))
            if isinstance(f, str) and f.startswith("="):
                val = eval_marks(f, lambda rr, cc, G=G1: G.get((rr, cc)))
                if val is None:
                    unev += 1
                    continue
                s = sig1(r)
                lst = orig_by_sig.get(s, [])
                exp = lst[used[s]] if used[s] < len(lst) else None
                used[s] += 1
                refs_ok = all(int(m) == r for m in re.findall(r"[A-Z]+(\d+)", f))
                if exp is None or abs(val - exp) > 1e-9 or not refs_ok:
                    bad += 1
        add(f"[{name}] marks formulas point at own row and reproduce original totals", bad == 0, f"{bad} mismatches, {unev} unevaluable")
        # reg checks
        codes = info["codes"]
        wrong, nohyp, blanks, regs = 0, 0, 0, []
        for r in r1:
            reg, edit = G1.get((r, 2)), G1.get((r, 3))
            d = _digits(G1.get((r, 1)))
            if reg in (None, ""):
                blanks += 1
                regs.append(None)
                continue
            regs.append(reg)
            ok = any(str(reg) in evidence.get(c, {}).get(d, set()) for c in codes) or \
                any(str(reg) in evidence.get("NAME:" + _nk(n), {}).get(d, set()) for n in info.get("names", []))
            exp_edit = str(reg).split("-", 1)[1] if "-" in str(reg) else None
            if (not ok) or (edit != exp_edit):
                wrong += 1
        add(f"[{name}] every Reg.No exists in raw source for THIS subject only; Reg Edit = text after first hyphen",
            wrong == 0, f"{wrong} bad of {len(r1) - blanks}")
        add(f"[{name}] number of Reg.No filled equals audit", (len(r1) - blanks) == info["expected_matched"],
            f"{len(r1) - blanks} vs {info['expected_matched']}")
        filled = [x for x in regs if x is not None]
        first_blank = next((i for i, x in enumerate(regs) if x is None), len(regs))
        add(f"[{name}] sorted ascending by full Reg.No (exceptions last)",
            filled == sorted(filled) and all(x is None for x in regs[first_blank:]))
        # non-student rows unchanged (column-shifted)
        stud0, stud1 = set(r0), set(r1)
        diff = 0
        for (r, c), v in G0.items():
            if r in stud0:
                continue
            nc = c if c < 2 else c + 2
            v1 = G1.get((r, nc))
            if isinstance(v, str) and v.startswith("="):
                continue
            if v1 != v:
                diff += 1
        add(f"[{name}] non-student rows/summary cells unchanged", diff == 0, f"{diff} cells differ")
        add(f"[{name}] images preserved", len(so._images) == len(sn._images), f"{len(so._images)} vs {len(sn._images)}")
        add(f"[{name}] merged-cell count preserved", len(so.merged_cells.ranges) == len(sn.merged_cells.ranges))
        add(f"[{name}] freeze panes preserved", so.freeze_panes == sn.freeze_panes)
    return dict(passed=all(ok for _, ok, _ in checks), checks=checks)
