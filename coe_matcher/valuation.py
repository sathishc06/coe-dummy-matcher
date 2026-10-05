"""Valuation workbook scanning: locate student tables, identify subjects, flag unsupported structures."""
import os
import re
from typing import List, Tuple

import openpyxl

from .models import Unit
from .normalize import clean_text, dummy_key, extract_codes

CODE_ANYWHERE = re.compile(r"\d{3}[A-Za-z]{3}\d{2}")
CELL_REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def codes_in(text):
    return list(dict.fromkeys(m.upper() for m in CODE_ANYWHERE.findall(clean_text(text))))


def _subject_name_from_c6(c6, codes):
    t = clean_text(c6)
    for c in codes:
        t = re.sub(c, " ", t, flags=re.I)
    t = re.sub(r"^[\s&/\-,:.]+", "", t)
    t = re.sub(r"^(OR|AND)\s+", "", t, flags=re.I)
    return t.strip(" &/-,:.")


def formula_issues(formula, own_row):
    """Return a reason string if the formula cannot be safely moved by a pure column shift / row sort."""
    f = formula
    if "!" in f:
        return "cross-sheet reference"
    if "$" in f:
        return "absolute reference"
    if "[" in f:
        return "external-link reference"
    for m in CELL_REF.finditer(re.sub(r'"[^"]*"', "", f)):
        if int(m.group(4)) != own_row:
            return f"reference to another row ({m.group(0)})"
    return ""


def scan_workbook(abs_path: str, rel_path: str) -> Tuple[List[Unit], List[str]]:
    """Returns (units, workbook_level_problems). Never raises."""
    problems: List[str] = []
    units: List[Unit] = []
    fname = os.path.basename(rel_path)
    ext = os.path.splitext(fname)[1].lower()
    if ext == ".xlsm":
        problems.append("macro-enabled workbook (.xlsm) is not supported for rewriting")
    try:
        wb = openpyxl.load_workbook(abs_path, read_only=True, data_only=False)
    except Exception as e:
        return [], [f"cannot open workbook: {e!r}"]
    file_codes = codes_in(os.path.splitext(fname)[0])
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hdr, dcol = None, None
        for i, r in enumerate(rows[:15], 1):
            for j, v in enumerate((r or ())[:4], 1):
                if clean_text(v).lower().startswith("dummy"):
                    hdr, dcol = i, j
                    break
            if hdr:
                break
        if not hdr:
            continue
        u = Unit(rel_path, abs_path, ws.title, header_row=hdr, dummy_col=dcol,
                 max_col=max((len(r) for r in rows), default=0), max_row=len(rows), n_sheets=len(wb.worksheets))
        g = lambda r, c: (rows[r - 1][c - 1] if r - 1 < len(rows) and c - 1 < len(rows[r - 1]) else None)
        u.c5, u.c6, u.q5, u.q6 = (clean_text(g(5, 3)), clean_text(g(6, 3)), clean_text(g(5, 17)), clean_text(g(6, 17)))
        if not u.c6:  # layouts with shifted label columns
            for c in range(2, 8):
                t = clean_text(g(6, c))
                if t:
                    u.c6 = t
                    break
        for i in range(hdr + 1, len(rows) + 1):
            v = g(i, dcol)
            if clean_text(v) == "":
                continue
            u.student_rows.append(i)
            u.dummy_raw[i] = v
            sig = tuple(x for j, x in enumerate(rows[i - 1] or (), 1)
                        if j != dcol and not (isinstance(x, str) and x.startswith("=")))
            while sig and sig[-1] in (None, ""):
                sig = sig[:-1]
            u.row_sig[i] = sig
        # structure checks -------------------------------------------------------
        bad_f = set()
        for i in range(hdr + 1, len(rows) + 1):
            for j, v in enumerate(rows[i - 1] or (), 1):
                if isinstance(v, str) and v.startswith("="):
                    why = formula_issues(v, i)
                    if why:
                        bad_f.add(why)
                elif v is not None and not isinstance(v, (str, int, float, bool)) and hasattr(v, "text"):
                    bad_f.add("array formula")
        for why in sorted(bad_f):
            u.structure_flags.append(f"unsupported formula structure: {why}")
        if ext == ".xlsm":
            u.structure_flags.append("macro-enabled workbook")
        # codes ------------------------------------------------------------------
        u.file_codes = file_codes
        u.content_codes = list(dict.fromkeys(codes_in(u.c6) + codes_in(ws.title)))
        fset, cset = set(u.file_codes), set(u.content_codes)
        if fset and cset:
            if fset & cset:
                u.id_status = "CONSISTENT"
                u.codes = sorted(fset | cset) if len(fset | cset) <= 3 else sorted(fset & cset)
                if (fset ^ cset):
                    u.id_note = f"filename codes {sorted(fset)} vs content codes {sorted(cset)}: overlap used as primary"
                    u.codes = sorted(fset & cset) if (fset & cset) else sorted(fset | cset)
            else:
                u.id_status = "CONFLICT"
                u.codes = sorted(fset | cset)
                u.id_note = f"filename code(s) {sorted(fset)} disagree with workbook content code(s) {sorted(cset)}"
        elif fset:
            u.id_status, u.codes = "FILENAME_ONLY", sorted(fset)
        elif cset:
            u.id_status, u.codes = "CONTENT_ONLY", sorted(cset)
        else:
            u.id_status = "UNIDENTIFIED"
        u.subject_name = _subject_name_from_c6(u.c6, u.content_codes or u.file_codes)
        if not u.student_rows:
            u.kind = "BLANK_TEMPLATE" if not u.codes else "NO_STUDENT_ROWS"
        units.append(u)
    if not units:
        problems.append("no worksheet with a 'Dummy No.' header was found")
    return units, problems
