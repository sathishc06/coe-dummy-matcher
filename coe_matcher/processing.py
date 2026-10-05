"""Valuation workbook rewriting: insert Reg.No / Reg Edit after the Dummy column and sort student rows.

Row-relative formulas (e.g. =SUM(C8:Q8)) are rewritten for the shifted columns and the student's NEW row.
Workbooks/sheets with structures that cannot be rewritten safely are blocked, never silently damaged.
"""
import copy
import os
import re
from collections import defaultdict
from typing import Dict, List

import openpyxl
from openpyxl.formatting.rule import Rule  # noqa: F401  (import check)
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string

from .models import MATCHED_STATUSES, RowResult, Unit

REF = re.compile(r"(?<![A-Za-z0-9_$.!'\"])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])")
INSERT_AT, N_NEW = 2, 2
UNMATCHED_FILL = PatternFill("solid", fgColor="FFF2CC")


def shift_formula(f: str, old_row: int = None, new_row: int = None) -> str:
    """Shift column refs >= B by +2; if old_row is given, also renumber refs to that row."""
    parts = re.split(r'("[^"]*")', f)          # keep string literals untouched
    out = []
    for p in parts:
        if p.startswith('"'):
            out.append(p)
            continue

        def sub(m):
            c = column_index_from_string(m.group(2))
            if c >= INSERT_AT:
                c += N_NEW
            r = int(m.group(4))
            if old_row is not None and r == old_row:
                r = new_row
            return f"{m.group(1)}{get_column_letter(c)}{m.group(3)}{r}"
        out.append(REF.sub(sub, p))
    return "".join(out)


def _shift_col(c):
    return c if c < INSERT_AT else c + N_NEW


def _shift_range(rng):
    c1, r1, c2, r2 = rng.min_col, rng.min_row, rng.max_col, rng.max_row
    n1 = c1 if c1 < INSERT_AT else c1 + N_NEW
    n2 = c2 if c2 < INSERT_AT else c2 + N_NEW
    if c1 < INSERT_AT <= c2:      # range spans the insertion point: expand
        n2 = c2 + N_NEW
    return f"{get_column_letter(n1)}{r1}:{get_column_letter(n2)}{r2}"


def _shift_ref_str(s):
    """'$A$1:$R$899' style area -> shifted."""
    def sub(m):
        c = column_index_from_string(m.group(2))
        return f"{m.group(1)}{get_column_letter(_shift_col(c))}{m.group(3)}{m.group(4)}"
    return REF.sub(sub, s)


def sheet_unsupported(ws) -> List[str]:
    why = []
    if ws.conditional_formatting and len(ws.conditional_formatting) > 0:
        why.append("conditional formatting present")
    if ws.data_validations and ws.data_validations.dataValidation:
        why.append("data validation present")
    if ws.tables:
        why.append("Excel table present")
    if ws._hyperlinks:
        why.append("hyperlinks present")
    if ws._charts:
        why.append("chart present")
    if ws.protection and ws.protection.sheet:
        why.append("protected sheet")
    return why


def merge_blockers(ws, unit: Unit) -> List[str]:
    sr = set(unit.student_rows)
    for m in ws.merged_cells.ranges:
        if m.min_row != m.max_row and any(r in sr for r in range(m.min_row, m.max_row + 1)):
            return [f"multi-row merged range {m} overlaps student rows"]
    return []


def process_sheet(ws, unit: Unit, rows: List[RowResult]):
    """Rewrite one worksheet in place. Returns dict describing what was done."""
    by_row = {r.sheet_row: r for r in rows}
    srows = sorted(unit.student_rows)
    matched = [r for r in srows if by_row[r].status in MATCHED_STATUSES]
    unmatched = [r for r in srows if by_row[r].status not in MATCHED_STATUSES]
    matched_sorted = sorted(matched, key=lambda r: (by_row[r].reg_no, r))
    order = matched_sorted + unmatched                      # new slot i receives original row order[i]
    new_pos = {orig: srows[i] for i, orig in enumerate(order)}  # original row -> new row
    maxc, maxr = ws.max_column, ws.max_row
    # snapshot
    snap = {}
    for r in range(1, maxr + 1):
        for c in range(1, maxc + 1):
            cell = ws.cell(r, c)
            snap[(r, c)] = (cell.value, copy.copy(cell._style))
    heights = {r: ws.row_dimensions[r].height for r in range(1, maxr + 1) if r in ws.row_dimensions}
    merges = [copy.copy(m) for m in ws.merged_cells.ranges]
    for m in merges:
        ws.unmerge_cells(str(m))
    # column widths (expand grouped dimensions)
    widths, hidden = {}, set()
    for key, cd in list(ws.column_dimensions.items()):
        lo, hi = (cd.min or column_index_from_string(key)), (cd.max or column_index_from_string(key))
        for ci in range(lo, min(hi, 60) + 1):
            if cd.width:
                widths[ci] = cd.width
            if cd.hidden:
                hidden.add(ci)
    # write cells
    for (r, c), (val, sty) in snap.items():
        nr = new_pos.get(r, r)
        nc = _shift_col(c)
        tgt = ws.cell(nr, nc)
        if isinstance(val, str) and val.startswith("="):
            val = shift_formula(val, r, nr) if r in new_pos else shift_formula(val)
        tgt.value = val
        tgt._style = copy.copy(sty)
    hdr = unit.header_row
    for r in range(1, maxr + 1):
        nr = new_pos.get(r, r)
        a_sty = snap[(r, 1)][1]
        for k in range(N_NEW):
            cell = ws.cell(nr, INSERT_AT + k)
            cell.value = None
            cell._style = copy.copy(a_sty)
    for r, h in heights.items():
        ws.row_dimensions[new_pos.get(r, r)].height = h
    # headers + data
    ws.cell(hdr, 2).value, ws.cell(hdr, 3).value = "Reg.No", "Reg Edit"
    for orig in srows:
        nr = new_pos[orig]
        rr = by_row[orig]
        b, c = ws.cell(nr, 2), ws.cell(nr, 3)
        b.number_format = c.number_format = "@"
        if rr.status in MATCHED_STATUSES:
            b.value = rr.reg_no
            c.value = rr.reg_edit if rr.reg_edit != "" else None
        else:
            b.fill = c.fill = UNMATCHED_FILL
    # merges
    for m in merges:
        rng = _shift_range(m)
        if m.min_row == m.max_row and m.min_row in new_pos:        # merge lives inside one student row: move with it
            nr = new_pos[m.min_row]
            a, b = rng.split(":")
            digits = re.compile(r"\d+")
            rng = digits.sub(str(nr), a) + ":" + digits.sub(str(nr), b)
        ws.merge_cells(rng)
    # widths
    for key in list(ws.column_dimensions.keys()):
        del ws.column_dimensions[key]
    new_w = {_shift_col(ci): w for ci, w in widths.items()}
    new_w[2] = max(widths.get(1, 17.7), 30)
    new_w[3] = max(widths.get(1, 17.7), 24)
    for ci, w in new_w.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    for ci in hidden:
        ws.column_dimensions[get_column_letter(_shift_col(ci))].hidden = True
    # sheet-level references
    if ws.print_area:
        pa = ws.print_area if isinstance(ws.print_area, str) else str(ws.print_area)
        ws.print_area = _shift_ref_str(pa.split("!")[-1].replace("'", ""))
    if ws.auto_filter and ws.auto_filter.ref:
        ws.auto_filter.ref = _shift_ref_str(ws.auto_filter.ref)
    fp = ws.freeze_panes
    if fp:
        m = REF.match(fp)
        if m and column_index_from_string(m.group(2)) >= INSERT_AT:
            ws.freeze_panes = _shift_ref_str(fp)
    return {"order": order, "new_pos": new_pos, "n_matched": len(matched), "n_unmatched": len(unmatched)}


def process_workbook(in_path: str, out_path: str, units: List[Unit], results_by_unit: Dict[str, List[RowResult]]):
    """Returns (ok, message, per_sheet_info). The input file is never modified."""
    info = {}
    try:
        wb = openpyxl.load_workbook(in_path)
    except Exception as e:
        return False, f"cannot open: {e!r}", info
    todo = []
    for u in units:
        rows = results_by_unit.get(f"{u.wb_rel}::{u.sheet}", [])
        if u.excluded or not any(r.status in MATCHED_STATUSES for r in rows):
            continue
        ws = wb[u.sheet]
        bad = list(u.structure_flags) + sheet_unsupported(ws) + merge_blockers(ws, u)
        if bad:
            return False, f"blocked ({u.sheet}): " + "; ".join(bad), info
        todo.append((u, ws, rows))
    if not todo:
        return False, "no verified rows to write", info
    for u, ws, rows in todo:
        info[u.sheet] = process_sheet(ws, u, rows)
    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    return True, "ok", info
