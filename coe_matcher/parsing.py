"""Parsers for Daywise source files (CSV / Excel) -> unified record list.
Original registration numbers are preserved byte-for-byte (only str() of the cell)."""
import csv
import io
import os
from dataclasses import dataclass, field
from typing import List, Optional

import openpyxl

from .normalize import clean_text, code_key

DUMMY_HDRS = {"dummyid", "dummyno", "dummynumber", "dummy"}
REG_HDRS = {"registrationnumber", "regno", "registerno", "registernumber", "regnumber", "registrationno"}
CODE_HDRS = {"subjectcode", "subcode"}
NAME_HDRS = {"subjectname", "subname"}
SEQ_HDRS = {"sequence", "seq"}
VER_HDRS = {"version"}
DATE_HDRS = {"generateddate"}


def _h(v):
    return "".join(ch for ch in clean_text(v).lower() if ch.isalnum())


@dataclass
class SourceRow:
    source_file: str          # path relative to the archive root
    sheet: str                # sheet name or "CSV"
    row_no: int               # 1-based row number in the source file
    dummy_raw: object
    reg_raw: object           # original, untouched value
    subject_code_raw: Optional[str]
    subject_name_raw: Optional[str]
    version: Optional[str] = None
    code_inherited: bool = False   # True if code was blank and inherited from the row above in the same sheet
    generated: Optional[str] = None


@dataclass
class SourceFileInfo:
    path: str
    kind: str                 # csv | xlsx
    status: str               # ok | skipped_lockfile | unreadable | no_recognised_columns
    detail: str = ""
    n_rows: int = 0
    layout: str = ""


def is_lock_file(path):
    return os.path.basename(path).startswith("~$")


def _read_csv_rows(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    return list(csv.reader(io.StringIO(text)))


def _map_header(row):
    m = {}
    for i, v in enumerate(row):
        k = _h(v)
        if k in DUMMY_HDRS and "dummy" not in m: m["dummy"] = i
        elif k in REG_HDRS and "reg" not in m: m["reg"] = i
        elif k in CODE_HDRS and "code" not in m: m["code"] = i
        elif k in NAME_HDRS and "name" not in m: m["name"] = i
        elif k in VER_HDRS and "version" not in m: m["version"] = i
        elif k in DATE_HDRS and "gen" not in m: m["gen"] = i
    return m if ("dummy" in m and "reg" in m) else None


def _looks_headerless(row):
    """Headerless layout seen in May-2026 data: [dummy(int-like), regno, code, name]."""
    if len(row) < 2:
        return False
    d = str(row[0]).strip().replace(".0", "")
    return d.isdigit() and len(d) >= 6 and row[1] not in (None, "")


def _cell(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _emit(rows, rel, sheet, start_row, m, info):
    out = []
    last_code = None
    last_name = None
    for off, row in enumerate(rows):
        rn = start_row + off
        get = lambda k: (row[m[k]] if k in m and m[k] < len(row) else None)
        d, r = get("dummy"), get("reg")
        if d in (None, "") and r in (None, ""):
            continue
        code, name = get("code"), get("name")
        inherited = False
        if clean_text(code) == "" and last_code is not None and "code" in m:
            code, inherited = last_code, True
        elif clean_text(code):
            last_code = code
        if clean_text(name) == "" and last_name is not None and "name" in m:
            name = last_name
        elif clean_text(name):
            last_name = name
        out.append(SourceRow(rel, sheet, rn, d, r, _cell(code), _cell(name),
                             version=_cell(get("version")), code_inherited=inherited,
                             generated=_cell(get("gen"))))
    info.n_rows += len(out)
    return out


def parse_source_file(abs_path, rel_path):
    """Returns (list[SourceRow], SourceFileInfo)."""
    ext = os.path.splitext(abs_path)[1].lower()
    kind = "csv" if ext == ".csv" else "xlsx"
    info = SourceFileInfo(rel_path, kind, "ok")
    if is_lock_file(abs_path):
        info.status, info.detail = "skipped_lockfile", "Office temporary lock file (~$), not a real workbook"
        return [], info
    rows_out = []
    try:
        if kind == "csv":
            allrows = _read_csv_rows(abs_path)
            if not allrows:
                info.status, info.detail = "no_recognised_columns", "empty file"
                return [], info
            m = _map_header(allrows[0])
            if m:
                info.layout = "csv-header"
                rows_out += _emit(allrows[1:], rel_path, "CSV", 2, m, info)
            else:
                info.status, info.detail = "no_recognised_columns", f"first row: {allrows[0][:6]}"
        else:
            wb = openpyxl.load_workbook(abs_path, read_only=True, data_only=True)
            for ws in wb.worksheets:
                allrows = [list(r) for r in ws.iter_rows(values_only=True)]
                # locate header within the first 10 rows
                hdr_i, m = None, None
                for i, r in enumerate(allrows[:10]):
                    mm = _map_header(r)
                    if mm:
                        hdr_i, m = i, mm
                        break
                if m:
                    info.layout = "xlsx-header"
                    rows_out += _emit(allrows[hdr_i + 1:], rel_path, ws.title, hdr_i + 2, m, info)
                elif allrows and _looks_headerless(allrows[0]):
                    info.layout = "xlsx-headerless[dummy,reg,code,name]"
                    mm = {"dummy": 0, "reg": 1, "code": 2, "name": 3}
                    rows_out += _emit(allrows, rel_path, ws.title, 1, mm, info)
            if not rows_out:
                info.status, info.detail = "no_recognised_columns", "no sheet with dummy+registration columns"
    except Exception as e:  # corrupt / unsupported: report, never crash
        info.status, info.detail = "unreadable", repr(e)[:200]
        return [], info
    return rows_out, info
