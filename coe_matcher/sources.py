"""Source catalog: parsed Daywise rows annotated with subject evidence and folder date/session."""
import datetime as dt
import os
import re
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from .archive import is_junk, is_lock_file
from .index import IndexData, norm_session, split_tokens
from .normalize import clean_text, dummy_key, name_key
from .parsing import SourceFileInfo, SourceRow, parse_source_file

DATE_IN_PATH = re.compile(r"(?<!\d)(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})(?!\d)(?:\s*(FN|AN|F\.N|A\.N))?", re.I)
SCI = re.compile(r"^\s*[+-]?\d+(\.\d+)?E[+-]?\d+\s*$", re.I)
HIDDEN = re.compile(r"[\u200b-\u200f\u2060\ufeff\u00a0]")


@dataclass
class SrcRec:
    row: SourceRow
    dkey: str
    reg: str                       # exact original string
    reg_key: str                   # comparison key (whitespace-normalised)
    codes: List[str]               # official-format subject tokens tied to this record
    evidence: str                  # code | filename | name | none
    folder_date: Optional[str]
    folder_session: str
    in_all_folder: bool
    corrupt: str = ""              # reason the registration value is unusable
    flags: List[str] = field(default_factory=list)


def path_date_session(rel_path):
    """Deepest dd.mm.yy[yy] [FN|AN] component in a path (folder or file name). Returns (iso, session)."""
    best = (None, "")
    for comp in rel_path.replace("\\", "/").split("/"):
        for m in DATE_IN_PATH.finditer(comp):
            d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            y = y + 2000 if y < 100 else y
            try:
                iso = dt.date(y, mo, d).isoformat()
            except ValueError:
                continue
            best = (iso, norm_session(m.group(4) or ""))
    return best


def reg_problem(raw):
    s = "" if raw is None else str(raw)
    if clean_text(s) == "":
        return "registration number is blank"
    if SCI.match(s):
        return f"registration number '{s}' was corrupted into scientific notation in the source file"
    return ""


class SourceCatalog:
    def __init__(self):
        self.recs: List[SrcRec] = []
        self.files: List[SourceFileInfo] = []
        self.by_code: Dict[str, List[SrcRec]] = defaultdict(list)
        self.by_dummy: Dict[str, List[SrcRec]] = defaultdict(list)
        self.name_assigned_files: Dict[str, List[str]] = {}

    @property
    def n_csv(self):
        return sum(1 for f in self.files if f.kind == "csv")

    @property
    def n_xlsx(self):
        return sum(1 for f in self.files if f.kind == "xlsx")


def build_catalog(root: str, index: Optional[IndexData] = None, skip_rel=()) -> SourceCatalog:
    cat = SourceCatalog()
    skip = {s.replace("\\", "/") for s in skip_rel}
    names = defaultdict(list)       # name_key -> index entries (official name evidence)
    if index:
        for e in index.entries:
            names[e.name_key].append(e)
    for dirpath, _, fns in os.walk(root):
        for fn in sorted(fns):
            ap = os.path.join(dirpath, fn)
            rel = os.path.relpath(ap, root).replace(os.sep, "/")
            if rel in skip or is_junk(rel) or not fn.lower().endswith((".csv", ".xlsx", ".xlsm")):
                continue
            rows, info = parse_source_file(ap, rel)
            cat.files.append(info)
            fdate, fsess = path_date_session(rel)
            in_all = any(c.upper() == "ALL" for c in rel.split("/")[:-1])
            fname_codes = list(dict.fromkeys(m.upper() for m in re.findall(r"\d{3}[A-Za-z]{3}\d{2}", os.path.splitext(fn)[0])))
            for r in rows:
                toks, bad = split_tokens(r.subject_code_raw or "")
                flags = []
                if toks:
                    codes, ev = toks, "code"
                elif fname_codes:
                    codes, ev = fname_codes, "filename"
                    flags.append("subject code taken from file name (no official code in the Subject_Code field)")
                else:
                    codes, ev = [], "none"
                    label = name_key(r.subject_code_raw) or name_key(r.subject_name_raw)
                    cand = {id(e): e for e in names.get(label, [])} if label else {}
                    if len(cand) == 1:
                        e = list(cand.values())[0]
                        if e.tokens and (fdate is None or fdate == e.date):
                            codes, ev = list(e.tokens), "name"
                            flags.append("subject assigned by exact equality with the official index subject name "
                                         f"'{e.name}' (source carries no subject code)")
                if r.code_inherited:
                    flags.append("subject code inherited from the row above (blank in source)")
                raw = "" if r.reg_raw is None else str(r.reg_raw)
                rec = SrcRec(r, dummy_key(r.dummy_raw) or "", raw, clean_text(raw), codes, ev, fdate, fsess, in_all,
                             corrupt=reg_problem(r.reg_raw), flags=flags)
                if HIDDEN.search(raw):
                    rec.flags.append("hidden characters present in source registration number (preserved as-is)")
                cat.recs.append(rec)
                if rec.dkey:
                    cat.by_dummy[rec.dkey].append(rec)
                for c in codes:
                    cat.by_code[c].append(rec)
    return cat
