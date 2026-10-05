"""Crosswalk (valuation sheet -> official index) and the dummy-number matching engine.

RULES (never relaxed):
  * a dummy number is matched ONLY inside source records tied to the sheet's own subject code(s);
  * registration numbers are copied byte-for-byte from the source and never derived or inferred;
  * a dummy with several different registration numbers in the usable evidence is AMBIGUOUS, never guessed.
"""
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from .index import IndexData, IndexEntry
from .models import (AMBIGUOUS, EXCLUDED, NOT_FOUND, VERIFIED, VERIFIED_ALT, RowResult, Unit)
from .normalize import clean_text, dummy_key, name_key
from .sources import SourceCatalog, SrcRec
from .subjects import build_code_lookup, identify

EXCLUDED_DEFAULT = ("225PST02",)


@dataclass
class Crosswalk:
    status: str                       # EXACT | COMPOSITE | MULTI_SESSION | AMBIGUOUS | VARIANT_TOKEN | NOT_IN_INDEX | UNIDENTIFIED
    entries: List[IndexEntry] = field(default_factory=list)
    note: str = ""


@dataclass
class UnitSummary:
    unit_key: str
    wb_rel: str
    sheet: str
    codes: str
    subject_name: str
    id_status: str
    id_note: str
    crosswalk: Crosswalk
    n_rows: int = 0
    counts: Dict[str, int] = field(default_factory=dict)
    sources_used: List[str] = field(default_factory=list)
    source_dummies: int = 0
    notes: List[str] = field(default_factory=list)
    excluded: bool = False
    structure_flags: List[str] = field(default_factory=list)
    val_dates_note: str = ""


def unit_key(u: Unit):
    return f"{u.wb_rel}::{u.sheet}"


def resolve_crosswalk(u: Unit, index: Optional[IndexData]) -> Crosswalk:
    if not u.codes:
        return Crosswalk("UNIDENTIFIED", note="no official-format subject code in the file name or sheet content")
    if not index or not index.entries:
        return Crosswalk("NOT_IN_INDEX", note="no official index loaded")
    ents = [e for e in index.entries if set(e.tokens) & set(u.codes)]
    if not ents:
        # malformed token variants (e.g. 622EEO015 for 622EEO15) are accepted ONLY with exact name corroboration
        var = []
        for e in index.entries:
            for t in e.malformed_tokens:
                if len(t) == 9 and t[6] == "0" and (t[:6] + t[7:]) in u.codes and name_key(u.subject_name) == e.name_key:
                    var.append(e)
        if var:
            return Crosswalk("VARIANT_TOKEN", var, "index code token is malformed; accepted via exact subject-name equality")
        return Crosswalk("NOT_IN_INDEX", note=f"code(s) {u.codes} do not appear in the official index")
    keyset = {(e.date, e.session) for e in ents}
    if len(ents) == 1:
        e = ents[0]
        return Crosswalk("COMPOSITE" if len(e.tokens) > 1 else "EXACT", ents)
    if len(keyset) == 1:
        return Crosswalk("EXACT", ents, "same code listed more than once with the same date/session")
    return Crosswalk("MULTI_SESSION", ents,
                     "code appears in several index entries with different date/session: "
                     + "; ".join(f"{e.date} {e.session}" for e in ents))


def _consistency(rec: SrcRec, ents: List[IndexEntry]):
    if not ents:
        return "UNVERIFIABLE"
    if rec.folder_date is None:
        return "UNDATED"
    for e in ents:
        if e.date == rec.folder_date and (rec.folder_session == "" or rec.folder_session == e.session):
            return "CONSISTENT"
    return "MISMATCH"


def _pri(rec: SrcRec, cons: str):
    return (cons != "CONSISTENT", rec.evidence != "code", rec.in_all_folder, rec.row.source_file)


def match_unit(u: Unit, index: Optional[IndexData], cat: SourceCatalog, excluded_codes=EXCLUDED_DEFAULT, lookup=None):
    key = unit_key(u)
    pre_codes = list(u.codes)
    excl_hit = [c for c in set(u.file_codes) | set(u.content_codes) if c in set(excluded_codes)]
    if getattr(u, "manual_codes", None):
        u.codes, u.id_status = list(u.manual_codes), "MANUAL"
        u.id_note = "subject mapping set manually by an authorised user"
    elif not excl_hit:
        identify(u, index, cat, lookup)
    else:
        u.codes = sorted(set(u.file_codes) | set(u.content_codes))
    cw = resolve_crosswalk(u, index)
    ents = cw.entries
    summ = UnitSummary(key, u.wb_rel, u.sheet, "/".join(u.codes), u.subject_name, u.id_status, u.id_note, cw,
                       n_rows=len(u.student_rows), structure_flags=list(u.structure_flags))
    results: List[RowResult] = []
    # metadata from the index when it is unambiguous
    e0 = ents[0] if len(ents) == 1 or len({(e.date, e.session) for e in ents}) == 1 else None
    meta = dict(subject_name=(e0.name if e0 else u.subject_name), exam_date=(e0.date or "") if e0 else "",
                session=(e0.session if e0 else ""), sem=(str(e0.sem) if e0 and e0.sem is not None else ""),
                dept=(e0.dept if e0 else ""))
    code_txt = "/".join(u.codes)
    # ---- exclusion --------------------------------------------------------------------
    hit = [c for c in u.codes if c in set(excluded_codes)]
    if hit:
        summ.excluded = True
        for r in u.student_rows:
            results.append(RowResult(key, u.wb_rel, u.sheet, r, u.dummy_raw[r], dummy_key(u.dummy_raw[r]) or "", EXCLUDED,
                                     reason=f"subject {hit[0]} is excluded by instruction (manual verification by COE)",
                                     subject_codes=code_txt, **meta))
        summ.counts = dict(Counter(x.status for x in results))
        return results, summ
    # ---- candidate source records -----------------------------------------------------
    cand_by_d: Dict[str, List[SrcRec]] = defaultdict(list)
    seen = set()
    for c in u.codes:
        for rec in cat.by_code.get(c, []):
            if id(rec) in seen:
                continue
            seen.add(id(rec))
            cand_by_d[rec.dkey].append(rec)
    used_files = Counter()
    summ.source_dummies = len(cand_by_d)
    dup_count = Counter(dummy_key(u.dummy_raw[r]) for r in u.student_rows)
    dup_sigs = defaultdict(set)
    for r in u.student_rows:
        dup_sigs[dummy_key(u.dummy_raw[r])].add(u.row_sig.get(r))
    for r in u.student_rows:
        raw = u.dummy_raw[r]
        dk = dummy_key(raw)
        rr = RowResult(key, u.wb_rel, u.sheet, r, raw, dk or "", NOT_FOUND, subject_codes=code_txt, **meta)
        if not dk:
            rr.reason = "dummy number is not numeric / has no usable digits"
            rr.flags.append("INVALID_DUMMY_VALUE")
            results.append(rr)
            continue
        if dup_count[dk] > 1:
            same = len(dup_sigs[dk]) == 1
            rr.flags.append("DUPLICATE_DUMMY_IN_WORKBOOK_IDENTICAL_MARKS" if same else "DUPLICATE_DUMMY_IN_WORKBOOK_DIFFERENT_MARKS")
        cands = cand_by_d.get(dk, [])
        usable = [c for c in cands if not c.corrupt]
        if not usable:
            if cands:
                rr.reason = ("; ".join(sorted({f"{c.row.source_file} row {c.row.row_no}: {c.corrupt}" for c in cands}))
                             + " - no other usable copy exists")
                rr.flags.append("SOURCE_REG_UNUSABLE")
            else:
                others = cat.by_dummy.get(dk, [])
                if others:
                    ds = sorted({f"{'/'.join(o.codes) or 'unknown'} ({o.row.source_file})" for o in others})
                    rr.reason = (f"dummy number {dk} has no record in any source tied to subject {code_txt}; it exists only in "
                                 f"source(s) of other subject(s): {'; '.join(ds[:4])}{' ...' if len(ds) > 4 else ''}. "
                                 "Cross-subject matching is prohibited, so no registration number was assigned")
                    rr.flags.append("WRONG_SUBJECT_ONLY")
                else:
                    rr.reason = f"dummy number {dk} is not present in any Daywise source file"
                    rr.flags.append("NO_SOURCE_RECORD")
            results.append(rr)
            continue
        consd = [(c, _consistency(c, ents)) for c in usable]
        cons_ok = [c for c, k in consd if k == "CONSISTENT"]
        pool = cons_ok if cons_ok else usable
        ignored = [c for c in usable if c not in pool and c.reg_key not in {p.reg_key for p in pool}]
        regs = {}
        for c in pool:
            regs.setdefault(c.reg_key, []).append(c)
        if len(regs) > 1:
            rr.status = AMBIGUOUS
            rr.reason = ("conflicting registration numbers for the same dummy within the subject's evidence: "
                         + " | ".join(f"'{v[0].reg}' ({v[0].row.source_file} row {v[0].row.row_no})" for v in regs.values()))
            rr.flags.append("CONFLICTING_MAPPING")
            results.append(rr)
            continue
        recs = sorted(pool, key=lambda c: _pri(c, "CONSISTENT" if c in cons_ok else ""))
        best = recs[0]
        rr.reg_no = best.reg
        rr.reg_edit = best.reg.split("-", 1)[1] if "-" in best.reg else ""
        rr.src_file, rr.src_sheet, rr.src_row = best.row.source_file, best.row.sheet, best.row.row_no
        rr.src_subject_code, rr.src_evidence = "/".join(best.codes), best.evidence
        rr.other_sources = " | ".join(sorted({f"{c.row.source_file}[{c.row.sheet}]#{c.row.row_no}" for c in recs[1:]}))
        used_files[best.row.source_file] += 1
        notes, why_alt = [], []
        if "-" not in best.reg:
            rr.flags.append("REG_HAS_NO_HYPHEN_REG_EDIT_BLANK")
        if best.flags:
            notes += best.flags
        if ignored:
            notes.append("ignored conflicting values from non-session copies: "
                         + "; ".join(sorted({f"'{c.reg}' in {c.row.source_file}" for c in ignored})))
        if cw.status == "MULTI_SESSION":
            why_alt.append("subject appears in several index sessions; session assigned only by dummy containment")
        if not ents:
            why_alt.append("subject/session cannot be confirmed in the official index")
        elif not cons_ok:
            why_alt.append("no source copy whose folder date matches the index exam date "
                           f"({', '.join(sorted({c.folder_date or 'undated' for c in usable}))})")
        elif not any(c.evidence in ("code", "filename") for c in cons_ok):
            why_alt.append("subject tied to source only by official-name equality")
        if u.id_status in ("NORMALIZED", "CONFLICT", "CONFLICT_RESOLVED", "MANUAL"):
            why_alt.append("subject identity not exact: " + u.id_note)
            rr.flags.append("SUBJECT_ID_" + u.id_status)
        if cw.status == "VARIANT_TOKEN":
            why_alt.append("index code token malformed; accepted by name equality")
        if why_alt:
            rr.status = VERIFIED_ALT
            notes = why_alt + notes
        else:
            rr.status = VERIFIED
        rr.reason = "; ".join(notes)
        if "DUPLICATE_DUMMY_IN_WORKBOOK_DIFFERENT_MARKS" in rr.flags:
            rr.status = AMBIGUOUS
            rr.reg_no, rr.reg_edit = "", ""
            rr.reason = ("the same dummy number appears on several rows of this sheet with DIFFERENT marks; "
                         "it cannot be determined which row belongs to the student (registration would have been "
                         f"'{best.reg}')")
        results.append(rr)
    summ.sources_used = [f"{f} ({n} rows)" for f, n in used_files.most_common()]
    summ.counts = dict(Counter(x.status for x in results))
    # context for unresolved rows: index arrear counts vs. regular-only sources
    unresolved = [x for x in results if x.status == NOT_FOUND]
    if unresolved and ents:
        arr = sum((e.arr or 0) for e in ents) if len(ents) else 0
        reg = sum((e.reg or 0) for e in ents) if len(ents) else 0
        if arr:
            summ.notes.append(f"official index lists REG={reg}, ARR={arr} for this subject; the Daywise archive supplied "
                              f"contains {summ.source_dummies} mapped dummy number(s) for it")
    return results, summ
