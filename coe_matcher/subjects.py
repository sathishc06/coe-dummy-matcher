"""Subject identification: file-name / workbook-content codes -> official index subject.

Rules:
  * exact official-format codes are preferred;
  * look-alike character fixes (O<->0, I<->1) are applied ONLY when the corrected code maps to exactly ONE official
    index subject AND is corroborated by exact official-name equality or by the sheet's dummy numbers being present
    in that subject's source files. Such assignments are never classed as plain VERIFIED;
  * conflicting file-name / content codes are resolved only by name equality or dummy containment, otherwise the
    union of codes is searched and any collision becomes AMBIGUOUS.
"""
import re
from collections import defaultdict
from typing import Dict, List, Optional

from .index import IndexData, IndexEntry
from .models import Unit
from .normalize import clean_text, dummy_key, name_key
from .sources import SourceCatalog

LOOSE = re.compile(r"(?<![A-Za-z0-9])\d{3}[A-Za-z0-9]{5}(?![0-9])")
STRICT = re.compile(r"\d{3}[A-Za-z]{3}\d{2}")


def canon(code: str) -> str:
    return code.upper().replace("O", "0").replace("I", "1")


def _loose_codes(text):
    return [m.upper() for m in LOOSE.findall(clean_text(text))]


def build_code_lookup(index: Optional[IndexData]):
    exact: Dict[str, List[IndexEntry]] = defaultdict(list)
    canon_map: Dict[str, List[IndexEntry]] = defaultdict(list)
    if index:
        for e in index.entries:
            for t in e.tokens:
                exact[t].append(e)
            for t in e.tokens + e.malformed_tokens:
                canon_map[canon(t)].append(e)
    return exact, canon_map


def _subject_groups(entries):
    return {frozenset(e.tokens) for e in entries}


def identify(u: Unit, index: Optional[IndexData], cat: SourceCatalog, lookup=None):
    """Sets u.codes / u.id_status / u.id_note. Returns nothing."""
    exact, canon_map = lookup or build_code_lookup(index)
    uname = name_key(u.subject_name)
    dkeys = {dummy_key(v) for v in u.dummy_raw.values()} - {None}
    cands = list(dict.fromkeys(u.file_codes + u.content_codes))
    notes = []
    if not cands:
        # try look-alike normalisation on loosely shaped codes
        import os
        loose = list(dict.fromkeys(_loose_codes(os.path.splitext(os.path.basename(u.wb_rel))[0]) +
                                   _loose_codes(u.c6) + _loose_codes(u.sheet)))
        hits = {}
        for lc in loose:
            for e in canon_map.get(canon(lc), []):
                hits[(id(e))] = (lc, e)
        subj = _subject_groups([e for _, e in hits.values()])
        if len(subj) == 1:
            lc, e = list(hits.values())[0]
            tok = next((t for t in e.tokens if canon(t) == canon(lc)), None) or \
                  next((t for t in e.tokens + e.malformed_tokens if canon(t) == canon(lc)), None)
            tok_ok = tok in e.tokens
            name_eq = bool(uname) and uname == e.name_key
            contained = 0
            if tok_ok:
                have = {r.dkey for r in cat.by_code.get(tok, [])}
                contained = len(dkeys & have)
            ratio = contained / max(1, len(dkeys))
            if tok_ok and (name_eq or ratio >= 0.8):
                u.codes = [tok]
                u.id_status = "NORMALIZED"
                u.id_note = (f"code '{lc}' read as official code '{tok}' (look-alike O/0, I/1 correction); unique official "
                             f"match; corroborated by "
                             + ("exact official subject-name equality" if name_eq else
                                f"{contained}/{len(dkeys)} dummy numbers present in that subject's source files"))
                return
            u.id_status = "UNIDENTIFIED"
            u.id_note = (f"code '{lc}' resembles official code(s) {sorted({t for t in e.tokens + e.malformed_tokens})} "
                         "but could not be corroborated; not assigned")
            return
        u.id_status = "UNIDENTIFIED"
        u.id_note = ("no official-format subject code found in file name or sheet content"
                     + (f"; look-alike candidates map to {len(subj)} different subjects" if len(subj) > 1 else ""))
        return
    known = {c: exact.get(c, []) for c in cands}
    groups = _subject_groups([e for es in known.values() for e in es])
    if len(groups) <= 1:
        u.codes = sorted(cands)
        if u.file_codes and u.content_codes and not (set(u.file_codes) & set(u.content_codes)):
            if groups:
                u.id_status = "CONSISTENT"
                u.id_note = f"file-name codes {u.file_codes} and content codes {u.content_codes} belong to the same official index entry"
            else:
                u.id_status = "CONFLICT"
                u.id_note = f"file-name codes {u.file_codes} differ from content codes {u.content_codes}; neither is in the index"
        elif u.file_codes and not u.content_codes:
            u.id_status = "FILENAME_ONLY"
        elif u.content_codes and not u.file_codes:
            u.id_status = "CONTENT_ONLY"
        else:
            u.id_status = "CONSISTENT"
        return
    # --- conflict between different official subjects ----------------------------------
    base_note = (f"file-name code(s) {u.file_codes} vs workbook content code(s) {u.content_codes} point to different "
                 "official subjects")
    by_name = {g for g in groups if any(e.name_key == uname for e in
                                          [e for es in known.values() for e in es] if frozenset(e.tokens) == g)} if uname else set()
    if len(by_name) == 1:
        g = list(by_name)[0]
        u.codes = sorted(set(cands) & set(g))
        u.id_status = "CONFLICT_RESOLVED"
        u.id_note = base_note + f"; resolved to {u.codes} because the sheet's subject name equals the official name"
        return
    score = {}
    for g in groups:
        have = {r.dkey for c in set(cands) & set(g) for r in cat.by_code.get(c, [])}
        score[g] = len(dkeys & have)
    ranked = sorted(score.items(), key=lambda kv: -kv[1])
    top, second = ranked[0][1], ranked[1][1]
    if top >= 0.5 * len(dkeys) and top >= 2 * max(second, 1) - (1 if second == 0 else 0):
        g = ranked[0][0]
        u.codes = sorted(set(cands) & set(g))
        u.id_status = "CONFLICT_RESOLVED"
        u.id_note = (base_note + f"; resolved to {u.codes} because {top}/{len(dkeys)} dummy numbers exist in that subject's "
                     f"source files versus {second} for the other")
        return
    u.codes = sorted(cands)
    u.id_status = "CONFLICT"
    u.id_note = base_note + "; could not be resolved by name or dummy containment, union of codes searched"
