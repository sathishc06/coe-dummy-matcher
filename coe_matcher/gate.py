"""Strict release gate. Official ZIP generation is refused unless every condition passes."""
from collections import defaultdict
from typing import Dict, List

from .models import AMBIGUOUS, EXCLUDED, MATCHED_STATUSES, NOT_FOUND, PROC_ERROR

REL_COMPLETE = "COMPLETE_VERIFIED"
REL_PARTIAL = "PARTIAL_EXCEPTIONS_REMAIN"
REL_BLOCKED = "BLOCKED"
REL_NONE = "NO_OUTPUT"
REL_EXCL = "EXCLUDED"
REL_NA = "NO_STUDENT_ROWS"


def workbook_release_status(unit_rows: Dict[str, list], unit_flags: Dict[str, list], written: bool, verified: bool, blocked_msg: str):
    """Per-workbook classification (all sheets that hold students)."""
    rows = [r for rs in unit_rows.values() for r in rs]
    if not rows:
        return REL_NA, "workbook has no student rows"
    sts = {r.status for r in rows}
    if sts == {EXCLUDED}:
        return REL_EXCL, "excluded subject"
    if blocked_msg:
        return REL_BLOCKED, blocked_msg
    matched = [r for r in rows if r.status in MATCHED_STATUSES]
    if not matched:
        return REL_NONE, "no row could be verified"
    reasons = []
    unresolved = [r for r in rows if r.status in (AMBIGUOUS, NOT_FOUND, PROC_ERROR)]
    if unresolved:
        reasons.append(f"{len(unresolved)} unresolved row(s)")
    dup = [r for r in rows if any(f.startswith("DUPLICATE_DUMMY") for f in r.flags)]
    if dup:
        reasons.append(f"{len(dup)} row(s) share a dummy number with another row of the same sheet (duplicate review required)")
    if written and not verified:
        return REL_BLOCKED, "independent verification FAILED"
    if not written:
        return REL_BLOCKED, "workbook was not written"
    if reasons:
        return REL_PARTIAL, "; ".join(reasons)
    return REL_COMPLETE, ""


def evaluate_gate(wb_status: Dict[str, tuple], unresolved_manual_conflicts=0) -> dict:
    """wb_status: rel -> (release_status, reason). Gate passes only if no included workbook is PARTIAL/BLOCKED/NONE."""
    problems = []
    for rel, (st, why) in sorted(wb_status.items()):
        if st in (REL_PARTIAL, REL_BLOCKED, REL_NONE):
            problems.append((rel, st, why))
    return {"passed": not problems and unresolved_manual_conflicts == 0, "problems": problems}
