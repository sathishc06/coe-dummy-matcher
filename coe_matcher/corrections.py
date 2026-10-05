"""Manual correction workflow. Original source data is never overwritten; every override is logged."""
import csv
import datetime as dt
import io
from dataclasses import dataclass, asdict
from typing import List

from .models import AMBIGUOUS, NOT_FOUND, VERIFIED_ALT


@dataclass
class Override:
    unit_key: str
    sheet_row: int
    chosen_src_file: str
    chosen_src_row: int
    reason: str
    source_evidence: str
    approval_note: str
    user: str = ""
    timestamp: str = ""
    previous_status: str = ""
    previous_reg: str = ""
    corrected_reg: str = ""


LOG_FIELDS = ["timestamp", "user", "unit_key", "sheet_row", "previous_status", "previous_value", "corrected_value",
              "reason", "source_evidence", "approval_note", "chosen_src_file", "chosen_src_row"]


def candidates_for(run, rr):
    """Every source record carrying this dummy number (any subject) - shown to the reviewer as evidence options."""
    out = []
    for c in run.catalog.by_dummy.get(rr.dummy_key, []):
        out.append(dict(source_file=c.row.source_file, sheet=c.row.sheet, row=c.row.row_no, reg=c.reg,
                        subject_codes="/".join(c.codes), unusable=c.corrupt, folder_date=c.folder_date or ""))
    return out


def apply_overrides(run, overrides: List[Override]):
    """Apply approved overrides to run.results (only unresolved rows). Returns the audit log rows."""
    idx = {(r.unit_key, r.sheet_row): r for r in run.results}
    log = []
    for o in overrides:
        rr = idx.get((o.unit_key, o.sheet_row))
        if rr is None or rr.status not in (AMBIGUOUS, NOT_FOUND):
            continue
        if not (o.reason.strip() and o.source_evidence.strip() and o.approval_note.strip()):
            continue                                   # incomplete evidence -> refused
        cand = next((c for c in run.catalog.by_dummy.get(rr.dummy_key, [])
                     if c.row.source_file == o.chosen_src_file and c.row.row_no == o.chosen_src_row), None)
        if cand is None or cand.corrupt:
            continue
        o.timestamp = o.timestamp or dt.datetime.now().isoformat(timespec="seconds")
        o.previous_status, o.previous_reg, o.corrected_reg = rr.status, rr.reg_no, cand.reg
        rr.status = VERIFIED_ALT
        rr.reg_no = cand.reg
        rr.reg_edit = cand.reg.split("-", 1)[1] if "-" in cand.reg else ""
        rr.src_file, rr.src_sheet, rr.src_row = cand.row.source_file, cand.row.sheet, cand.row.row_no
        rr.manual_override = True
        rr.flags.append("MANUAL_OVERRIDE")
        rr.reason = f"MANUAL: {o.reason} | evidence: {o.source_evidence} | approval: {o.approval_note} ({o.user} {o.timestamp})"
        log.append(dict(timestamp=o.timestamp, user=o.user, unit_key=o.unit_key, sheet_row=o.sheet_row,
                        previous_status=o.previous_status, previous_value=o.previous_reg or "(none)",
                        corrected_value=o.corrected_reg, reason=o.reason, source_evidence=o.source_evidence,
                        approval_note=o.approval_note, chosen_src_file=o.chosen_src_file, chosen_src_row=o.chosen_src_row))
    run.manual_log = getattr(run, "manual_log", []) + log
    return log


def log_to_csv(log):
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=LOG_FIELDS)
    w.writeheader()
    w.writerows(log)
    return buf.getvalue()
