"""Shared data model and status vocabulary."""
from dataclasses import dataclass, field
from typing import Dict, List, Optional

# ---- match statuses (spec: Phase 3) --------------------------------------------------
VERIFIED = "VERIFIED"
VERIFIED_ALT = "VERIFIED_ALTERNATE_SOURCE"
AMBIGUOUS = "AMBIGUOUS"
NOT_FOUND = "NOT_FOUND"
EXCLUDED = "EXCLUDED"
PROC_ERROR = "PROCESSING_ERROR"
MATCHED_STATUSES = (VERIFIED, VERIFIED_ALT)
ALL_STATUSES = (VERIFIED, VERIFIED_ALT, AMBIGUOUS, NOT_FOUND, EXCLUDED, PROC_ERROR)


@dataclass
class Unit:
    """One valuation worksheet that contains a 'Dummy No.' student table."""
    wb_rel: str                     # workbook path relative to valuation archive root
    wb_abs: str
    sheet: str
    header_row: int = 0
    dummy_col: int = 1
    max_col: int = 0
    max_row: int = 0
    student_rows: List[int] = field(default_factory=list)       # sheet row numbers that hold a dummy number
    dummy_raw: Dict[int, object] = field(default_factory=dict)   # row -> raw dummy cell value
    c5: str = ""
    c6: str = ""
    q5: str = ""
    q6: str = ""
    file_codes: List[str] = field(default_factory=list)
    content_codes: List[str] = field(default_factory=list)
    codes: List[str] = field(default_factory=list)              # codes used for matching
    subject_name: str = ""
    id_status: str = ""             # CONSISTENT | CONFLICT | FILENAME_ONLY | CONTENT_ONLY | UNIDENTIFIED
    id_note: str = ""
    structure_flags: List[str] = field(default_factory=list)    # reasons the workbook cannot be safely rewritten
    kind: str = "STUDENT_SHEET"     # STUDENT_SHEET | BLANK_TEMPLATE | NO_STUDENT_ROWS
    excluded: bool = False
    exclude_reason: str = ""
    n_sheets: int = 1
    row_sig: Dict[int, tuple] = field(default_factory=dict)       # row -> non-formula data values (for duplicate checks)


@dataclass
class RowResult:
    unit_key: str                   # "<workbook rel>::<sheet>"
    wb_rel: str
    sheet: str
    sheet_row: int
    dummy_raw: object
    dummy_key: str
    status: str
    reg_no: str = ""                # exact original registration number (only when matched)
    reg_edit: str = ""
    reason: str = ""                # exact reason when not matched, or notes when matched
    flags: List[str] = field(default_factory=list)
    subject_codes: str = ""
    subject_name: str = ""
    exam_date: str = ""
    session: str = ""
    sem: str = ""
    dept: str = ""
    src_file: str = ""
    src_sheet: str = ""
    src_row: object = ""
    src_subject_code: str = ""
    src_evidence: str = ""          # how the source was tied to the subject: code | filename | name
    other_sources: str = ""         # additional agreeing copies
    manual_override: bool = False
    verification: str = ""          # filled by verification stage
