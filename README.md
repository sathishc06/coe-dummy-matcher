# COE Dummy Number Matcher

Reusable Streamlit application for the Controller of Examinations office. It maps valuation-sheet **dummy numbers** to
**registration numbers** using the Daywise source files, writes corrected valuation workbooks, and produces a full audit.

## Install and run
```bash
python -m venv .venv && source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
export COE_ADMIN_PASSWORD="choose-a-password"             # enables manual corrections; leave unset to disable them
streamlit run app.py
```
Docker: `docker build -t coe-matcher . && docker run -p 8501:8501 -e COE_ADMIN_PASSWORD=secret coe-matcher`
Upload limit is 2 GB (`.streamlit/config.toml`). A full May-2026-sized batch (700 workbooks, 590 source files) takes ~9 minutes
on 8 workers; use the worker slider on the Upload page.

## Using it
1. **Upload** the Valuation ZIP, Daywise ZIP and (optionally) the official Daywise date/session index `.xlsx`
   (auto-detected inside the Daywise ZIP if omitted). List subject codes to exclude (default `225PST02`).
2. Review **File inventory**, **Subject mapping** (duplicate code/name conflicts are listed; authorised users can set a mapping).
3. Watch **Reconciliation progress**; browse **Verified matches** and **Unmatched records** with filters by department,
   subject, semester, exam date and session.
4. **Conflict resolution**: authorised users pick a source record that carries the dummy number and must give reason,
   evidence and approval note; the change is logged (previous value, new value, time, user) and the run is reprocessed.
   Original source files are never overwritten.
5. **Workbook verification** shows the independent checks for each output workbook.
6. **Final download**. The official ZIP exists only if the strict release gate passes; otherwise audit-only downloads are
   labelled **NOT APPROVED FOR OFFICIAL USE**.

## Rules the engine enforces
* Dummy numbers are matched only inside source records tied to the sheet's own subject code (never globally).
* Registration numbers are copied byte-for-byte from the source; never derived from patterns or neighbouring rows.
* Several different registration numbers for one dummy, or one dummy repeated on several sheet rows with different marks -> AMBIGUOUS.
* Source copies with a corrupted registration number (e.g. scientific notation `2.40E+15`) are never used.
* Statuses: VERIFIED, VERIFIED_ALTERNATE_SOURCE (identity or session needed corroboration - always shown with the reason),
  AMBIGUOUS, NOT_FOUND, EXCLUDED, PROCESSING_ERROR.
* Output sheets get `Reg.No` and `Reg Edit` right after `Dummy No.`; student rows are sorted by full Reg.No; unmatched rows
  follow, highlighted. Row formulas are rewritten for the shifted columns and the student's new row.
  `Reg Edit` is blank when the Reg.No has no hyphen.
* Workbooks with unsupported structure (cross-row/cross-sheet formulas, multi-row merges over student rows, macros, charts,
  tables, conditional formatting, data validation, protection) are blocked, not rewritten.
* Output workbooks are saved with *recalculate on open*; cached marks values appear after Excel/LibreOffice opens them.

## Release gate
Official ZIP is refused if any included workbook has an unresolved or duplicate row, fails independent verification,
is blocked, or could not be written.

## Modules (`coe_matcher/`)
`archive` (safe ZIP) · `parsing`/`sources` (Daywise CSV/Excel) · `index` (official schedule) · `valuation` (sheet scan) ·
`subjects` (code identification) · `matching` (engine) · `processing` (rewrite) · `verification` (independent checks) ·
`gate` · `audit` · `corrections` · `pipeline`.

## Tests
`python -m pytest tests -v` - unit tests, integration tests on a synthetic dataset with deliberate mismatches
(`tests/make_test_dataset.py`), and Streamlit page-render tests.

## Known limitations
See TEST_REPORT.md ("Not tested / incomplete").

## Usable subject/department release

The release page automatically creates `USABLE_SUBJECT_DEPARTMENT_RELEASE.zip`. It includes only rows from workbooks classified `COMPLETE_VERIFIED` and independently verified as PASS. Unresolved, ambiguous, blocked, excluded, and unverified subjects/workbooks are omitted automatically.

Output layout inside the ZIP:

```text
OFFICIAL_USABLE_SUBJECTS/
  <SUBJECT_CODE>/
    <SUBJECT_CODE>_<DEPT_CODE>.xlsx
```

The department code is derived from the complete `Reg.No` by taking the first three consecutive digits occurring after the `6176` prefix. Example: a registration containing `6176...101...` is grouped under department code `101`. Each department gets a separate Excel file. No manual correction or second manual re-check is performed during release packaging.
