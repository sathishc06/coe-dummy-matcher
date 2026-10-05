# Test Report - COE Dummy Number Matcher

All results below were produced by commands actually executed in this session (logs: `pytest_full.log`, `real_check.log`).

## 1. Automated test suite - `python -m pytest tests -v`
**34 passed, 0 failed.**
* `tests/test_unit.py` (14): dummy-key formats (int/float/text/leading zero/hyphen), hidden characters, name/code keys, formula column-shift and row-renumber,
  marks-formula evaluator, index date/session/token parsing, weekday-based year correction, folder date/session detection,
  scientific-notation detection, header/headerless source parsing, corrupt + lock files, zip-slip blocking, bad ZIP, gate logic.
* `tests/test_integration.py` (16) on a synthetic dataset with deliberate faults (`tests/make_test_dataset.py`): missing source record,
  dummy that exists only in another subject, duplicate dummy with different marks, conflicting registrations across two source copies,
  corrupted (scientific-notation) registration with a good alternate copy, float/leading-zero dummy formats, file-name vs content
  subject-code conflict, excluded subject, corrupt/lock/summary-only workbooks, corrupt source file. Also: output layout, sort order,
  marks stay with their dummy, formulas follow rows, unmatched rows last, independent verification passes, release gate blocks the
  official ZIP, official ZIP is produced for a clean batch, manual correction refused without evidence / logged with evidence,
  audit sheets present, blank cells stay blank in CSV/audit.
* `tests/test_app.py` (4): Streamlit `AppTest` - upload page renders; other pages ask for an upload first; all 9 pages render with a loaded run and the final page shows NOT APPROVED FOR OFFICIAL USE; manual correction through the Conflict-resolution page (wrong password locks it, correct password + reason/evidence/approval note re-runs the reconciliation and logs the change).

## 2. Real May 2026 archives - full pipeline run
Executed on `VALUATION_MAY_2026.zip`, `Daywise.zip`, `DAYWISE_MAY_2026.xlsx` (8 worker processes, 565 s).
* 702 workbook files (10 Office `~$` lock files skipped), 31,938 student rows, 591 Daywise files (496 CSV / 95 Excel; one lock file skipped), 45,950 source records.
* **Independent verification: 376 of 376 written workbooks passed** (checks per workbook: reopen, sheet names, header positions, row/column extent,
  student count, per-dummy marks multiset, dummy multiset, own-row formulas reproducing original totals, Reg.No found in raw source for the
  same subject, Reg Edit rule, ascending sort, non-student cells unchanged, images, merges, freeze panes).
* **`tests/real_data_check.py` (reads outputs from disk only): 9 of 9 passed** - CSV/audit counts agree; no Reg.No on unresolved or excluded rows;
  225PST02 not written; all 26,109 matched Reg.No values re-found in raw source for the same subject; 376 workbooks keep sheets and non-empty rows;
  no official ZIP present while the gate fails; output ZIPs open cleanly.
* **LibreOffice recalculation (sample only):** 14 corrected workbooks (12 complete, 2 partial; 3,622 student rows) were recalculated headless;
  every recalculated total equals that student's original cached total, 0 error cells. The other 362 workbooks were NOT recalculated this way.

## 3. Defects found by testing and fixed during the session
1. A single-row merged range inside a student row was left behind when rows were sorted (found by independent verification on 622ME017) - merges now move with their row; multi-row merges over student rows block the workbook.
2. Independent verifier did not understand source files that carry only a subject *name* (422AEI02) - name evidence added to the verifier.
3. Blank text cells were written as `'` in the CSV/audit (found by `real_data_check.py`) - fixed, regression test added.
4. Non-breaking spaces were deleted instead of treated as spaces in comparisons - fixed.
5. Pipeline refactor left an undefined variable (found by the first synthetic run) - fixed.

## 4. Implemented / tested / incomplete
| Feature | Implemented | Tested |
|---|---|---|
| ZIP upload, safe extraction, corrupt-file reporting | yes | yes (unit + synthetic + real) |
| Automatic index detection | yes | yes (real: found inside Daywise ZIP) |
| Subject identification + index crosswalk | yes | yes |
| Matching engine (scoped, ambiguity, corruption) | yes | yes |
| Workbook rewrite (Reg.No/Reg Edit, sort, formulas) | yes | yes (synthetic, real, LibreOffice sample) |
| Independent verification | yes | yes |
| Release gate | yes | yes |
| Audit XLSX / CSV / report / ZIPs | yes | yes |
| Streamlit pages (9) incl. filters, indicators | yes | render-tested (all 9 pages, empty and with a loaded run); filter widgets not exercised |
| Manual correction (engine) | yes | yes (logic) |
| Manual correction through the UI (AppTest) | yes | yes (simulated Streamlit session, not a real browser) |
| Manual subject mapping through the UI | yes | **not tested** |
| Password gate | yes (single shared env password) | wrong/right password tested on the correction page only |
| Large (multi-GB) uploads, concurrent users | configured (2 GB limit) | **not tested** |
| `.xlsm`, Windows, Excel desktop opening files | blocked/untested | **not tested** |
| Docker image | Dockerfile written | **not built or run** |
