# Remaining unresolved issues (May 2026)

Release gate: **FAILED - NOT APPROVED FOR OFFICIAL USE.** 100% completion is not claimed.

## Unresolved rows: 4,661 of 30,770 included rows (excluded 225PST02 rows: 1,168)
| Count | Exact reason | Evidence needed |
|---:|---|---|
| 3,826 | Dummy number is in **no** Daywise file (largest: 122MAT02-ARR 216, 222MAT02-ARR 178, 125CYT04 140, 422ECI03 133, 125MLT02-ARR 123, 622CST03 120, 125PPI05 111, 822ECE29 101, 122CYT04 93, 422ECT01 90, 322CST04 88) | Daywise dummy-to-register lists for the arrear / supplementary sessions |
| 820 | Dummy exists only in a **different subject's** source (e.g. 154 rows in 622CSI04 appear only in 622CIT01) - cross-subject matching is prohibited | Confirmation of the correct subject for those dummy numbers, or the correct Daywise file |
| 15 | Same dummy on several rows of one sheet with different marks (422CST06, 622MGE03, 622CHE22) | Corrected valuation sheet |

## Needs human review even though matched
* **399 alternate-source rows** come from sheets whose file-name code and workbook-content code point to *different* official subjects and could not be resolved by name or dummy containment; both subjects' sources were searched and each Reg.No was confirmed in a source (conflicts would have been AMBIGUOUS). Confirm the subject of these workbooks.
* **627 rows** where the same conflict was resolved by exact subject-name equality or dummy containment; **1,500 rows** where a look-alike code (O/0, I/1) was corrected; **43** subject tied only by exact official-name equality; **36** other session/date corroboration cases.
* **238 rows** share a dummy number with another row of the same sheet and identical marks (e.g. repeated blocks in 422CIT04-OS). They were kept (no rows deleted) and make the workbook PARTIAL.
* **11,208 of 26,109** matched registration numbers contain no hyphen, so `Reg Edit` is blank for them. Say if you want the full number copied instead.

## Workbooks
* 339 COMPLETE, 37 PARTIAL (exceptions remain), 4 BLOCKED, 279 with no verifiable row (no output), 29 with no student rows, 4 excluded.
* BLOCKED: 422ECT01 COMMUNICATION THEORY (two copies; marks cell M70 `=-G118` points at another row), 622MEI03-FEA (formulas in row 28 point at row 29), Maths/TODAY TOTAL PAPER.xlsx (no Dummy No. table). These need manual handling.
* Output files recalculate on open; only 14 were recalculated in LibreOffice (all matched).
