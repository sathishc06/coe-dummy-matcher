import io
import os
import zipfile

import openpyxl
import pytest

from coe_matcher.archive import safe_extract
from coe_matcher.gate import REL_COMPLETE, REL_PARTIAL, evaluate_gate
from coe_matcher.index import load_index, norm_session, parse_date, split_tokens
from coe_matcher.normalize import clean_text, code_key, dummy_key, name_key
from coe_matcher.parsing import parse_source_file
from coe_matcher.processing import shift_formula
from coe_matcher.sources import path_date_session, reg_problem
from coe_matcher.verification import eval_marks


def test_dummy_key_formats():
    assert dummy_key(36430017) == dummy_key("36430017") == dummy_key(36430017.0) == dummy_key(" 0036430017 ") == "36430017"
    assert dummy_key("36 430-017") == "36430017"
    assert dummy_key("abc") is None and dummy_key(None) is None and dummy_key(float("nan")) is None


def test_hidden_characters_removed_for_comparison_only():
    s = "AB\u200b C\u00a0D"
    assert clean_text(s) == "AB C D"


def test_name_and_code_keys():
    assert name_key("AERODYNAMICS - I") == name_key("Aerodynamics-I") and name_key("A & B") == name_key("A AND B")
    assert code_key("422-bmt 03") == "422BMT03"


def test_shift_formula():
    assert shift_formula("=SUM(C8:Q8)") == "=SUM(E8:S8)"
    assert shift_formula("=SUM(C8:Q8)", 8, 20) == "=SUM(E20:S20)"
    assert shift_formula('=IF(A8>0,"C8",B8)', 8, 9) == '=IF(A9>0,"C8",D9)'     # string literal untouched, A not shifted
    assert shift_formula("=-R70") == "=-T70"


def test_eval_marks():
    d = {(8, 3): 1, (8, 4): 2, (8, 5): 3}
    g = lambda r, c: d.get((r, c))
    assert eval_marks("=SUM(C8:E8)", g) == 6 and eval_marks("=C8+D8+E8", g) == 6 and eval_marks("=-C8", g) == -1
    assert eval_marks("=VLOOKUP(1,2,3)", g) is None


def test_index_date_and_tokens():
    assert str(parse_date("20.05.2026")) == "2026-05-20" and parse_date("32.01.2026") is None
    assert norm_session("f.n") == "FN" and norm_session("AFTERNOON") == "AN"
    good, bad = split_tokens("622ITT01/CIT04/522CIT01&ITO14")
    assert good == ["622ITT01", "522CIT01"] and bad == ["CIT04", "ITO14"]


def test_path_date_session():
    assert path_date_session("A/9.5.26 FN/x.csv") == ("2026-05-09", "FN")
    assert path_date_session("A/x_2026.csv") == (None, "")


def test_reg_problem_detects_scientific_notation():
    assert "scientific" in reg_problem("2.40361E+15") and reg_problem("") and reg_problem("AC24-123") == ""


def test_parse_headerless_and_header_csv(tmp_path):
    p = tmp_path / "a.csv"
    p.write_text("Dummy_ID,Registration_Number,Subject_Code,Subject_Name\n12345678,AB-1,222ABC01,X\n", encoding="utf-8-sig")
    rows, info = parse_source_file(str(p), "a.csv")
    assert len(rows) == 1 and rows[0].reg_raw == "AB-1" and info.status == "ok"
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append([12345678, "AB-2", "222ABC01", "X"]); wb.save(tmp_path / "b.xlsx")
    rows, info = parse_source_file(str(tmp_path / "b.xlsx"), "b.xlsx")
    assert len(rows) == 1 and info.layout.startswith("xlsx-headerless")


def test_corrupt_and_lock_files_do_not_crash(tmp_path):
    (tmp_path / "bad.xlsx").write_bytes(b"junk")
    rows, info = parse_source_file(str(tmp_path / "bad.xlsx"), "bad.xlsx")
    assert rows == [] and info.status == "unreadable"
    rows, info = parse_source_file(str(tmp_path / "~$x.xlsx"), "~$x.xlsx")
    assert info.status == "skipped_lockfile"


def test_zip_slip_blocked(tmp_path):
    z = tmp_path / "evil.zip"
    with zipfile.ZipFile(z, "w") as zf:
        zf.writestr("../escape.txt", "x"); zf.writestr("ok/a.csv", "a")
    rep = safe_extract(str(z), str(tmp_path / "out"))
    assert not (tmp_path / "escape.txt").exists() and rep.files == ["ok/a.csv"] and any("unsafe" in s for s in rep.skipped)


def test_bad_zip_reported_not_raised(tmp_path):
    (tmp_path / "x.zip").write_bytes(b"not a zip")
    rep = safe_extract(str(tmp_path / "x.zip"), str(tmp_path / "o"))
    assert not rep.ok and "not a valid ZIP" in rep.error


def test_gate_logic():
    assert evaluate_gate({"a": (REL_COMPLETE, "")})["passed"]
    assert not evaluate_gate({"a": (REL_COMPLETE, ""), "b": (REL_PARTIAL, "x")})["passed"]


def test_index_weekday_year_correction(tmp_path):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["Exam Date", "Day", "Session", "Sem", "Sub. Code", "Sub. Name", "Dept"])
    ws.append(["01.06.2026", "Monday", "FN", 2, "111AAA01", "A", "D"])
    ws.append(["02.06.2026", "Tuesday", "FN", 2, "111AAA02", "B", "D"])
    ws.append(["03.06.2026", "Wednesday", "FN", 2, "111AAA03", "C", "D"])
    ws.append(["01.06.2025", "Monday", "FN", 2, "111AAA04", "D", "D"])      # wrong year, weekday says 2026
    wb.save(tmp_path / "i.xlsx")
    idx = load_index(str(tmp_path / "i.xlsx"))
    e = [x for x in idx.entries if "111AAA04" in x.tokens][0]
    assert e.date == "2026-06-01" and "year corrected" in e.date_note
