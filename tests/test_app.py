import os

from streamlit.testing.v1 import AppTest

APP = os.path.join(os.path.dirname(os.path.dirname(__file__)), "app.py")


def test_upload_page_renders_without_exception():
    at = AppTest.from_file(APP, default_timeout=60).run()
    assert not at.exception and at.title[0].value == "COE Dummy Number Matcher"


def test_other_pages_ask_for_upload_first():
    at = AppTest.from_file(APP, default_timeout=60).run()
    for page in ["2. File inventory", "5. Verified matches", "9. Final download"]:
        at.sidebar.radio[0].set_value(page).run()
        assert not at.exception and any("Upload files" in i.value for i in at.info)


def _loaded_run(tmp_path):
    import sys
    sys.path.insert(0, os.path.dirname(__file__))
    import make_test_dataset as mk
    from conftest import subset_zip
    from coe_matcher.pipeline import run_reconciliation
    vz, dz, idx, _ = mk.build(str(tmp_path / "ds"))
    vz2 = subset_zip(vz, str(tmp_path / "v.zip"), ["111AAA02-GEOMETRY"])
    return run_reconciliation(vz2, dz, idx, str(tmp_path / "w"), processes=1)


def test_all_pages_render_with_loaded_run(tmp_path):
    run = _loaded_run(tmp_path)
    at = AppTest.from_file(APP, default_timeout=120)
    at.session_state["run"], at.session_state["procs"] = run, 1
    at.session_state["overrides"], at.session_state["subject_overrides"] = [], {}
    at.run()
    for page in ["2. File inventory", "3. Subject mapping", "4. Reconciliation progress", "5. Verified matches",
                 "6. Unmatched records", "7. Conflict resolution", "8. Workbook verification", "9. Final download"]:
        at.sidebar.radio[0].set_value(page).run()
        assert not at.exception, (page, [e.value for e in at.exception])
    at.sidebar.radio[0].set_value("9. Final download").run()
    assert any("NOT APPROVED FOR OFFICIAL USE" in e.value for e in at.error)


def test_manual_correction_through_ui(tmp_path, monkeypatch):
    monkeypatch.setenv("COE_ADMIN_PASSWORD", "secret")
    run = _loaded_run(tmp_path)
    at = AppTest.from_file(APP, default_timeout=120)
    at.session_state["run"], at.session_state["procs"] = run, 1
    at.session_state["overrides"], at.session_state["subject_overrides"] = [], {}
    at.run()
    at.sidebar.radio[0].set_value("7. Conflict resolution").run()
    # wrong password -> locked
    at.text_input(key="pw7").set_value("wrong").run()
    assert any("Authorised users only" in i.value for i in at.info)
    at.text_input(key="pw7").set_value("secret").run()
    assert not at.exception
    ta = at.text_area
    ta[0].set_value("student sat GEOMETRY under ALGEBRA bundle")
    ta[1].set_value("attendance sheet p3")
    ta[2].set_value("approved by COE")
    at.text_input[1].set_value("tester")
    # pick the candidate row dummy 1000001 (the NOT_FOUND wrong-subject row)
    sel = [s for s in at.selectbox if s.label == "Row"][0]
    opt = [o for o in sel.options if "1000001" in o][0]
    sel.set_value(opt)
    at.run()
    [b for b in at.button if b.label == "Record correction and re-run"][0].click().run()
    assert not at.exception
    assert any("Correction applied" in s.value for s in at.success)
    row = [r for r in at.session_state["run"].results if r.dummy_key == "1000001"][0]
    assert row.status == "VERIFIED_ALTERNATE_SOURCE" and row.manual_override
