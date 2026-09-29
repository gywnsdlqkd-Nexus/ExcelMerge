# -*- coding: utf-8 -*-
"""키 열을 파일별로 기억한다 — 전역 키 하나로 모든 파일을 비교하던 문제.

키는 오랫동안 prefs 에 **전역 하나**였다. 한 파일에서 B열을 키로 잡으면 그 뒤 여는 모든
파일이 B열로 비교됐고, B가 키가 아닌 파일에서는 중복 키로 수천 행이 조용히 빠졌다
(실측: Data_SkillModule_CS.xlsx 를 B열 키로 비교 → 14,659행 제외).

파일별로 기억해 두면 한 번 제대로 잡은 파일은 다음부터 틀리지 않는다.
"""
import pytest

from excelmerge.prefs import (load_key_prefs, load_last_key, save_key_prefs,
                              save_last_key)

KEY_COL = 1
HEADER = ["ID", "GRP", "V"]
A_DATA = [HEADER] + [[f"k{i}", "같은키", f"a{i}"] for i in range(1, 5)]
B_DATA = [HEADER] + [[f"k{i}", "같은키", (f"a{i}" if i % 2 else f"B{i}")]
                     for i in range(1, 5)]


@pytest.fixture(autouse=True)
def isolated_prefs(monkeypatch, tmp_path):
    """전역 설정을 건드리지 않는다 — 실제 %APPDATA% 를 오염시키면 안 된다."""
    monkeypatch.setenv("APPDATA", str(tmp_path))


# ── 저장/조회 ────────────────────────────────────────────────────────────────

def test_roundtrip(tmp_path):
    p = str(tmp_path / "a.xlsx")
    assert load_last_key(p) is None, "처음엔 기억이 없어야 한다"
    save_last_key(p, 1, 3)
    assert load_last_key(p) == (1, 3)


def test_each_file_remembers_its_own_key(tmp_path):
    a, b = str(tmp_path / "a.xlsx"), str(tmp_path / "b.xlsx")
    save_last_key(a, 0, 0)
    save_last_key(b, 1, 5)
    assert load_last_key(a) == (0, 0)
    assert load_last_key(b) == (1, 5)


def test_row_order_mode_is_remembered(tmp_path):
    """키 열 해제(-1)도 그대로 기억해야 한다 — 그 파일은 ROW 순서로 보고 싶다는 뜻."""
    p = str(tmp_path / "a.xlsx")
    save_last_key(p, 0, -1)
    assert load_last_key(p) == (0, -1)


def test_empty_path_is_ignored(tmp_path):
    save_last_key("", 1, 1)          # 예외 없이 무시
    assert load_last_key("") is None


def test_broken_entry_is_ignored(tmp_path, monkeypatch):
    import excelmerge.prefs as prefs
    monkeypatch.setattr(prefs, "_read_prefs",
                        lambda: {"last_keys": {str(tmp_path / "a.xlsx"): "엉터리"}})
    assert load_last_key(str(tmp_path / "a.xlsx")) is None


def test_global_default_is_untouched_by_per_file_memory(tmp_path):
    """파일별 기억은 전역 기본값을 덮어쓰지 않는다 — 처음 여는 파일은 전역을 쓴다."""
    save_key_prefs(0, 0)
    save_last_key(str(tmp_path / "a.xlsx"), 1, 7)
    assert load_key_prefs() == (0, 0)


def test_old_entries_are_dropped_over_the_cap(tmp_path):
    import excelmerge.prefs as prefs
    for i in range(prefs._LAST_SHEETS_MAX + 10):
        save_last_key(str(tmp_path / f"f{i}.xlsx"), 0, i % 5)
    m = prefs._read_prefs()["last_keys"]
    assert len(m) == prefs._LAST_SHEETS_MAX
    assert str(tmp_path / "f0.xlsx") not in m, "오래된 항목이 남았다"
    assert str(tmp_path / f"f{prefs._LAST_SHEETS_MAX + 9}.xlsx") in m


# ── 비교 화면에 실제로 적용되는가 ────────────────────────────────────────────

@pytest.fixture
def dv(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    win = MainWindow()
    win.show()
    try:
        yield win.tabs.currentWidget()
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def _load(view, a_path, b_path, qapp):
    view.panel_a.set_path(a_path)
    view.panel_b.set_path(b_path)
    view._on_loaded(A_DATA, B_DATA)
    w = getattr(view, "_diff_worker", None)
    if w is not None:
        w.wait(8000)
    for _ in range(20):
        qapp.processEvents()


def test_changing_the_key_is_remembered_for_that_file(dv, qapp, tmp_path):
    a, b = str(tmp_path / "x.xlsx"), str(tmp_path / "y.xlsx")
    _load(dv, a, b, qapp)
    dv._on_key_col_changed(2)
    for _ in range(20):
        qapp.processEvents()
    assert load_last_key(a) == (dv._key_row, 2)
    assert load_last_key(b) == (dv._key_row, 2), "B 파일도 함께 기억해야 한다"


def test_remembered_key_is_applied_on_next_load(dv, qapp, tmp_path):
    a, b = str(tmp_path / "x.xlsx"), str(tmp_path / "y.xlsx")
    save_last_key(a, 0, 2)
    save_key_prefs(0, 0)              # 전역 기본은 A열
    _load(dv, a, b, qapp)
    assert dv._key_col == 2, "기억한 키가 아니라 전역 기본이 적용됐다"
    assert dv.panel_a.table._key_col == 2, "표에도 반영돼야 한다"


def test_unknown_file_falls_back_to_the_global_default(dv, qapp, tmp_path):
    save_key_prefs(0, 1)
    dv._key_row, dv._key_col = 0, 1   # 전역을 읽어 온 상태를 흉내
    _load(dv, str(tmp_path / "새파일.xlsx"), "", qapp)
    assert dv._key_col == 1


def test_remembered_key_out_of_range_falls_back(dv, qapp, tmp_path):
    """좁은 시트에 넓은 시트의 키가 기억돼 있으면 기존 안전 처리(A열 리셋)가 받는다."""
    a = str(tmp_path / "x.xlsx")
    save_last_key(a, 0, 99)
    _load(dv, a, "", qapp)
    assert dv._key_col == 0, "범위 밖 키가 그대로 적용됐다"
