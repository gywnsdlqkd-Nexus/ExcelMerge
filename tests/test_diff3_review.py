# -*- coding: utf-8 -*-
"""3파일 병합 — 적대적 검토에서 재현된 결함들. 하나하나가 실제로 틀린 결과를 냈던 경우다.

합성 데이터만 쓴다(공개 저장소). 화면은 띄우지 않는다(offscreen).
"""
import os
import shutil
import stat
import time

import openpyxl
import pytest
from PyQt5.QtCore import QItemSelection, QItemSelectionModel
from PyQt5.QtWidgets import QApplication

from excelmerge.diff3_engine import (S3_CONFLICT, SRC_LEFT, SRC_RIGHT, auto_merge_plan,
                                     compute_diff3, match_columns3, usable_col_meta3)
from excelmerge.diff3_workers import write_merge3
from excelmerge.loaders import clear_values_cache, load_values_any


# ── 도구 ────────────────────────────────────────────────────────────────────

def _xlsx(path, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    wb.save(str(path))
    return str(path)


def _vals(path, sheet="Data"):
    clear_values_cache()
    return load_values_any(path, None, sheet)


def _wait(pred, timeout=20.0):
    app = QApplication.instance()
    end = time.time() + timeout
    while time.time() < end:
        app.processEvents()
        if pred():
            return True
        time.sleep(0.01)
    return bool(pred())


def _engine_merge(tmp_path, left, base, right, key_col=0):
    """엔진 + 자동 병합 + 저장(P4V 결과 파일) → 결과 값."""
    l = _xlsx(tmp_path / "l.xlsx", {"Data": left})
    b = _xlsx(tmp_path / "b.xlsx", {"Data": base})
    r = _xlsx(tmp_path / "r.xlsx", {"Data": right})
    data = [_vals(p) for p in (l, b, r)]
    cm = match_columns3(*data, 0)
    m, meta = compute_diff3(*data, key_col, 0, usable_col_meta3(cm, key_col))
    out = str(tmp_path / "out.xlsx")
    write_merge3(b, out, [("Data", m, meta, usable_col_meta3(cm, key_col),
                           auto_merge_plan(m, meta))], {SRC_LEFT: l, SRC_RIGHT: r})
    return _vals(out), m, meta


# ── 엔진: 보이지 않던 곳이 기준값으로 되돌아가지 않는다 ─────────────────────

def test_unnamed_trailing_column_is_compared_and_merged(tmp_path):
    """이름 없는 꼬리 열의 양쪽 수정이 결과에서 사라졌다(P4V 결과는 원본에서 다시 만든다)."""
    base = [["ID", "V", ""], [1, "a", "memo1"], [2, "b", "memo2"]]
    left = [["ID", "V", ""], [1, "a", "THEIR memo"], [2, "b", "memo2"]]
    right = [["ID", "V", ""], [1, "a", "memo1"], [2, "B", "MY memo"]]
    got, _m, _meta = _engine_merge(tmp_path, left, base, right)
    assert got == [["ID", "V", ""], ["1", "a", "THEIR memo"], ["2", "B", "MY memo"]]


def test_duplicate_and_blank_key_rows_are_matched_by_order(tmp_path):
    base = [["ID", "V"], [1, "a"], [1, "dup-base"], ["", "note-base"], [2, "b"]]
    left = [["ID", "V"], [1, "a"], [1, "dup-THEIRS"], ["", "note-base"], [2, "b"]]
    right = [["ID", "V"], [1, "a"], [1, "dup-base"], ["", "note-MINE"], [2, "B"]]
    got, _m, _meta = _engine_merge(tmp_path, left, base, right)
    assert got == [["ID", "V"], ["1", "a"], ["1", "dup-THEIRS"], ["", "note-MINE"], ["2", "B"]]


def test_column_change_on_one_side_does_not_make_deleted_rows_conflicts(tmp_path):
    """오른쪽이 열을 더했고 왼쪽이 행을 지웠다 — 그 행은 충돌이 아니라 왼쪽 삭제다."""
    base = [["ID", "N", "V"], [1, "a", 10], [2, "b", 20], [3, "c", 30]]
    left = [["ID", "N", "V"], [1, "a", 10], [3, "c", 30]]
    right = [["ID", "N", "V", "New"], [1, "a", 10, "x"], [2, "b", 20, "x"], [3, "c", 30, "x"]]
    got, m, meta = _engine_merge(tmp_path, left, base, right)
    row2 = next(i for i, mm in enumerate(meta) if mm == (None, 2, 2))
    assert S3_CONFLICT not in [c[0] for c in m[row2]]
    assert got == [["ID", "N", "V"], ["1", "a", "10"], ["3", "c", "30"]]


def test_new_row_with_a_conflict_is_left_whole_to_the_user():
    """양쪽이 같은 키로 다르게 만든 새 행 — 키 칸만 가져가면 키만 있는 빈 행이 생겼다."""
    base = [["ID", "N"], ["1", "a"]]
    left = [["ID", "N"], ["1", "a"], ["7", "L"]]
    right = [["ID", "N"], ["1", "a"], ["7", "R"]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    plan = auto_merge_plan(m, meta)
    row7 = next(i for i, mm in enumerate(meta) if mm == (2, None, 2))
    assert not any(r == row7 for r, _c in plan)


# ── 화면 ────────────────────────────────────────────────────────────────────

ANSWER = [None]


@pytest.fixture
def win(qapp, monkeypatch):
    from excelmerge import diff3_view as dv
    from excelmerge.main_window import MainWindow
    shown = []
    for name in ("information", "warning", "critical"):
        monkeypatch.setattr(dv.QMessageBox, name,
                            staticmethod(lambda *a, _n=name, **k: shown.append((_n, a[1]))))
    monkeypatch.setattr(dv.QMessageBox, "question",
                        staticmethod(lambda *a, **k: (shown.append(("question", a[1])),
                                                      ANSWER[0])[1]))
    ANSWER[0] = dv.QMessageBox.Yes
    w = MainWindow()
    w.show()
    w.shown = shown
    yield w
    w.close()
    w.deleteLater()
    qapp.processEvents()


def _compared(view):
    return bool(view._matrix) and not view._computing and \
        view.status.currentMessage().startswith("비교 완료")


def _open(win, tmp_path, left, base, right, result=False):
    l = _xlsx(tmp_path / "left.xlsx", left)
    b = _xlsx(tmp_path / "base.xlsx", base)
    r = _xlsx(tmp_path / "right.xlsx", right)
    res = str(tmp_path / "result.xlsx") if result else ""
    view = win.open_merge3(l, b, r, res)
    assert _wait(lambda: _compared(view)), view.status.currentMessage()
    return view, l, b, r, res


def _select(view, cells, role="b"):
    t = view.panels[role].table
    sel = QItemSelection()
    for r, c in cells:
        idx = t.model().index(r, c)
        sel.select(idx, idx)
    t.selectionModel().select(sel, QItemSelectionModel.ClearAndSelect)


def _save(view):
    view._save()
    _wait(lambda: not view._saving_now() and not view.status.currentMessage().startswith("저장 중"))
    QApplication.instance().processEvents()


def _goto_sheet(view, name):
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index(name))
    assert _wait(lambda: view._current_sheet == name and _compared(view))


H = ["ID", "V", "W"]
DATA_B = [H, [1, "a", "p"], [2, "b", "q"]]
DATA_L = [H, [1, "A", "p"], [2, "b", "q"]]
S2_B = [H, [7, "x", "y"], [8, "z", "w"]]
S2_L = [H, [7, "X", "y"], [8, "z", "w"]]


def test_key_follows_the_middle_files_column_when_a_column_is_inserted_before_it(win, tmp_path):
    """왼쪽이 키 열 앞에 열을 끼웠다 — 위치 기준으로 떨어져 진짜 행을 지우던 문제."""
    base = [["ID", "Name", "Val"], [1, "a", 10], [2, "b", 20], [3, "c", 30]]
    left = [["Group", "ID", "Name", "Val"], ["g", 1, "a", 10], ["g", 2, "b", 20], ["g", 3, "c", 30]]
    right = [["ID", "Name", "Val"], [1, "a", 10], [2, "b", 99], [3, "c", 30]]
    view, l, b, r, res = _open(win, tmp_path, {"Data": left}, {"Data": base}, {"Data": right},
                               result=True)
    assert view._key_col == 1 and not view._auto_risk
    _save(view)
    assert _vals(res) == [["ID", "Name", "Val"], ["1", "a", "10"], ["2", "b", "99"],
                          ["3", "c", "30"]]


def test_carry_does_not_leak_into_another_sheet(win, tmp_path):
    """키를 바꿔 다시 비교하는 중에 시트를 옮기면 앞 시트의 결정이 다른 시트에 붙었다."""
    view, l, b, r, _ = _open(win, tmp_path, {"Data": DATA_L, "S2": S2_L},
                             {"Data": DATA_B, "S2": S2_B}, {"Data": DATA_B, "S2": S2_B})
    _select(view, [(1, 1)])
    view._stage_selected(SRC_LEFT)
    view._on_key_col_changed(-1)                 # 다시 비교 시작 — 끝나기 전에
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("S2"))
    assert _wait(lambda: view._current_sheet == "S2" and _compared(view))
    assert view._staged == {}, "손대지 않은 시트에 앞 시트의 결정이 붙었다"
    assert view._sheet_cache["Data"]["key"] == (0, 0), "보관한 시트의 키가 매트릭스와 다르다"


def test_nothing_changes_while_saving(win, tmp_path):
    view, l, b, r, _ = _open(win, tmp_path, {"Data": DATA_L}, {"Data": DATA_B},
                             {"Data": [H, [1, "a", "p"], [2, "b", "Q"]]})
    _select(view, [(1, 1)])
    view._stage_selected(SRC_LEFT)
    view._save()
    assert view._saving_now()
    _select(view, [(2, 2)])
    view._stage_selected(SRC_RIGHT)              # 저장 중 — 받지 않는다
    _wait(lambda: not view._saving_now() and view.status.currentMessage().startswith("저장 완료"))
    assert (2, 2) not in view._merged and (2, 2) not in view._staged
    assert _vals(b)[1] == ["1", "A", "p"]


def test_sheet_missing_from_the_middle_is_compare_only(win, tmp_path):
    """왼쪽에만 있는 시트 — 자동 준비했다가 저장이 언제나 실패했다."""
    view, l, b, r, res = _open(win, tmp_path, {"Data": DATA_L, "New": S2_B},
                               {"Data": DATA_B}, {"Data": DATA_B}, result=True)
    assert _wait(lambda: view._changed_sheets is not None)
    assert view._unvisited_changed_sheets() == []
    _goto_sheet(view, "New")
    assert view._staged == {}
    _select(view, [(1, 1)], role="l")
    view._stage_selected(SRC_LEFT)
    assert view._staged == {}
    _goto_sheet(view, "Data")
    _save(view)
    assert _vals(res)[1] == ["1", "A", "p"]


def test_sheet_deleted_on_one_side_is_not_emptied(win, tmp_path):
    """왼쪽에서 지운 시트 — 자동 병합이 머리글까지 지워 시트가 비었다."""
    view, l, b, r, res = _open(win, tmp_path, {"Data": DATA_L},
                               {"Data": DATA_B, "Old": S2_B}, {"Data": DATA_B, "Old": S2_L},
                               result=True)
    assert _wait(lambda: view._changed_sheets is not None)
    _goto_sheet(view, "Old")
    assert view._staged == {}
    _select(view, [(1, 1)])
    view._stage_selected(SRC_RIGHT)              # 남은 쪽(오른쪽) 값은 가져올 수 있다
    assert view._staged == {(1, 1): SRC_RIGHT}
    _goto_sheet(view, "Data")
    _save(view)
    assert _vals(res, "Old") == [["ID", "V", "W"], ["7", "X", "y"], ["8", "z", "w"]]


def test_picking_the_keeping_side_releases_the_deletion(win, tmp_path):
    """지운 쪽을 행째 고른 뒤 남긴 쪽 칸 하나를 고르면 키가 빈 반쪽 행이 저장됐다."""
    base = [H, [1, "a", "p"], [2, "b", "q"], [3, "c", "r"]]
    left = [H, [1, "a", "p"], [3, "c", "r"]]
    right = [H, [1, "a", "p"], [2, "b", "Q"], [3, "c", "r"]]
    view, l, b, r, _ = _open(win, tmp_path, {"Data": left}, {"Data": base}, {"Data": right})
    r2 = next(i for i, mm in enumerate(view._row_meta) if mm == (None, 2, 2))
    _select(view, [(r2, 0), (r2, 1), (r2, 2)])
    view._stage_selected(SRC_LEFT)
    _select(view, [(r2, 2)])
    view._stage_selected(SRC_RIGHT)
    assert view._staged == {(r2, 2): SRC_RIGHT}
    _save(view)
    assert _vals(b) == [H, ["1", "a", "p"], ["2", "b", "Q"], ["3", "c", "r"]]


def test_unchanged_cells_of_a_row_conflict_can_be_resolved(win, tmp_path):
    base = [H, [1, "a", "p"], [2, "b", "q"]]
    left = [H, [1, "a", "p"]]
    right = [H, [1, "a", "p"], [2, "b", "Q"]]
    view, *_ = _open(win, tmp_path, {"Data": left}, {"Data": base}, {"Data": right})
    r2 = next(i for i, mm in enumerate(view._row_meta) if mm == (None, 2, 2))
    _select(view, [(r2, 0), (r2, 1), (r2, 2)], role="r")
    view._stage_selected(SRC_RIGHT)
    assert len(view._staged) == 3
    assert "(남은 0)" in view._summary()


def test_p4v_save_asks_about_open_conflicts_on_other_sheets(win, tmp_path):
    other_b = [H, [1, "x", "y"]]
    view, l, b, r, res = _open(win, tmp_path, {"Data": DATA_L, "Other": [H, [1, "LX", "y"]]},
                               {"Data": DATA_B, "Other": other_b},
                               {"Data": DATA_B, "Other": [H, [1, "RX", "y"]]}, result=True)
    assert _wait(lambda: view._changed_sheets is not None)
    _goto_sheet(view, "Other")
    _goto_sheet(view, "Data")
    ANSWER[0] = __import__("PyQt5.QtWidgets", fromlist=["QMessageBox"]).QMessageBox.No
    _save(view)
    assert ("question", "결과에 빠지는 변경") in win.shown
    assert not os.path.exists(res), "남은 충돌을 묻지 않고 썼다"


def test_read_only_middle_file_leaves_no_temp_copy(win, tmp_path):
    view, l, b, r, _ = _open(win, tmp_path, {"Data": DATA_L}, {"Data": DATA_B}, {"Data": DATA_B})
    os.chmod(b, stat.S_IREAD)
    try:
        view._run_refresh()
        assert _wait(lambda: _compared(view))
        tmpdir = view._tmpdir
        assert os.access(view._template, os.W_OK)
        view.shutdown()
        assert not os.path.exists(tmpdir), "읽기 전용 사본이 임시 폴더에 남았다"
    finally:
        os.chmod(b, stat.S_IWRITE | stat.S_IREAD)


# ── 2파일 회귀 ──────────────────────────────────────────────────────────────

@pytest.fixture
def win2(qapp, monkeypatch):
    from excelmerge import diff_view as dv
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv.QMessageBox, "information", staticmethod(lambda *a, **k: None))
    w = MainWindow()
    w.show()
    yield w
    w.close()
    w.deleteLater()
    qapp.processEvents()


def test_two_way_post_save_carry_stays_on_its_sheet(win2, tmp_path):
    from excelmerge.constants import DIR_A2B, DIR_B2A
    a = _xlsx(tmp_path / "a.xlsx", {"S1": [["ID", "V1", "V2"], [1, "a1", "x"], [2, "y", "a2"]],
                                    "S2": [["ID", "V1", "V2"], [1, "s2a1", "p"], [2, "q", "s2a2"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"S1": [["ID", "V1", "V2"], [1, "b1", "x"], [2, "y", "b2"]],
                                    "S2": [["ID", "V1", "V2"], [1, "s2b1", "p"], [2, "q", "s2b2"]]})
    view = win2.open_file_compare(a, b)
    assert _wait(lambda: bool(view._diff_matrix) and view._current_sheet == "S1"
                 and view.status.currentMessage().startswith("비교 완료"))
    view._staged[(1, 1)] = DIR_A2B
    view._staged[(2, 2)] = DIR_B2A
    view._set_save_btn_state()
    view._save_staged("b")
    _wait(lambda: not view._saving_now() and view._post_save_carry is not None)
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("S2"))   # 다시 읽기 끝나기 전에
    assert _wait(lambda: view._current_sheet == "S2" and bool(view._diff_matrix)
                 and view.status.currentMessage().startswith("비교 완료"))
    assert view._staged == {} and not view._merged_cells, "다른 시트에 준비·병합이 붙었다"


def test_two_way_staging_during_save_stays_staged(win2, tmp_path):
    from excelmerge.constants import DIR_A2B
    a = _xlsx(tmp_path / "a.xlsx", {"S": [["ID", "V"], [1, "a1"], [2, "a2"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"S": [["ID", "V"], [1, "b1"], [2, "b2"]]})
    view = win2.open_file_compare(a, b)
    assert _wait(lambda: bool(view._diff_matrix) and view.status.currentMessage().startswith("비교 완료"))
    view._staged[(1, 1)] = DIR_A2B
    view._set_save_btn_state()
    view._save_staged("b")
    view._staged[(2, 1)] = DIR_A2B                # 저장 중에 건 준비
    _wait(lambda: not view._saving_now() and view._post_save_carry is None
          and view.status.currentMessage().startswith("비교 완료"))
    r2 = next(r for r, row in enumerate(view._diff_matrix) if row[0][1] == "2")
    assert view._staged.get((r2, 1)) == DIR_A2B, "저장 중에 건 준비가 병합됨으로 사라졌다"
    assert _vals(b, "S")[2] == ["2", "b2"]
