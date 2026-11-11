# -*- coding: utf-8 -*-
"""3파일 병합 — 2차 검증(1차 수정이 만든 결함)에서 재현된 경우들. 합성 데이터, offscreen."""
import os
import time

import openpyxl
import pytest
from PyQt5.QtCore import QItemSelection, QItemSelectionModel
from PyQt5.QtWidgets import QApplication, QMessageBox

from excelmerge.diff3_engine import (S3_CONFLICT, S3_SAME, SRC_LEFT, SRC_RIGHT, auto_merge_plan,
                                     compute_diff3)
from excelmerge.diff3_workers import write_merge3
from excelmerge.loaders import clear_values_cache, load_values_any

H = ["ID", "V", "W"]


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


# ── 엔진: 같은 키 행은 내용으로 짝짓는다 ────────────────────────────────────

def _by_meta(m, meta):
    return {mm: [c[0] for c in row] for row, mm in zip(m, meta)}


def test_inserting_a_blank_key_row_does_not_shift_the_others():
    """구분 행(키 빈칸) 하나를 끼우면 그 뒤 구분 행이 모두 한 칸씩 밀려 '바뀜'으로 자동 병합됐다."""
    base = [["ID", "N"], ["", "[Weapon]"], ["101", "sword"], ["", "[Armor]"], ["201", "plate"]]
    left = [["ID", "N"], ["", "[New]"], ["", "[Weapon]"], ["101", "sword"], ["", "[Armor]"],
            ["201", "plate"]]
    m, meta = compute_diff3(left, base, base, 0, 0)
    changed = [mm for mm, sts in _by_meta(m, meta).items() if any(s != S3_SAME for s in sts)]
    assert changed == [(1, None, None)], "끼운 행 하나만 바뀜이어야 한다"


def test_duplicate_key_group_edit_and_insert_are_kept_apart():
    """같은 키 묶음 앞에 행을 끼운 쪽과 그 묶음의 행을 고친 쪽 — 고친 값이 끼운 행에 붙었다."""
    base = [["ID", "Item", "Count"], ["1", "gold", "100"], ["1", "gem", "5"], ["2", "gold", "200"]]
    left = [["ID", "Item", "Count"], ["1", "potion", "100"], ["1", "gold", "100"],
            ["1", "gem", "5"], ["2", "gold", "200"]]
    right = [["ID", "Item", "Count"], ["1", "gold", "150"], ["1", "gem", "5"], ["2", "gold", "200"]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    plan = auto_merge_plan(m, meta)
    got = {}
    for r, row in enumerate(m):
        vals = [(c[1] if plan.get((r, i)) == SRC_LEFT else c[3] if plan.get((r, i)) == SRC_RIGHT
                 else c[2]) for i, c in enumerate(row)]
        got[meta[r]] = vals
    assert ["1", "potion", "100"] in got.values()
    assert ["1", "gold", "150"] in got.values()
    assert S3_CONFLICT not in {c[0] for row in m for c in row}


def test_both_sides_adding_different_rows_to_a_group_keeps_both():
    base = [["ID", "Item"], ["1", "gold"]]
    left = [["ID", "Item"], ["1", "gold"], ["1", "potion"]]
    right = [["ID", "Item"], ["1", "gold"], ["1", "elixir"]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert S3_CONFLICT not in {c[0] for row in m for c in row}
    assert sum(1 for mm in meta if mm[1] is None) == 2


def test_cells_equal_to_the_middle_are_not_written(tmp_path):
    """남긴 쪽 행을 통째로 골라도 아무도 안 바꾼 수식 칸은 수식으로 남는다."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in (H, [1, "a", "p"], [2, '="b"&""', "q"]):
        ws.append(row)
    base = str(tmp_path / "b.xlsx")
    wb.save(base)
    from excelmerge.merge3_build import build_merge3_patches
    m = [[("same", "ID", "ID", "ID"), ("same", "V", "V", "V"), ("same", "W", "W", "W")],
         [("same", "1", "1", "1"), ("same", "a", "a", "a"), ("same", "p", "p", "p")],
         [("conflict", "", "2", "2"), ("conflict", "", "b", "b"), ("conflict", "", "q", "Q")]]
    meta = [(0, 0, 0), (1, 1, 1), (None, 2, 2)]
    cm = [(0, 0, 0), (1, 1, 1), (2, 2, 2)]
    patches, _ins, _ss = build_merge3_patches(m, meta, cm, {(2, 0): SRC_RIGHT, (2, 1): SRC_RIGHT,
                                                           (2, 2): SRC_RIGHT})
    assert patches == {"C3": "Q"}, "기준과 같은 값(수식 칸)까지 썼다"


# ── 화면 ────────────────────────────────────────────────────────────────────

ANSWER = {"default": QMessageBox.Yes}


@pytest.fixture
def win(qapp, monkeypatch):
    from excelmerge import diff3_view as dv
    from excelmerge.main_window import MainWindow
    shown = []
    for name in ("information", "warning", "critical"):
        monkeypatch.setattr(dv.QMessageBox, name,
                            staticmethod(lambda *a, _n=name, **k: shown.append((_n, a[1]))))

    def question(*a, **k):
        shown.append(("question", a[1]))
        return ANSWER.get(a[1], ANSWER["default"])
    monkeypatch.setattr(dv.QMessageBox, "question", staticmethod(question))
    ANSWER.clear()
    ANSWER["default"] = QMessageBox.Yes
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


def test_picking_a_conflict_in_a_new_row_brings_the_rest_of_the_row(win, tmp_path):
    base = [H, [1, "a", "p"]]
    left = [H, [1, "a", "p"], [7, "same", "L"]]
    right = [H, [1, "a", "p"], [7, "same", "R"]]
    view, l, b, r, res = _open(win, tmp_path, {"Data": left}, {"Data": base}, {"Data": right},
                               result=True)
    r7 = next(i for i, mm in enumerate(view._row_meta) if mm == (2, None, 2))
    assert not any(k[0] == r7 for k in view._staged)
    _select(view, [(r7, 2)])
    view._stage_selected(SRC_RIGHT)
    assert {c for (rr, c) in view._staged if rr == r7} == {0, 1, 2}
    _save(view)
    assert _vals(res) == [H, ["1", "a", "p"], ["7", "same", "R"]]


def test_unstaging_the_key_of_a_new_row_unstages_the_row(win, tmp_path):
    base = [H, [1, "a", "p"]]
    left = [H, [1, "a", "p"], [7, "v7", "w7"]]
    view, l, b, r, res = _open(win, tmp_path, {"Data": left}, {"Data": base}, {"Data": base},
                               result=True)
    r7 = next(i for i, mm in enumerate(view._row_meta) if mm == (2, None, None))
    assert {c for (rr, c) in view._staged if rr == r7} == {0, 1, 2}
    _select(view, [(r7, 0)])
    view._unstage_selected()
    assert not any(k[0] == r7 for k in view._staged)


def test_p4v_save_asks_about_added_or_deleted_sheets(win, tmp_path):
    view, l, b, r, res = _open(
        win, tmp_path, {"Data": [H, [1, "A", "p"]], "Added": [H, [100, "new", "x"]]},
        {"Data": [H, [1, "a", "p"]]}, {"Data": [H, [1, "a", "p"]]}, result=True)
    assert _wait(lambda: view._changed_sheets is not None)
    ANSWER["결과에 빠지는 변경"] = QMessageBox.No
    _save(view)
    assert ("question", "결과에 빠지는 변경") in win.shown
    assert not os.path.exists(res), "시트가 빠진다는 걸 묻지 않고 썼다"


def test_carry_completes_a_deletion_row_that_lost_columns(win, tmp_path):
    """키를 바꿔 결정을 되살릴 때 열 일부를 못 찾아 삭제 행이 반쪽만 걸렸다."""
    base = [H, [1, "a", "p"], [2, "b", "q"], [3, "c", "r"]]
    left = [["ID", "V"], [1, "a"], [2, "b"]]
    view, l, b, r, res = _open(win, tmp_path, {"Data": left}, {"Data": base}, {"Data": base},
                               result=True)
    r3 = next(i for i, mm in enumerate(view._row_meta) if mm == (None, 3, 3))
    assert {c for (rr, c) in view._staged if rr == r3} == {0, 1, 2}
    view._on_key_col_changed(-1)
    assert _wait(lambda: _compared(view))
    r3 = next(i for i, mm in enumerate(view._row_meta) if mm == (None, 3, 3))
    assert {c for (rr, c) in view._staged if rr == r3} == {0, 1, 2}
    _save(view)
    got = _vals(res)
    assert not any(row and row[0] == "" for row in got[1:]), f"키 없는 행이 남았다: {got}"


def test_positional_fallback_is_restored_with_the_sheet(win, tmp_path):
    base = [["ID", "N"], [1, "a"]]
    left = [["G", "ID", "N"], ["g", 1, "a"]]
    view, *_ = _open(win, tmp_path, {"Data": left, "S2": [H, [1, "A", "p"]]},
                     {"Data": base, "S2": [H, [1, "a", "p"]]},
                     {"Data": [["N", "X"], ["a", 1]], "S2": [H, [1, "a", "p"]]})
    assert view._auto_risk
    _select(view, [(1, 1)])
    view._stage_selected(SRC_LEFT)
    _goto_sheet(view, "S2")
    _goto_sheet(view, "Data")
    assert view._auto_risk, "시트를 다녀오자 위치 기준 표시가 풀렸다"


def test_key_change_while_a_sheet_is_loading_is_refused(win, tmp_path):
    view, *_ = _open(win, tmp_path, {"Data": [H, [1, "A", "p"]], "S2": [H, [1, "LX", "y"]]},
                     {"Data": [H, [1, "a", "p"]], "S2": [H, [1, "x", "y"]]},
                     {"Data": [H, [1, "a", "p"]], "S2": [H, [1, "RX", "y"]]})
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("S2"))
    assert view._loading
    view._on_key_col_changed(-1)
    assert view._key_col == 0
    assert _wait(lambda: view._current_sheet == "S2" and _compared(view))
    assert view._staged == {}


def test_both_changed_cells_come_from_the_side_that_has_the_sheet(win, tmp_path):
    """왼쪽이 지운 시트에서 오른쪽이 지운 행('양쪽 같게')이 자동 병합에서 빠졌다."""
    old_b = [H, [7, "x", "y"], [8, "z", "w"]]
    old_r = [H, [7, "x", "y"]]
    view, l, b, r, res = _open(win, tmp_path, {"Data": [H, [1, "a", "p"]]},
                               {"Data": [H, [1, "a", "p"]], "Old": old_b},
                               {"Data": [H, [1, "a", "p"]], "Old": old_r}, result=True)
    assert _wait(lambda: view._changed_sheets is not None)
    _goto_sheet(view, "Old")
    view._auto_merge()
    r8 = next(i for i, mm in enumerate(view._row_meta) if mm == (None, 2, None))
    assert {view._staged.get((r8, c)) for c in range(3)} == {SRC_RIGHT}


# ── 2파일 ───────────────────────────────────────────────────────────────────

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


def test_two_way_excluded_columns_survive_the_save_reload(win2, tmp_path):
    from excelmerge.constants import DIR_A2B
    hdr = ["ID", "V", "#", "W", "#"]
    a = _xlsx(tmp_path / "a.xlsx", {"S": [hdr, [1, "a1", "c", "w", "x1"], [2, "v", "c", "w", "x2"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"S": [hdr, [1, "b1", "C", "w", "x1"],
                                          [2, "v", "c", "w", "X2"]]})
    view = win2.open_file_compare(a, b)
    assert _wait(lambda: bool(view._diff_matrix) and view.status.currentMessage().startswith("비교 완료"))
    view._on_columns_exclude_set([2], True)
    view._staged[(1, 1)] = DIR_A2B
    view._set_save_btn_state()
    view._save_staged("b")
    assert _wait(lambda: not view._saving_now() and view._post_save_carry is None
                 and view.status.currentMessage().startswith("비교 완료"))
    assert view._excluded_cols == {2}, f"제외 열이 바뀌었다: {view._excluded_cols}"


def test_two_way_sheet_switch_is_refused_while_saving(win2, tmp_path):
    from excelmerge.constants import DIR_A2B
    a = _xlsx(tmp_path / "a.xlsx", {"S1": [["ID", "V"], [1, "a1"]], "S2": [["ID", "V"], [1, "s2a"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"S1": [["ID", "V"], [1, "b1"]], "S2": [["ID", "V"], [1, "s2b"]]})
    view = win2.open_file_compare(a, b)
    assert _wait(lambda: bool(view._diff_matrix) and view._current_sheet == "S1"
                 and view.status.currentMessage().startswith("비교 완료"))
    view._staged[(1, 1)] = DIR_A2B
    view._set_save_btn_state()
    view._save_staged("b")
    assert view._saving_now()
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("S2"))
    assert view._current_sheet == "S1"
    assert _wait(lambda: not view._saving_now() and view.status.currentMessage().startswith("비교 완료"))
    assert view._current_sheet == "S1"
