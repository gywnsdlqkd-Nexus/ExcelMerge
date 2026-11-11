# -*- coding: utf-8 -*-
"""3파일 병합 — 4차 검증에서 재현된 경우들. 합성 데이터, offscreen."""
import time

import openpyxl
import pytest
from PyQt5.QtWidgets import QApplication, QMessageBox

from excelmerge.diff3_engine import S3_CONFLICT, compute_diff3
from excelmerge.loaders import clear_values_cache, load_values_any


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


HG = ["ID", "Item", "Cnt"]
BASE = [HG, ["A", "grpA", "0"], ["1", "gold", "100"], ["1", "gem", "5"], ["B", "grpB", "0"],
        ["2", "x", "1"]]


@pytest.mark.parametrize("left", [
    [HG, ["1", "gold", "100"], ["1", "gem", "5"], ["B", "grpB", "0"], ["2", "x", "1"]],   # 머리 행 지움
    [HG, ["A2", "grpA", "0"], ["1", "gold", "100"], ["1", "gem", "5"], ["B", "grpB", "0"],
     ["2", "x", "1"]],                                                                   # 머리 행 키 바꿈
])
def test_losing_the_head_row_does_not_split_the_group(left):
    """구간 머리 행을 한쪽이 지우거나 키를 바꾸면 아래 같은 키 행이 '지움+새 행'과 거짓 충돌로 갈렸다."""
    right = [r[:] for r in BASE]
    right[2] = ["1", "gold", "200"]
    m, meta = compute_diff3(left, BASE, right, 0, 0)
    assert S3_CONFLICT not in {c[0] for row in m for c in row}
    gold = next(mm for row, mm in zip(m, meta) if row[1][2] == "gold")
    assert None not in gold, f"gold 행이 세 파일에서 짝지어지지 않았다: {gold}"


def test_blank_key_children_follow_a_head_present_everywhere():
    base = [HG, ["A", "grpA", ""], ["", "gold", "100"], ["B", "grpB", ""], ["", "gold", "100"]]
    left = [HG, ["", "gold", "100"], ["B", "grpB", ""], ["", "gold", "100"]]   # A 를 지움
    right = [HG, ["A", "grpA", ""], ["", "gold", "200"], ["B", "grpB", ""], ["", "gold", "100"]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert S3_CONFLICT not in {c[0] for row in m for c in row}


def test_similarity_breaks_the_tie_between_an_inserted_and_an_edited_row():
    base = [HG, ["1", "gold", "100"]]
    left = [HG, ["1", "potion", "100"], ["1", "gold", "120"]]
    right = [HG, ["1", "gold", "150"]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert (2, 1, 1) in meta, f"고친 gold 가 가운데 gold 와 짝지어져야 한다: {meta}"
    row = m[meta.index((2, 1, 1))]
    assert row[2][0] == S3_CONFLICT, "120 대 150 충돌이 숨었다"


def test_small_groups_with_repeated_content_stay_fast():
    rows = [HG] + [["", "same", "x"] for _ in range(2000)]
    left = [r[:] for r in rows]
    for i in range(1, 2000, 7):
        left[i] = ["", "same", "L"]
    t0 = time.perf_counter()
    compute_diff3(left, rows, rows, 0, 0)
    assert time.perf_counter() - t0 < 5.0


# ── 화면 ────────────────────────────────────────────────────────────────────

@pytest.fixture
def win(qapp, monkeypatch):
    from excelmerge import diff3_view as dv
    from excelmerge.main_window import MainWindow
    shown = []
    for name in ("information", "warning", "critical"):
        monkeypatch.setattr(dv.QMessageBox, name,
                            staticmethod(lambda *a, _n=name, **k: shown.append((_n, a[1]))))
    monkeypatch.setattr(dv.QMessageBox, "question",
                        staticmethod(lambda *a, **k: QMessageBox.Yes))
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


def test_row_order_mode_never_auto_merges(win, tmp_path):
    H = ["ID", "V"]
    l = _xlsx(tmp_path / "l.xlsx", {"Data": [H, [1, "A"], [2, "b"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"Data": [H, [1, "a"], [2, "b"]]})
    r = _xlsx(tmp_path / "r.xlsx", {"Data": [H, [1, "a"], [2, "b"]]})
    view = win.open_merge3(l, b, r)
    assert _wait(lambda: _compared(view))
    view._on_key_col_changed(-1)
    assert _wait(lambda: _compared(view))
    assert view._auto_risk
    view._auto_merge()
    assert view._staged == {}


def test_three_way_save_refuses_when_the_path_box_differs(win, tmp_path):
    H = ["ID", "V"]
    l = _xlsx(tmp_path / "l.xlsx", {"Data": [H, [1, "A"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"Data": [H, [1, "a"]]})
    r = _xlsx(tmp_path / "r.xlsx", {"Data": [H, [1, "a"]]})
    other = _xlsx(tmp_path / "other.xlsx", {"Data": [H, [9, "zz"]]})
    view = win.open_merge3(l, b, r)
    assert _wait(lambda: _compared(view))
    view._auto_merge()
    view.panel_b.path_edit.setText(other)          # Enter 없이
    view._save()
    assert ("warning", "저장하지 않음") in win.shown
    assert _vals(b)[1] == ["1", "a"] and _vals(other)[1] == ["9", "zz"]


# ── 2파일 ───────────────────────────────────────────────────────────────────

def test_two_way_new_file_during_save_does_not_inherit_merged_marks(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv
    from excelmerge.constants import DIR_A2B
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv.QMessageBox, "information", staticmethod(lambda *a, **k: None))
    a = _xlsx(tmp_path / "a.xlsx", {"S": [["ID", "V"], [1, "a1"], [3, "a3"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"S": [["ID", "V"], [1, "b1"], [2, "b2"], [3, "b3"]]})
    a2 = _xlsx(tmp_path / "a2.xlsx", {"S": [["ID", "V"], [1, "z1"], [3, "z3"]]})
    win = MainWindow()
    win.show()
    try:
        view = win.open_file_compare(a, b)
        assert _wait(lambda: bool(view._diff_matrix)
                     and view.status.currentMessage().startswith("비교 완료"))
        r2 = next(r for r, row in enumerate(view._diff_matrix) if row[0][2] == "2")
        view._staged[(r2, 0)] = DIR_A2B            # B 에서 ID2 행을 지운다
        view._staged[(r2, 1)] = DIR_A2B
        view._set_save_btn_state()
        view._save_staged("b")
        view.panel_a.set_path(a2)                   # 저장 중에 A 에 새 파일
        view._on_file_loaded("a", a2)
        assert _wait(lambda: not view._saving_now() and bool(view._diff_matrix)
                     and view.panel_a.get_path() == a2
                     and view.status.currentMessage().startswith("비교 완료"))
        assert not view._merged_cells, "새 비교에 병합 표시가 붙었다"
        keys = [row[0][2] for row in view._diff_matrix[1:]]
        assert "2" not in keys, "저장 전 B 를 읽은 비교가 남았다"
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()
