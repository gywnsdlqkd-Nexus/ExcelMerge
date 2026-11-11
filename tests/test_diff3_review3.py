# -*- coding: utf-8 -*-
"""3파일 병합 — 3차 검증에서 재현된 경우들. 합성 데이터, offscreen."""
import os
import time

import openpyxl
import pytest
from PyQt5.QtCore import QItemSelection, QItemSelectionModel
from PyQt5.QtWidgets import QApplication, QMessageBox

from excelmerge.diff3_engine import S3_CONFLICT, S3_SAME, SRC_LEFT, compute_diff3
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


# ── 엔진: 같은 키 행 짝짓기 ─────────────────────────────────────────────────

def test_blank_key_rows_are_matched_within_their_section():
    """같은 내용의 빈 키 행이 여러 구간에 있으면 엉뚱한 구간 것이 지워졌다."""
    base = [["ID", "N", "V"], ["s1", "stage1", ""], ["", "gold", "100"], ["s2", "stage2", ""],
            ["", "gold", "100"]]
    left = [["ID", "N", "V"], ["s1", "stage1", ""], ["s2", "stage2", ""], ["", "gold", "100"]]
    m, meta = compute_diff3(left, base, base, 0, 0)
    gone = [mm for mm in meta if mm[0] is None]
    assert gone == [(None, 2, 2)], f"stage1 의 gold(가운데 2행)가 지워져야 한다: {gone}"


def test_a_column_added_with_values_does_not_break_the_alignment():
    base = [["ID", "N"], ["", "[A]"], ["", "[B]"]]
    left = [["ID", "N", "New"], ["", "[New]", "x"], ["", "[A]", "x"], ["", "[B]", "x"]]
    m, meta = compute_diff3(left, base, base, 0, 0)
    assert (1, None, None) in meta and (2, 1, 1) in meta and (3, 2, 2) in meta, meta


def test_insert_and_edit_in_a_duplicate_group_pair_by_similarity():
    hdr = ["ID", "Item", "Count", "Grade"]
    base = [hdr, ["1", "gold", "100", "A"], ["1", "gem", "5", "A"]]
    left = [hdr, ["1", "potion", "1", "A"], ["1", "gold", "150", "A"], ["1", "gem", "5", "A"]]
    m, meta = compute_diff3(left, base, base, 0, 0)
    assert (2, 1, 1) in meta, f"고친 gold 가 가운데 gold 와 짝지어져야 한다: {meta}"
    assert (1, None, None) in meta


def test_many_blank_keys_with_repeated_content_stay_fast():
    rows = [["ID", "N"]] + [["", f"v{i % 3}"] for i in range(20000)]
    left = rows[:5000] + [["", "new"]] + rows[5000:]
    amb = set()
    t0 = time.perf_counter()
    m, _meta = compute_diff3(left, rows, rows, 0, 0, None, amb)
    took = time.perf_counter() - t0
    assert took < 10.0, f"{took:.1f}초"
    assert len(amb) == len(m) - 1, "내용으로 짝지은 행은 모두 '불확실'로 표시돼야 한다"


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
        shown.append(("question", a[1], a[2]))
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


def _save(view):
    view._save()
    _wait(lambda: not view._saving_now() and not view.status.currentMessage().startswith("저장 중"))
    QApplication.instance().processEvents()


def _goto_sheet(view, name):
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index(name))
    assert _wait(lambda: view._current_sheet == name and _compared(view))


H = ["ID", "V", "W"]


def test_p4v_save_lists_a_column_that_cannot_be_merged(win, tmp_path):
    """theirs 가 더한 열은 가운데에 없어 병합할 수 없다 — 묻지 않고 빼던 문제."""
    base = [H, [1, "a", "p"], [2, "b", "q"]]
    left = [["ID", "V", "Stat", "W"], [1, "a", "s1", "p"], [2, "b", "s2", "q"]]
    right = [H, [1, "a", "p"], [2, "b", "Q"]]
    view, l, b, r, res = _open(win, tmp_path, {"Data": left}, {"Data": base}, {"Data": right},
                               result=True)
    ANSWER["결과에 빠지는 변경"] = QMessageBox.No
    _save(view)
    asked = [q for q in win.shown if q[0] == "question" and q[1] == "결과에 빠지는 변경"]
    assert asked and "병합할 수 없는 칸" in asked[0][2], asked
    assert not os.path.exists(res)


def test_three_way_never_starts_in_row_order_mode(win, tmp_path):
    from excelmerge.prefs import save_key_prefs
    save_key_prefs(0, -1)
    view, *_ = _open(win, tmp_path, {"Data": [H, [1, "A", "p"]]}, {"Data": [H, [1, "a", "p"]]},
                     {"Data": [H, [1, "a", "p"]]})
    assert view._key_col == 0


def test_positional_columns_with_different_headers_block_auto_merge(win, tmp_path):
    """머리글 가운데 빈 칸으로 열을 위치로 맞춘 시트에서 한쪽이 열을 끼우면 값이 밀려 써졌다."""
    base = [["ID", "", "W"], [1, "m1", "p"], [2, "m2", "q"]]
    left = [["ID", "NEW", "", "W"], [1, "n1", "m1", "p"], [2, "n2", "m2", "q"]]
    right = [["ID", "", "W"], [1, "m1", "p"], [2, "m2", "Q"]]
    view, l, b, r, res = _open(win, tmp_path, {"Data": left}, {"Data": base}, {"Data": right},
                               result=True)
    assert view._auto_risk
    assert view._staged == {}, "위치로 맞춘 열을 자동 병합 준비했다"


def test_ambiguous_rows_are_not_auto_merged(win, tmp_path):
    hdr = ["ID", "Item", "Count"]
    base = [hdr, [1, "gold", 100], [1, "gem", 5], [2, "x", 1]]
    left = [hdr, [1, "potion", 1], [1, "gold", 150], [1, "gem", 5], [2, "x", 1]]
    right = [hdr, [1, "gold", 100], [1, "gem", 5], [2, "x", 9]]
    view, *_ = _open(win, tmp_path, {"Data": left}, {"Data": base}, {"Data": right}, result=True)
    assert view._ambiguous
    assert not any(k[0] in view._ambiguous for k in view._staged), "불확실한 행을 자동 준비했다"
    r2 = next(i for i, mm in enumerate(view._row_meta) if mm == (4, 3, 3))
    assert view._staged.get((r2, 2)) is not None, "확실한 행(키 2)은 자동 준비돼야 한다"


def test_risk_of_the_stashed_sheet_matches_its_matrix(win, tmp_path):
    """키를 바꾸고 다시 비교가 끝나기 전에 시트를 옮겼다 돌아오면 위험 표시가 새 키 것이었다."""
    base = [["ID", "A", "B"], [1, "x", "y"]]
    left = [["Key", "A", "New", "B"], [1, "x", "n", "y"]]
    view, *_ = _open(win, tmp_path, {"Data": left, "S2": [H, [1, "L", "p"]]},
                     {"Data": base, "S2": [H, [1, "a", "p"]]},
                     {"Data": base, "S2": [H, [1, "a", "p"]]})
    risk = view._auto_risk
    assert risk
    cell = next((r, c) for r, row in enumerate(view._matrix) for c, x in enumerate(row)
                if x[0] != S3_SAME and x[1] != x[2])
    t = view.panel_b.table
    idx = t.model().index(*cell)
    sel = QItemSelection()
    sel.select(idx, idx)
    t.selectionModel().select(sel, QItemSelectionModel.ClearAndSelect)
    view._stage_selected(SRC_LEFT)
    assert view._staged, "고른 칸이 준비되지 않았다"
    view._on_key_col_changed(1)                  # 다시 비교 — 끝나기 전에
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("S2"))
    assert _wait(lambda: view._current_sheet == "S2" and _compared(view))
    assert view._sheet_cache["Data"]["risk"] == risk
    _goto_sheet(view, "Data")
    assert view._auto_risk == risk


# ── 2파일 ───────────────────────────────────────────────────────────────────

@pytest.fixture
def win2(qapp, monkeypatch):
    from excelmerge import diff_view as dv
    from excelmerge.main_window import MainWindow
    seen = []
    for name in ("information", "warning"):
        monkeypatch.setattr(dv.QMessageBox, name,
                            staticmethod(lambda *a, _n=name, **k: seen.append((_n, a[1]))))
    w = MainWindow()
    w.show()
    w.seen = seen
    yield w
    w.close()
    w.deleteLater()
    qapp.processEvents()


def test_two_way_save_is_refused_when_the_path_box_differs(win2, tmp_path):
    from excelmerge.constants import DIR_A2B
    a = _xlsx(tmp_path / "a.xlsx", {"S": [["ID", "V"], [1, "a1"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"S": [["ID", "V"], [1, "b1"]]})
    view = win2.open_file_compare(a, b)
    assert _wait(lambda: bool(view._diff_matrix) and view.status.currentMessage().startswith("비교 완료"))
    view._staged[(1, 1)] = DIR_A2B
    view.panel_b.path_edit.setText(str(tmp_path / "other.xlsx"))   # Enter 없이
    view._save_staged("b")
    assert ("warning", "저장하지 않음") in win2.seen
    assert _vals(b, "S")[1] == ["1", "b1"]


def test_two_way_unreloadable_save_drops_staged_and_blocks_saving(win2, tmp_path):
    from excelmerge.constants import DIR_A2B, DIR_B2A
    a = _xlsx(tmp_path / "a.xlsx", {"S": [["ID", "V"], [1, "a1"], [2, "a2"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"S": [["ID", "V"], [1, "b1"], [2, "b2"]]})
    view = win2.open_file_compare(a, b)
    assert _wait(lambda: bool(view._diff_matrix) and view.status.currentMessage().startswith("비교 완료"))
    view._staged[(1, 1)] = DIR_A2B
    view._staged[(2, 1)] = DIR_B2A
    view._set_save_btn_state()
    view._save_staged("b")
    view.panel_a.path_edit.setText(str(tmp_path / "typing"))      # 저장 중에 경로칸을 고침
    assert _wait(lambda: not view._saving_now() and view._stale_after_save)
    assert view._staged == {}
    view._staged[(2, 1)] = DIR_B2A
    view._save_staged("a")
    assert _vals(a, "S")[2] == ["2", "a2"], "다시 비교하지 못한 낡은 좌표로 저장했다"
