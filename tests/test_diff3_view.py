# -*- coding: utf-8 -*-
"""3파일 비교 화면 — 실제 화면·워커·파일 경로로(화면 없이) 돈다. 합성 데이터만 쓴다."""
import os
import time

import openpyxl
import pytest
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication

from excelmerge.diff3_engine import SRC_LEFT, SRC_RIGHT
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


def _wait(pred, timeout=20.0):
    app = QApplication.instance()
    end = time.time() + timeout
    while time.time() < end:
        app.processEvents()
        if pred():
            return True
        time.sleep(0.01)
    return bool(pred())


def _vals(path, sheet="Data"):
    clear_values_cache()
    return load_values_any(path, None, sheet)


ANSWER = [None]              # '남은 충돌' 질문 창의 대답(fixture 가 Yes 로 둔다)

BASE = [H, [1, "a", "p"], [2, "b", "q"], [3, "c", "r"]]
LEFT = [H, [1, "A", "p"], [2, "b", "q"], [3, "L3", "r"]]        # 1 왼쪽만, 3 충돌
RIGHT = [H, [1, "a", "P"], [2, "b", "q"], [3, "R3", "r"]]       # 1 오른쪽만, 3 충돌


@pytest.fixture
def win(qapp, monkeypatch):
    from excelmerge import diff3_view as dv
    from excelmerge.main_window import MainWindow
    for name in ("information", "warning", "critical"):
        monkeypatch.setattr(dv.QMessageBox, name, staticmethod(lambda *a, **k: None))
    asked = []

    def question(*a, **k):
        asked.append(a[2] if len(a) > 2 else "")
        return ANSWER[0]
    monkeypatch.setattr(dv.QMessageBox, "question", staticmethod(question))
    ANSWER[0] = dv.QMessageBox.Yes
    w = MainWindow()
    w.asked = asked
    w.show()
    yield w
    w.close()
    w.deleteLater()
    qapp.processEvents()


def _open(win, tmp_path, left=LEFT, base=BASE, right=RIGHT, result=False, extra=None):
    sheets = lambda rows: {"Data": rows, **(extra or {})}
    l = _xlsx(tmp_path / "left.xlsx", sheets(left))
    b = _xlsx(tmp_path / "base.xlsx", sheets(base))
    r = _xlsx(tmp_path / "right.xlsx", sheets(right))
    res = str(tmp_path / "result.xlsx") if result else ""
    view = win.open_merge3(l, b, r, res)
    assert _wait(lambda: bool(view._matrix) and view.status.currentMessage().startswith("비교 완료")), \
        view.status.currentMessage()
    return view, l, b, r, res


def _row(view, key):
    return next(i for i, row in enumerate(view._matrix) if row[0][2] == key or row[0][1] == key)


def _select(view, cells, role="b"):
    from PyQt5.QtCore import QItemSelection, QItemSelectionModel
    t = view.panels[role].table
    sel = QItemSelection()
    for r, c in cells:
        idx = t.model().index(r, c)
        sel.select(idx, idx)
    t.selectionModel().select(sel, QItemSelectionModel.ClearAndSelect)


def _save(view):
    view._save()
    assert _wait(lambda: not view._saving_now() and view.status.currentMessage().startswith(
        ("저장 완료", "오류"))), view.status.currentMessage()
    assert view.status.currentMessage().startswith("저장 완료"), view.status.currentMessage()


def test_opens_replacing_the_empty_start_tab_and_counts(win, tmp_path):
    from excelmerge.diff3_view import Diff3View
    view, *_ = _open(win, tmp_path)
    assert win.tabs.count() == 1 and isinstance(win.tabs.widget(0), Diff3View)
    msg = view.status.currentMessage()
    assert "변경 3칸" in msg and "충돌 1칸(남은 1)" in msg, msg
    assert win.tabs.tabText(0) == "base.xlsx · 3파일"


def test_panel_colors_follow_what_each_file_changed(win, tmp_path):
    from excelmerge.theme import DIFF_COLORS
    view, *_ = _open(win, tmp_path)
    r1, r3 = _row(view, "1"), _row(view, "3")
    bg = lambda role, r, c: view.panels[role].table.model().data(
        view.panels[role].table.model().index(r, c), Qt.BackgroundRole)
    assert bg("l", r1, 1) == DIFF_COLORS["modified"]       # 왼쪽이 바꾼 칸
    assert bg("l", r1, 2) == DIFF_COLORS["same"]           # 오른쪽만 바꾼 칸 — 왼쪽은 기준 그대로
    assert bg("r", r1, 2) == DIFF_COLORS["modified"]
    assert bg("b", r1, 1) == bg("b", r1, 2) == DIFF_COLORS["modified"]
    assert bg("b", r3, 1) == DIFF_COLORS["conflict"]


def test_auto_merge_then_pick_a_conflict_then_save_into_the_middle(win, tmp_path):
    view, l, b, r, _ = _open(win, tmp_path)
    view._auto_merge()
    r3 = _row(view, "3")
    _select(view, [(r3, 1)])
    view._stage_selected(SRC_RIGHT)
    assert view._staged[(r3, 1)] == SRC_RIGHT
    m = view.panel_b.table.model()
    assert m.display_text(r3, 1) == "R3", "가운데가 고른 값을 보여야 한다"
    _save(view)
    assert _vals(b) == [H, ["1", "A", "P"], ["2", "b", "q"], ["3", "R3", "r"]]
    assert view._merged and not view._staged and not view.save_btn.isEnabled()
    # 저장 뒤에도 비교 기준은 원본 — 병합한 칸은 파랑으로 남고, 다시 저장해도 같은 결과
    assert m.cell_kind(r3, 1) == "merged"


def test_saving_twice_after_revoking_a_merge_rebuilds_from_the_original(win, tmp_path):
    view, l, b, r, _ = _open(win, tmp_path)
    view._auto_merge()
    _save(view)
    r1 = _row(view, "1")
    _select(view, [(r1, 2)])
    view._unstage_selected()                  # 오른쪽 값 병합을 푼다
    assert view.save_btn.isEnabled()
    _save(view)
    assert _vals(b) == [H, ["1", "A", "p"], ["2", "b", "q"], ["3", "c", "r"]], \
        "푼 병합이 결과에 남았다"


def test_p4v_mode_writes_the_result_file_and_auto_stages(win, tmp_path):
    view, l, b, r, res = _open(win, tmp_path, result=True)
    before = open(b, "rb").read()
    msg = view.status.currentMessage()
    assert "자동으로 병합 준비함" in msg, msg
    assert len(view._staged) == 2                 # 1행의 V(왼쪽)·W(오른쪽) — 충돌은 빼고
    _save(view)
    assert _vals(res) == [H, ["1", "A", "P"], ["2", "b", "q"], ["3", "c", "r"]]
    assert open(b, "rb").read() == before, "결과 파일을 받았는데 가운데 원본을 고쳤다"


def test_external_change_blocks_the_save(win, tmp_path):
    view, l, b, r, _ = _open(win, tmp_path)
    view._auto_merge()
    time.sleep(0.05)
    _xlsx(b, {"Data": BASE + [[9, "x", "y"]]})     # 바깥에서 가운데가 바뀜
    view._save()
    assert not view._saving_now()
    assert _vals(b)[-1] == ["9", "x", "y"], "바깥 변경을 덮어썼다"


def test_decisions_survive_switching_sheets_and_are_saved_together(win, tmp_path):
    other_b = [["K", "X"], ["k1", "x"]]
    other_l = [["K", "X"], ["k1", "LX"]]
    view, l, b, r, _ = _open(win, tmp_path, extra={"Other": other_b})
    # 오른쪽 파일의 Other 는 기준 그대로, 왼쪽만 고친다
    _xlsx(l, {"Data": LEFT, "Other": other_l})
    view._run_refresh()
    assert _wait(lambda: bool(view._matrix) and view.status.currentMessage().startswith("비교 완료"))
    view._auto_merge()                         # Data 시트
    n_data = len(view._staged)
    idx = view._sheet_names.index("Other")
    view.sheet_tabs.setCurrentIndex(idx)
    assert _wait(lambda: view._current_sheet == "Other" and bool(view._matrix)
                 and view.status.currentMessage().startswith("비교 완료"))
    view._auto_merge()
    assert "Data" in view._sheet_cache and len(view._sheet_cache["Data"]["staged"]) == n_data
    _save(view)
    assert _vals(b, "Other") == [["K", "X"], ["k1", "LX"]]
    assert _vals(b, "Data")[1] == ["1", "A", "P"]
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("Data"))
    assert _wait(lambda: view._current_sheet == "Data" and bool(view._merged))
    assert len(view._merged) == n_data, "되돌아온 시트의 병합 표시가 사라졌다"


def test_p4v_mode_refuses_to_save_with_unvisited_changed_sheets(win, tmp_path):
    other_b = [["K", "X"], ["k1", "x"]]
    view, l, b, r, res = _open(win, tmp_path, result=True, extra={"Other": other_b})
    _xlsx(r, {"Data": RIGHT, "Other": [["K", "X"], ["k1", "RX"]]})
    view._run_refresh()
    assert _wait(lambda: bool(view._matrix) and view._changed_sheets is not None)
    view._save()
    assert not os.path.exists(res), "열어 보지 않은 시트 변경을 버린 결과를 썼다"


def test_key_change_carries_decisions(win, tmp_path):
    view, *_ = _open(win, tmp_path)
    view._auto_merge()
    staged = dict(view._staged)
    view._on_key_col_changed(-1)              # 행 순서 모드 — 여기선 같은 행끼리 짝지어진다
    assert _wait(lambda: view.status.currentMessage().startswith("비교 완료"))
    assert len(view._staged) == len(staged)


def test_deleted_row_taken_whole(win, tmp_path):
    left = [H, [1, "a", "p"], [3, "c", "r"]]            # 2 지움
    right = [H, [1, "a", "p"], [2, "b", "Q"], [3, "c", "r"]]   # 2 고침 → 충돌
    view, l, b, r, _ = _open(win, tmp_path, left=left, right=right)
    r2 = _row(view, "2")
    _select(view, [(r2, 2)], role="l")
    view._stage_selected(SRC_LEFT)
    assert {c for (rr, c) in view._staged if rr == r2} == {0, 1, 2}, "지운 행을 반만 골랐다"
    _save(view)
    assert _vals(b) == [H, ["1", "a", "p"], ["3", "c", "r"]]


def test_noop_pick_is_refused(win, tmp_path):
    view, *_ = _open(win, tmp_path)
    r1 = _row(view, "1")
    _select(view, [(r1, 1)])                  # 왼쪽만 바뀐 칸에 '오른쪽 → 가운데'
    view._stage_selected(SRC_RIGHT)
    assert not view._staged


def test_cli_parses_p4v_merge_and_three_positional():
    import excel_diff_merge as e
    assert e._parse_args(["-b", "B", "-s", "T", "-d", "Y", "-r", "R"]) == ("T", "Y", "B", "R")
    assert e._parse_args(["L", "M", "R"]) == ("L", "R", "M", "")
    assert e._parse_args(["-s", "A", "-d", "B"]) == ("A", "B", "", "")
    assert e._parse_args(["A", "B"]) == ("A", "B", "", "")
