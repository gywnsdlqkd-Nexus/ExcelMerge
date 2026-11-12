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


# ── 2파일 화면과 같은 UI — 툴바 1행 · 저장 버튼 자리 · 찾기 ─────────────────
def _toolbar_row(view, index=0):
    """view 의 index 번째 가로 줄에 놓인 것 — 버튼은 글자(없으면 툴팁 첫 줄), 입력란은 objectName."""
    from PyQt5.QtWidgets import QLineEdit, QPushButton
    out = []

    def walk(lay):
        for i in range(lay.count()):
            it = lay.itemAt(i)
            if it.layout() is not None:
                walk(it.layout())
            elif isinstance(it.widget(), QPushButton):
                b = it.widget()
                out.append(b.text() or b.toolTip().splitlines()[0])
            elif isinstance(it.widget(), QLineEdit):
                out.append(it.widget().objectName())
    walk(view.layout().itemAt(index).layout())
    return out


def test_toolbar_first_row_is_the_same_as_the_two_file_view(win, tmp_path):
    from excelmerge.diff_view import DiffView
    view, *_ = _open(win, tmp_path)
    two = DiffView(win.status)
    try:
        assert _toolbar_row(view) == _toolbar_row(two), "3파일 툴바 1행이 2파일과 달라졌다"
    finally:
        two.shutdown()
        two.deleteLater()
    # 3파일에만 있는 것은 아래 줄에 — 1행의 버튼 자리를 밀지 않는다
    assert _toolbar_row(view, 1) == ["◀ 이전 충돌", "다음 충돌 ▶", "자동 병합"]
    assert win.minimumSizeHint().width() <= 1400, "툴바가 기본 창 폭을 넘는다"


def test_save_button_sits_on_the_middle_panel_like_the_two_file_view(win, tmp_path):
    view, l, b, r, _ = _open(win, tmp_path)
    assert view.save_btn is view.panel_b.save_btn
    assert view.panel_b.save_btn.isVisible()
    assert not view.panel_l.save_btn.isVisible() and not view.panel_r.save_btn.isVisible(), \
        "양옆은 읽기만 하는 파일 — 저장 버튼이 보이면 안 된다"
    assert not view.save_btn.isEnabled()            # 결정이 없으면 꺼져 있다(2파일과 같음)
    view._auto_merge()
    assert view.save_btn.isEnabled()
    view.save_btn.click()
    assert _wait(lambda: not view._saving_now()
                 and view.status.currentMessage().startswith("저장 완료")), \
        view.status.currentMessage()
    assert _vals(b)[1] == ["1", "A", "P"]
    assert "가운데 파일" in view.save_btn.toolTip()


def test_save_button_tooltip_names_the_result_file_in_p4v_mode(win, tmp_path):
    view, l, b, r, res = _open(win, tmp_path, result=True)
    assert res in view.save_btn.toolTip() and "결과 파일" in view.save_btn.toolTip()
    assert res in view.result_label.toolTip()


def _find(view, term):
    view.find_edit.setText(term)
    return view._find_cells(term)


def test_find_searches_what_the_three_panels_show(win, tmp_path):
    view, *_ = _open(win, tmp_path)
    r1, r2, r3 = _row(view, "1"), _row(view, "2"), _row(view, "3")
    assert view.find_edit.isEnabled() and view.find_next_btn.isEnabled()
    assert _find(view, "L3") == [(r3, 1)]             # 왼쪽 패널에만 있는 값
    assert _find(view, "R3") == [(r3, 1)]             # 오른쪽 패널에만 있는 값
    assert _find(view, "c") == [(r3, 1)]              # 가운데(기준) 값
    # '변경점만 보기'로 숨은 행은 찾지 않는다 — 끄면 찾는다
    assert view.diff_only_btn.isChecked() and _find(view, "q") == []
    view.diff_only_btn.setChecked(False)
    assert _find(view, "q") == [(r2, 2)]
    # 검사 제외 열은 뺀다
    view._on_columns_exclude_set([2], True)
    assert _find(view, "q") == []
    view._on_columns_exclude_set([2], False)
    # 결정이 걸린 칸 — 가운데는 고른 값을 보이므로 기준값("c")은 화면에 없다
    _select(view, [(r3, 1)])
    view._stage_selected(SRC_RIGHT)
    assert _find(view, "c") == [], "화면에 없는 기준값을 찾았다"
    assert _find(view, "R3") == [(r3, 1)]
    view._undo_last()
    assert _find(view, "c") == [(r3, 1)], "결정을 되돌렸는데 찾기 결과가 옛것이다"


def test_find_moves_all_three_panels_and_wraps(win, tmp_path):
    view, *_ = _open(win, tmp_path)
    r1, r3 = _row(view, "1"), _row(view, "3")
    view.find_case_btn.setChecked(False)              # 대소문자 구분 — "A" 는 왼쪽 1행뿐
    view.find_edit.setText("A")
    view._goto_find(+1)
    for p in view.panels.values():
        assert p.table._current_cell() == (r1, 1)
    assert '1/1개 일치' in view.status.currentMessage()
    view._goto_find(+1)                               # 하나뿐 — 처음으로 돌아온다
    assert "처음부터 다시" in view.status.currentMessage()
    view.find_edit.setText("없는글자")
    view._goto_find(+1)
    assert "일치 항목이 없습니다" in view.status.currentMessage()
    assert view.panel_b.table._current_cell() == (r1, 1)


def test_ctrl_f_focuses_the_find_box(win, tmp_path):
    # 키 입력을 흉내 내지 않는다 — 창 활성화·수정자 키 상태가 뒤 테스트(헤더 클릭)를 오염시켰다.
    # 단축키가 이 화면 범위로 등록돼 있고, 그 신호가 찾기 입력란을 잡는지만 본다.
    from PyQt5.QtGui import QKeySequence
    from PyQt5.QtWidgets import QShortcut
    view, *_ = _open(win, tmp_path)
    sc = [s for s in view.findChildren(QShortcut)
          if s.parent() is view and s.key() == QKeySequence("Ctrl+F")]
    assert len(sc) == 1 and sc[0].context() == Qt.WidgetWithChildrenShortcut
    view.find_edit.setText("abc")
    calls = []
    view.find_edit.setFocus = lambda *a: calls.append("focus")
    sc[0].activated.emit()
    assert calls == ["focus"] and view.find_edit.selectedText() == "abc"


def test_cli_parses_p4v_merge_and_three_positional():
    import excel_diff_merge as e
    assert e._parse_args(["-b", "B", "-s", "T", "-d", "Y", "-r", "R"]) == ("T", "Y", "B", "R")
    assert e._parse_args(["L", "M", "R"]) == ("L", "R", "M", "")
    assert e._parse_args(["-s", "A", "-d", "B"]) == ("A", "B", "", "")
    assert e._parse_args(["A", "B"]) == ("A", "B", "", "")
