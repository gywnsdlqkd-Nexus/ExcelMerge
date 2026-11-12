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
    assert m.decision(r3, 1) == (SRC_RIGHT, "merged")


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
    # 3파일에만 있던 충돌 이동 · 자동 병합 버튼은 두지 않는다 — 툴바는 한 줄뿐이다
    from PyQt5.QtWidgets import QPushButton
    texts = {b.text() for b in view.findChildren(QPushButton)}
    assert not texts & {"◀ 이전 충돌", "다음 충돌 ▶", "자동 병합", "저장"}, texts
    assert win.minimumSizeHint().width() <= 1400, "툴바가 기본 창 폭을 넘는다"


def test_every_panel_has_its_own_save_button_like_the_two_file_view(win, tmp_path):
    view, l, b, r, _ = _open(win, tmp_path)
    assert view.save_btn is view.panel_b.save_btn
    for p in view.panels.values():
        assert p.save_btn.isVisible() and not p.save_btn.isEnabled()   # 결정이 없으면 꺼져 있다
    view._auto_merge()
    assert view.save_btn.isEnabled()
    view.save_btn.click()
    assert _wait(lambda: not view._saving_now()
                 and view.status.currentMessage().startswith("저장 완료")), \
        view.status.currentMessage()
    assert _vals(b)[1] == ["1", "A", "P"]
    assert "가운데 파일" in view.save_btn.toolTip()


def test_p4v_mode_writes_only_the_result(win, tmp_path):
    """결과 파일을 따로 받은 병합 — 쓰는 곳은 결과 하나. 양옆 저장 버튼·양옆으로 가는 방향이 없다."""
    view, l, b, r, res = _open(win, tmp_path, result=True)
    assert res in view.save_btn.toolTip() and "결과 파일" in view.save_btn.toolTip()
    assert os.path.basename(res) in view.panel_b.title_label.text()
    assert not view.panel_l.save_btn.isVisible() and not view.panel_r.save_btn.isVisible()
    for p in view.panels.values():
        assert [d for _lbl, d in p.table.stage_choices] == ["l>b", "r>b"]
    r3 = _row(view, "3")
    _select(view, [(r3, 1)])
    view._stage_selected("b>l")                      # 메뉴엔 없지만 불러도 받지 않는다
    assert not view._targets["l"][0]


# ── 여섯 방향 병합 — 어느 파일에서 어느 파일로든 ────────────────────────────
def _save_role(view, role):
    view.panels[role].save_btn.click()
    assert _wait(lambda: not view._saving_now() and view.status.currentMessage().startswith(
        ("저장 완료", "오류"))), view.status.currentMessage()
    assert view.status.currentMessage().startswith("저장 완료"), view.status.currentMessage()


def test_menu_offers_six_directions_and_alt_arrows_still_go_to_the_middle(win, tmp_path):
    view, *_ = _open(win, tmp_path)
    for p in view.panels.values():
        assert p.table.stage_choices == (
            ("왼쪽 → 가운데", "l>b"), ("오른쪽 → 가운데", "r>b"),
            ("가운데 → 왼쪽", "b>l"), ("오른쪽 → 왼쪽", "r>l"),
            ("가운데 → 오른쪽", "b>r"), ("왼쪽 → 오른쪽", "l>r"))


def test_perspective_identity_matches_the_engine():
    from excelmerge.diff3_engine import compute_diff3, perspective
    left = [H, [1, "a", "p"], [3, "c", "r"], [5, "n", "n"]]
    right = [H, [1, "a", "P"], [2, "b", "Q"], [3, "R3", "r"]]
    m, rm = compute_diff3([[str(v) for v in r] for r in left],
                          [[str(v) for v in r] for r in BASE],
                          [[str(v) for v in r] for r in right], 0, 0)
    pm, prm, pcm = perspective(m, rm, None, (0, 1, 2))
    assert pm == m and [tuple(x) for x in prm] == [tuple(x) for x in rm] and pcm is None
    # 왼쪽을 가운데에 둔 모양 — 값 자리가 (가운데, 왼쪽, 오른쪽)
    lm, lrm, _ = perspective(m, rm, None, (1, 0, 2))
    for row, lrow, meta, lmeta in zip(m, lm, rm, lrm):
        assert tuple(lmeta) == (meta[1], meta[0], meta[2])
        for cell, lcell in zip(row, lrow):
            assert lcell[1:] == (cell[2], cell[1], cell[3])


def test_middle_to_left_and_right_to_left_save_into_the_left_file(win, tmp_path):
    view, l, b, r, _ = _open(win, tmp_path)
    r1, r3 = _row(view, "1"), _row(view, "3")
    before_b, before_r = open(b, "rb").read(), open(r, "rb").read()
    _select(view, [(r1, 1)])
    view._stage_selected("b>l")                      # 왼쪽이 고친 A 를 가운데 값 a 로
    _select(view, [(r3, 1)])
    view._stage_selected("r>l")                      # 충돌 칸에 오른쪽 값
    lm = view.panel_l.table.model()
    assert lm.display_text(r1, 1) == "a" and lm.display_text(r3, 1) == "R3", \
        "받는 패널이 가져올 값을 보여야 한다"
    assert view.panel_l.save_btn.isEnabled() and not view.panel_b.save_btn.isEnabled()
    assert not view._staged, "가운데로 가는 결정이 아니다"
    _save_role(view, "l")
    assert _vals(l) == [H, ["1", "a", "p"], ["2", "b", "q"], ["3", "R3", "r"]]
    assert open(b, "rb").read() == before_b and open(r, "rb").read() == before_r, \
        "왼쪽에 저장했는데 다른 파일이 바뀌었다"
    assert not view.panel_l.save_btn.isEnabled()
    assert lm.decision(r1, 1) is not None and lm.decision(r1, 1)[1] == "merged"


def test_one_value_can_go_to_two_files_and_each_saves_on_its_own(win, tmp_path):
    view, l, b, r, _ = _open(win, tmp_path)
    r1 = _row(view, "1")
    _select(view, [(r1, 1)])
    view._stage_selected("l>b")                      # 왼쪽 A → 가운데
    _select(view, [(r1, 1)])
    view._stage_selected("l>r")                      # 같은 값을 오른쪽에도
    assert view.panel_b.save_btn.isEnabled() and view.panel_r.save_btn.isEnabled()
    view._save()                                     # Ctrl+S — 대상이 둘이라 아무 쪽도 쓰지 않는다
    assert not view._saving_now() and "저장할 패널" in view.status.currentMessage()
    _save_role(view, "r")
    assert _vals(r)[1] == ["1", "A", "P"] and _vals(b)[1] == ["1", "a", "p"]
    assert view.panel_b.save_btn.isEnabled(), "오른쪽만 저장했는데 가운데 저장 대기가 꺼졌다"
    _save_role(view, "b")
    assert _vals(b)[1] == ["1", "A", "p"]


def test_second_save_into_a_side_file_rebuilds_from_its_original(win, tmp_path):
    """양옆도 가운데처럼 원본 사본에서 새로 만든다 — 저장 뒤에 병합을 풀고 다시 저장하면 빠진다."""
    view, l, b, r, _ = _open(win, tmp_path)
    r1, r3 = _row(view, "1"), _row(view, "3")
    _select(view, [(r1, 2), (r3, 1)])
    view._stage_selected("l>r")                      # 오른쪽 P → p, R3 → L3
    _save_role(view, "r")
    assert _vals(r) == [H, ["1", "a", "p"], ["2", "b", "q"], ["3", "L3", "r"]]
    _select(view, [(r1, 2)])
    view._unstage_selected()                         # P → p 를 푼다
    assert view.panel_r.save_btn.isEnabled()
    _save_role(view, "r")
    assert _vals(r) == [H, ["1", "a", "P"], ["2", "b", "q"], ["3", "L3", "r"]], \
        "푼 병합이 오른쪽 파일에 남았다"


def test_new_row_goes_to_a_side_file_whole_and_deleted_row_is_removed(win, tmp_path):
    base = [H, [1, "a", "p"], [2, "b", "q"], [3, "c", "r"], [4, "d", "s"]]   # 가운데에만 4
    left = [H, [1, "a", "p"], [3, "c", "r"]]                                 # 왼쪽엔 2·4 없음
    right = [H, [1, "a", "p"], [2, "b", "q"], [3, "c", "r"]]
    view, l, b, r, _ = _open(win, tmp_path, left=left, base=base, right=right)
    r2, r4 = _row(view, "2"), _row(view, "4")
    _select(view, [(r4, 1)])
    view._stage_selected("b>r")                      # 가운데의 새 행 4 를 오른쪽에 — 칸 하나만 골라도 행째
    assert {c for (rr, c) in view._targets["r"][0] if rr == r4} == {0, 1, 2}
    _select(view, [(r2, 1)])
    view._stage_selected("l>r")                      # 왼쪽이 지운 2 를 오른쪽에서도 지운다 — 행째
    assert {c for (rr, c) in view._targets["r"][0] if rr == r2} == {0, 1, 2}
    _save_role(view, "r")
    assert _vals(r) == [H, ["1", "a", "p"], ["3", "c", "r"], ["4", "d", "s"]]


def test_row_missing_in_the_target_comes_whole_from_the_picked_file(win, tmp_path):
    """받는 파일에 없는 행 — 다른 두 파일의 값이 서로 달라도 고른 쪽에서 행째로 온다.
    충돌 칸을 남겨 두면 그 칸이 빈 채로 들어갔다(실측: 20열 행이 3칸만 들어감)."""
    base = [H, [1, "a", "p"], [4, "d", "s"]]
    left = [H, [1, "a", "p"]]                         # 왼쪽엔 4 없음
    right = [H, [1, "a", "p"], [4, "D", "S"]]         # 오른쪽의 4 는 값이 다르다
    view, l, b, r, _ = _open(win, tmp_path, left=left, base=base, right=right)
    r4 = _row(view, "4")
    _select(view, [(r4, 0)])
    view._stage_selected("b>l")                      # 키 칸 하나만 고른다
    assert {c for (rr, c) in view._targets["l"][0] if rr == r4} == {0, 1, 2}
    _select(view, [(r4, 2)])
    view._stage_selected("r>l")                      # 한 칸만 오른쪽 값으로 바꾼다
    _save_role(view, "l")
    assert _vals(l) == [H, ["1", "a", "p"], ["4", "d", "S"]]


def test_a_cell_is_either_a_receiver_or_a_giver_not_both(win, tmp_path):
    view, *_ = _open(win, tmp_path)
    r3 = _row(view, "3")
    _select(view, [(r3, 1)])
    view._stage_selected("l>b")                      # 가운데가 왼쪽 값을 받는다
    _select(view, [(r3, 1)])
    view._stage_selected("b>r")                      # 그 가운데 칸을 오른쪽으로 — 화면 값과 쓰는 값이 갈린다
    assert not view._targets["r"][0], "받기로 한 칸을 다시 출처로 썼다"
    _select(view, [(r3, 1)])
    view._stage_selected("r>l")                      # 왼쪽은 가운데의 출처 — 받는 쪽이 될 수 없다
    assert not view._targets["l"][0], "출처로 쓰인 칸이 다른 값을 받았다"


def test_unstage_and_undo_cover_every_target(win, tmp_path):
    view, *_ = _open(win, tmp_path)
    r1 = _row(view, "1")
    _select(view, [(r1, 1)])
    view._stage_selected("l>b")
    _select(view, [(r1, 1)])
    view._stage_selected("l>r")
    _select(view, [(r1, 1)])
    view._unstage_selected()                         # 한 번에 두 대상 모두
    assert not view._staged and not view._targets["r"][0]
    view._undo_last()
    assert view._staged == {(r1, 1): SRC_LEFT} and len(view._targets["r"][0]) == 1
    view._undo_last()
    assert view._staged and not view._targets["r"][0]


def test_side_decisions_survive_sheet_switch_and_key_change(win, tmp_path):
    other = [["K", "X"], ["k1", "x"]]
    view, l, b, r, _ = _open(win, tmp_path, extra={"Other": other})
    r1 = _row(view, "1")
    _select(view, [(r1, 1)])
    view._stage_selected("b>l")
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("Other"))
    assert _wait(lambda: view._current_sheet == "Other" and bool(view._matrix)
                 and view.status.currentMessage().startswith("비교 완료"))
    assert not view._targets["l"][0] and view.panel_l.save_btn.isEnabled()
    _save_role(view, "l")                            # 다른 시트에 있는 결정도 함께 쓴다
    assert _vals(l)[1] == ["1", "a", "p"]
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("Data"))
    assert _wait(lambda: view._current_sheet == "Data" and bool(view._targets["l"][1]))
    view._on_key_col_changed(-1)                     # 키를 바꿔 다시 비교해도 결정이 따라온다
    assert _wait(lambda: view.status.currentMessage().startswith("비교 완료"))
    assert len(view._targets["l"][1]) == 1


def test_side_file_that_cannot_be_written_is_refused(win, tmp_path):
    view, l, b, r, _ = _open(win, tmp_path)
    csv = tmp_path / "left.csv"
    csv.write_text("ID,V,W\n1,A,p\n2,b,q\n3,L3,r\n", encoding="utf-8")
    view.panel_l.set_path(str(csv))
    view._start_session()
    assert _wait(lambda: bool(view._matrix) and view.status.currentMessage().startswith("비교 완료"))
    r1 = _row(view, "1")
    _select(view, [(r1, 1)])
    view._stage_selected("b>l")
    assert not view._targets["l"][0], "저장할 수 없는 형식의 파일에 병합 준비가 걸렸다"


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
