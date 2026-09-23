# -*- coding: utf-8 -*-
"""비교 직후 키보드로 바로 스크롤되는가 — 포커스 인계 + 선택 없는 PageUp/PageDown.

비교는 대개 경로칸에서 Enter 로 시작하므로 끝난 뒤에도 포커스가 경로칸에 남았고,
그래서 **셀을 한 번 클릭하기 전에는 PageUp/PageDown 이 표에 닿지도 않았다**.
표에 포커스가 있어도 현재 셀이 없으면 Qt 는 '격자 맨 앞'으로 쳐서 PageDown 에
오히려 화면이 위로 튀었다(실측 스크롤 1 → 0). 둘 다 여기서 고정한다.
"""
import pytest
from PyQt5.QtCore import Qt
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QApplication

N_COLS = 6
N_ROWS = 200      # 스크롤이 생길 만큼 충분히

HEADER = [f"H{c}" for c in range(N_COLS)]
A_DATA = [HEADER] + [[f"a{i}_{c}" for c in range(N_COLS)] for i in range(1, N_ROWS)]
# 3의 배수 행만 값이 같아(=변경 없음) 필터에 숨는다 — 나머지는 보인다.
B_DATA = [HEADER] + [[(f"a{i}_{c}" if i % 3 == 0 else f"B{i}_{c}") for c in range(N_COLS)]
                     for i in range(1, N_ROWS)]


@pytest.fixture
def dv(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    monkeypatch.setenv("APPDATA", str(tmp_path))
    win = MainWindow()
    win.resize(900, 500)
    win.show()
    try:
        view = win.tabs.currentWidget()
        view.panel_a.set_path("a.xlsx")
        view.panel_b.set_path("b.xlsx")
        view.panel_a.path_edit.setFocus()     # 경로 입력 후 Enter 친 직후와 같은 상태
        qapp.processEvents()
        view._on_loaded(A_DATA, B_DATA)
        w = getattr(view, "_diff_worker", None)
        if w is not None:
            w.wait(8000)
        for _ in range(30):
            qapp.processEvents()
        yield view
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def _vbar(dv):
    return dv.panel_a.table.verticalScrollBar()


# ── 비교가 끝나면 키보드가 표로 간다 ─────────────────────────────────────────

def test_grid_takes_focus_after_compare(dv):
    assert QApplication.instance().focusWidget() is dv.panel_a.table


def test_pagedown_works_right_after_compare_without_clicking(dv):
    """보고된 요청 그대로 — 셀을 한 번도 안 고르고 PageDown 만 눌러도 화면이 내려간다."""
    bar = _vbar(dv)
    assert bar.maximum() > 0, "전제: 스크롤이 생길 만큼 행이 있어야 한다"
    before = bar.value()
    QTest.keyClick(QApplication.instance().focusWidget(), Qt.Key_PageDown, Qt.NoModifier)
    QApplication.instance().processEvents()
    assert bar.value() > before


def test_find_box_keeps_focus(dv):
    """비교는 비동기다 — 그 사이 사용자가 찾기 칸에 들어갔으면 뺏지 않는다."""
    dv.find_edit.setEnabled(True)
    dv.find_edit.setFocus()
    QApplication.instance().processEvents()
    dv._focus_grid()
    assert QApplication.instance().focusWidget() is dv.find_edit


def test_path_box_hands_focus_over(dv):
    dv.panel_a.path_edit.setFocus()
    QApplication.instance().processEvents()
    dv._focus_grid()
    assert QApplication.instance().focusWidget() is dv.panel_a.table


# ── 툴바 버튼을 눌러도 키보드는 표에 남는다 ──────────────────────────────────
# 버튼이 포커스를 가져가면 그다음 PageUp/PageDown 이 표가 아니라 버튼으로 간다 —
# '변경점만 보기'를 끄고 스크롤하려던 순간 키가 죽는 게 이 때문이었다.

def test_toolbar_buttons_do_not_take_focus(dv):
    from PyQt5.QtWidgets import QPushButton
    btns = [dv.diff_only_btn, dv.refresh_btn, dv.prev_diff_btn, dv.next_diff_btn,
            dv.find_case_btn, dv.find_word_btn, dv.find_prev_btn, dv.find_next_btn,
            dv.panel_a.save_btn, dv.panel_b.save_btn]
    for b in btns:
        assert isinstance(b, QPushButton)
        assert b.focusPolicy() == Qt.NoFocus, b.toolTip()


def test_pagedown_still_works_after_turning_the_filter_off_by_click(dv):
    """보고된 경로 그대로 — '변경점만 보기'를 클릭으로 끄고 바로 PageDown."""
    from PyQt5.QtCore import QPoint
    btn = dv.diff_only_btn
    tbl = dv.panel_a.table
    tbl.setFocus()
    QApplication.instance().processEvents()
    assert btn.isChecked(), "전제: 비교 직후엔 필터가 켜져 있다"
    QTest.mouseClick(btn, Qt.LeftButton, Qt.NoModifier,
                     QPoint(btn.width() // 2, btn.height() // 2))
    for _ in range(20):
        QApplication.instance().processEvents()
    assert not btn.isChecked(), "전제: 클릭으로 필터가 꺼졌다"
    assert QApplication.instance().focusWidget() is tbl, "포커스가 버튼에 눌러앉았다"
    bar = _vbar(dv)
    before = bar.value()
    QTest.keyClick(QApplication.instance().focusWidget(), Qt.Key_PageDown, Qt.NoModifier)
    QApplication.instance().processEvents()
    assert bar.value() > before
    assert not tbl.selectionModel().hasSelection()


# ── 고른 셀이 없으면 PageUp/PageDown 은 순수 스크롤 ─────────────────────────
# 표가 포커스를 받으면 Qt 가 선택과 무관하게 커서를 첫 칸에 꽂아 둔다. 그 커서를 기준으로
# 페이지를 옮기면 사용자가 고른 적 없는 셀이 선택돼 버린다 — 보고받은 증상이 이것이다.

def test_paging_after_compare_selects_nothing(dv):
    """비교 직후(클릭 한 번 없이) PageUp/PageDown — 화면만 움직이고 선택은 생기지 않는다."""
    tbl = dv.panel_a.table
    sm = tbl.selectionModel()
    bar = _vbar(dv)
    assert not sm.hasSelection(), "전제: 아직 아무것도 고르지 않았다"
    cursor_before = tbl._current_cell()
    seen = [bar.value()]
    for _ in range(2):
        QTest.keyClick(tbl, Qt.Key_PageDown, Qt.NoModifier)
        QApplication.instance().processEvents()
        seen.append(bar.value())
    QTest.keyClick(tbl, Qt.Key_PageUp, Qt.NoModifier)
    QApplication.instance().processEvents()
    seen.append(bar.value())
    assert seen[1] > seen[0] and seen[2] > seen[1] and seen[3] < seen[2], seen
    assert not sm.hasSelection(), "선택이 생기면 안 된다"
    assert tbl._current_cell() == cursor_before, "커서도 움직이면 안 된다"


def test_shift_paging_without_a_pick_also_only_scrolls(dv):
    """Shift 를 눌러도 늘릴 기준이 없으니 스크롤만 한다."""
    tbl = dv.panel_a.table
    bar = _vbar(dv)
    before = bar.value()
    QTest.keyClick(tbl, Qt.Key_PageDown, Qt.ShiftModifier)
    QApplication.instance().processEvents()
    assert bar.value() > before
    assert not tbl.selectionModel().hasSelection()


# ── 현재 셀이 없을 때의 PageUp/PageDown = 순수 스크롤 ────────────────────────

def _clear_cursor(dv):
    tbl = dv.panel_a.table
    tbl.setFocus()
    QApplication.instance().processEvents()
    tbl.selectionModel().clearCurrentIndex()
    tbl.clearSelection()
    QApplication.instance().processEvents()
    assert tbl._current_cell() == (-1, -1)
    return tbl


def test_pagedown_scrolls_down_without_a_current_cell(dv):
    """Qt 기본은 여기서 커서를 첫 칸에 놓으며 화면을 위로 되돌린다 — 그게 버그였다."""
    tbl = _clear_cursor(dv)
    bar = _vbar(dv)
    bar.setValue(10)
    QApplication.instance().processEvents()
    QTest.keyClick(tbl, Qt.Key_PageDown, Qt.NoModifier)
    QApplication.instance().processEvents()
    assert bar.value() > 10
    assert tbl._current_cell() == (-1, -1), "선택도 커서도 만들지 않는다 — 스크롤만"
    assert not tbl.selectionModel().selectedIndexes()


def test_pageup_scrolls_up_without_a_current_cell(dv):
    tbl = _clear_cursor(dv)
    bar = _vbar(dv)
    bar.setValue(30)
    QApplication.instance().processEvents()
    QTest.keyClick(tbl, Qt.Key_PageUp, Qt.NoModifier)
    QApplication.instance().processEvents()
    assert bar.value() < 30
    assert tbl._current_cell() == (-1, -1)


def test_page_scroll_stops_at_the_ends(dv):
    """격자 끝에서는 그냥 멈춘다(값이 범위를 벗어나지 않는다)."""
    tbl = _clear_cursor(dv)
    bar = _vbar(dv)
    bar.setValue(0)
    QApplication.instance().processEvents()
    QTest.keyClick(tbl, Qt.Key_PageUp, Qt.NoModifier)
    QApplication.instance().processEvents()
    assert bar.value() == 0
    bar.setValue(bar.maximum())
    QApplication.instance().processEvents()
    QTest.keyClick(tbl, Qt.Key_PageDown, Qt.NoModifier)
    QApplication.instance().processEvents()
    assert bar.value() == bar.maximum()


def test_b_panel_follows_the_scroll(dv):
    """A 를 굴리면 B 도 같이 간다 — 기존 스크롤 동기화가 그대로 동작해야 한다."""
    tbl = _clear_cursor(dv)
    _vbar(dv).setValue(5)
    QApplication.instance().processEvents()
    QTest.keyClick(tbl, Qt.Key_PageDown, Qt.NoModifier)
    QApplication.instance().processEvents()
    assert dv.panel_b.table.verticalScrollBar().value() == _vbar(dv).value()


# ── 비회귀: 현재 셀이 있으면 예전처럼 커서가 움직인다 ────────────────────────

def test_pagedown_with_a_current_cell_still_moves_the_cursor(dv):
    tbl = dv.panel_a.table
    r = next(r for r in range(len(dv._diff_matrix)) if not tbl.isRowHidden(r))
    tbl._move_current_cell(r, 4)
    QApplication.instance().processEvents()
    QTest.keyClick(tbl, Qt.Key_PageDown, Qt.NoModifier)
    QApplication.instance().processEvents()
    nr, nc = tbl._current_cell()
    assert nr > r and nc == 4
    assert not tbl.isRowHidden(nr), "볼 수 없는 행에 착지하면 안 된다"
