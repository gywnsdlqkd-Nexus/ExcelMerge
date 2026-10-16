# -*- coding: utf-8 -*-
"""틀 고정(키 열·키 행)을 넘나드는 Shift+클릭이 **누른 대로** 선택되는가.

실기 제보: 키 열(A)의 셀을 클릭하고 다른 셀을 Shift+클릭하면 선택 영역이 엉뚱하게
틀어졌다. Qt 는 Shift+클릭의 기준(앵커)을 **뷰마다 따로** 기억하고 그 셀의 화면 위치에서
사각형을 그리는데, 키 열은 본체에서 숨긴 채 오버레이가 그린다. 그래서 둘 다 틀렸다.

  ① A열 클릭은 오버레이가 받아 본체의 앵커가 안 바뀐다 — 예전에 본체에서 누른 셀부터
     선택됐다(실측: 기대 5~9행 × A~C → 9~12행 × C~D).
  ② 앵커가 맞아도 본체에서 숨은 열의 화면 위치는 빈 사각형이다 — 보이는 첫 칸부터
     선택됐다(실측: 1~9행 × B~C).

전부 실제 마우스 이벤트(QTest)로, 화면 없이 돈다.
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from PyQt5.QtCore import QEvent, QPoint, Qt
from PyQt5.QtGui import QMouseEvent
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QApplication


@pytest.fixture
def dv(qapp):
    """키 앵커 A1(키 열 A · 키 행 1), 스크롤이 생길 만큼 긴 표. 틀 고정 ON, 전체 행 표시."""
    from excelmerge.main_window import MainWindow
    win = MainWindow()
    win.resize(1200, 700)
    win.show()
    try:
        v = win.tabs.currentWidget()
        v.panel_a.set_path("a.xlsx"); v.panel_b.set_path("b.xlsx")
        a = [["ID", "V", "W", "U"]] + [[str(i), "a%d" % i, "p", "q"] for i in range(200)]
        b = [row[:] for row in a]
        b[7][1] = "CH"
        v._on_loaded(a, b)
        w = getattr(v, "_diff_worker", None)
        if w is not None:
            w.wait(8000)
        qapp.processEvents()
        v._freeze["a"].refresh()
        v.diff_only_btn.setChecked(False)
        qapp.processEvents()
        assert v._freeze["a"].active, "전제: 틀 고정이 켜져 있어야 한다"
        yield v
    finally:
        win.close(); win.deleteLater(); qapp.processEvents()


def _pos(view, r, c):
    host = getattr(view, "_host", view)
    rect = view.visualRect(host._model.index(r, c))
    assert rect.isValid() and not rect.isEmpty(), f"({r},{c}) 가 이 뷰에 안 보인다"
    return rect.center()


def _click(view, r, c, mod=Qt.NoModifier):
    QTest.mouseClick(view.viewport(), Qt.LeftButton, mod, _pos(view, r, c))
    QApplication.instance().processEvents()


def _move_event(vp, lp, mods):
    """누른 채 움직임 — 실제 이벤트처럼 로컬·창·전역 좌표를 모두 채운다.

    위치 하나만 주는 생성자는 전역 좌표를 **로컬 값으로** 채운다. 선택 엔진은 경계를 넘는
    드래그를 전역 좌표로 판정하므로, 그렇게 만든 이벤트는 엉뚱한 칸을 가리킨다.
    """
    from PyQt5.QtCore import QPointF
    return QMouseEvent(QEvent.MouseMove, QPointF(lp), QPointF(vp.mapTo(vp.window(), lp)),
                       QPointF(vp.mapToGlobal(lp)), Qt.NoButton, Qt.LeftButton, mods)


def _box(table):
    sel = table.selectionModel().selection()
    if not sel:
        return None
    return (min(x.top() for x in sel), min(x.left() for x in sel),
            max(x.bottom() for x in sel), max(x.right() for x in sel))


def _current(table):
    i = table.currentIndex()
    return (i.row(), i.column())


# ── 제보 그대로: 키 열에서 시작 ─────────────────────────────────────────────

def test_from_the_key_column_to_the_body(dv):
    t, fc = dv.panel_a.table, dv._freeze["a"]
    _click(fc.left, 5, 0)
    _click(t, 9, 2, Qt.ShiftModifier)
    assert _box(t) == (5, 0, 9, 2), "보이는 첫 칸부터 선택됐다(숨은 열의 빈 화면 위치)"
    assert _current(t) == (9, 2)


def test_an_earlier_body_click_is_not_the_anchor(dv):
    """예전에 본체에서 누른 셀이 기준이 되면 안 된다 — 키 열 클릭이 새 기준이다."""
    t, fc = dv.panel_a.table, dv._freeze["a"]
    _click(t, 12, 3)
    _click(fc.left, 5, 0)
    _click(t, 9, 2, Qt.ShiftModifier)
    assert _box(t) == (5, 0, 9, 2)


def test_from_the_body_to_the_key_column(dv):
    """반대 방향 — 오버레이도 자기 낡은 기준을 쓰면 안 된다."""
    t, fc = dv.panel_a.table, dv._freeze["a"]
    _click(fc.left, 20, 0)          # 오버레이에 낡은 기준을 남겨 둔다
    _click(t, 9, 2)
    _click(fc.left, 4, 0, Qt.ShiftModifier)
    assert _box(t) == (4, 0, 9, 2)


def test_from_the_key_row_to_the_body(dv):
    """키 행(위쪽 고정)도 같은 구조다."""
    t, fc = dv.panel_a.table, dv._freeze["a"]
    _click(fc.top, 0, 2)
    _click(t, 6, 3, Qt.ShiftModifier)
    assert _box(t) == (0, 2, 6, 3)


# ── 엑셀처럼: 기준은 유지되고 범위만 바뀐다 ────────────────────────────────

def test_repeated_shift_clicks_keep_the_anchor(dv):
    t, fc = dv.panel_a.table, dv._freeze["a"]
    _click(fc.left, 5, 0)
    _click(t, 9, 2, Qt.ShiftModifier)
    _click(t, 7, 3, Qt.ShiftModifier)
    assert _box(t) == (5, 0, 7, 3)
    _click(fc.left, 2, 0, Qt.ShiftModifier)
    assert _box(t) == (2, 0, 5, 0)


def test_a_keyboard_move_resets_the_anchor(dv):
    """클릭 뒤 방향키로 옮겼으면 그 셀이 새 기준이다."""
    t, fc = dv.panel_a.table, dv._freeze["a"]
    _click(t, 5, 1)
    t.setFocus()
    QTest.keyClick(t, Qt.Key_Down)
    QTest.keyClick(t, Qt.Key_Down)
    QApplication.instance().processEvents()
    assert _current(t) == (7, 1)
    _click(t, 9, 3, Qt.ShiftModifier)
    assert _box(t) == (7, 1, 9, 3)


def test_a_jittery_shift_click_keeps_its_selection(dv):
    """실제 클릭엔 누른 채 1~2px 움직임이 섞인다 — 그 이동을 Qt 기본에 넘기면 낡은 누름
    위치에서 사각형을 다시 그려 방금 선택을 덮어쓴다."""
    t, fc = dv.panel_a.table, dv._freeze["a"]
    _click(t, 12, 3)
    _click(fc.left, 5, 0)
    vp, p = t.viewport(), _pos(t, 9, 2)
    QTest.mousePress(vp, Qt.LeftButton, Qt.ShiftModifier, p)
    QApplication.sendEvent(vp, _move_event(vp, p + QPoint(2, 1), Qt.ShiftModifier))
    QTest.mouseRelease(vp, Qt.LeftButton, Qt.ShiftModifier, p + QPoint(2, 1))
    QApplication.instance().processEvents()
    assert _box(t) == (5, 0, 9, 2)


def test_a_plain_click_on_the_last_target_starts_over_in_the_key_column(dv):
    """방금 Shift+클릭으로 끝난 셀을 그냥 클릭하면 거기가 새 기준이다.

    현재 셀이 그대로라 '키보드로 옮겨졌나' 판정으로는 못 알아챈다 — 클릭이 직접 기준을
    남겨야 한다(키 열은 오버레이가 받으므로 오버레이도).
    """
    t, fc = dv.panel_a.table, dv._freeze["a"]
    _click(t, 9, 2)
    _click(fc.left, 4, 0, Qt.ShiftModifier)     # 현재 셀 = (4, 0)
    _click(fc.left, 4, 0)                       # 그 셀을 그냥 클릭 → 새 기준
    _click(t, 6, 3, Qt.ShiftModifier)
    assert _box(t) == (4, 0, 6, 3)


def test_a_plain_click_on_the_last_target_starts_over_in_the_body(dv):
    t, fc = dv.panel_a.table, dv._freeze["a"]
    _click(fc.left, 5, 0)
    _click(t, 9, 2, Qt.ShiftModifier)           # 현재 셀 = (9, 2)
    _click(t, 9, 2)
    _click(t, 11, 3, Qt.ShiftModifier)
    assert _box(t) == (9, 2, 11, 3)


# ── 기존 동작은 그대로 ─────────────────────────────────────────────────────

def test_a_plain_click_still_selects_one_cell(dv):
    t, fc = dv.panel_a.table, dv._freeze["a"]
    _click(fc.left, 5, 0)
    _click(t, 9, 2, Qt.ShiftModifier)
    _click(t, 11, 1)
    assert _box(t) == (11, 1, 11, 1)


def test_the_other_panel_mirrors_the_range(dv, qapp):
    """A 패널의 선택은 B 패널에 미러된다 — 직접 만든 범위도 마찬가지여야 한다."""
    t, fc = dv.panel_a.table, dv._freeze["a"]
    _click(fc.left, 5, 0)
    _click(t, 9, 2, Qt.ShiftModifier)
    qapp.processEvents()
    assert _box(dv.panel_b.table) == (5, 0, 9, 2)
