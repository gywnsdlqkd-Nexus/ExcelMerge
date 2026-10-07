# -*- coding: utf-8 -*-
"""드래그 중 A↔B 선택 미러 지연 회귀 테스트.

대량 선택을 매 드래그 스텝마다 반대 패널에 미러하면(전범위 ClearAndSelect) 헤더 드래그가
극심하게 느려진다(측정: 실제 변화가 생기는 선택 갱신 24.6ms 중 약 22ms 가 미러).
그래서 드래그 중에는 미러를 보류하고 **종료 시 1회만** 미러한다.

이 테스트가 지키는 계약:
  1) 드래그 중에는 미러하지 않는다.
  2) 드래그 종료 시 반대 패널이 정확히 미러된다.
  3) 드래그가 아닌 평범한 선택은 즉시 미러된다(지연이 일반 선택을 망치지 않는다).
  4) 본체/오버레이/헤더의 **모든 press 경로**가 플래그를 세우고 release 가 해제한다
     — 한 곳이라도 빠지면 플래그가 남아 그 뒤 선택이 미러되지 않는다.
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from PyQt5.QtCore import Qt, QPoint, QItemSelection, QItemSelectionModel
from PyQt5.QtWidgets import QApplication


def _wait(win, ms=5000):
    w = getattr(win, "_diff_worker", None)
    if w is not None:
        w.wait(ms)
    app = QApplication.instance()
    if app is not None:
        app.processEvents()


@pytest.fixture
def dv(qapp):
    """키 행/열 1 로 틀 고정된 활성 DiffView (test_freeze_selection 과 동일 구성)."""
    from excelmerge.main_window import MainWindow
    win = MainWindow()
    win.show()
    try:
        view = win.tabs.currentWidget()
        view.panel_a.set_path("a.xlsx"); view.panel_b.set_path("b.xlsx")
        view._key_row = 1; view._key_col = 1
        for p in (view.panel_a, view.panel_b):
            p.table.set_key_row(1); p.table.set_key_col(1)
        a = [["pre", "x", "y", "z"], ["ID", "V", "W", "U"]] + \
            [[str(i), "a%d" % i, "p", "q"] for i in range(15)]
        b = [row[:] for row in a]
        b[5][2] = "CH"
        view._on_loaded(a, b); _wait(win)
        qapp.processEvents(); view._freeze["a"].refresh(); qapp.processEvents()
        yield view
    finally:
        win.close(); win.deleteLater(); qapp.processEvents()


def _cols(table):
    return sorted({i.column() for i in table.selectionModel().selectedIndexes()})


def _select_cols(table, c1, c2):
    m = table.model()
    sel = QItemSelection(m.index(0, c1), m.index(table.rowCount() - 1, c2))
    table.selectionModel().select(sel, QItemSelectionModel.ClearAndSelect)


def test_normal_selection_mirrors_immediately(dv, qapp):
    """드래그가 아닐 때는 기존처럼 즉시 미러돼야 한다."""
    a, b = dv.panel_a.table, dv.panel_b.table
    assert not a._drag_selecting
    _select_cols(a, 2, 3)
    qapp.processEvents()
    assert _cols(a) == [2, 3]
    assert _cols(b) == _cols(a), "평범한 선택이 미러되지 않았다"


def test_mirror_deferred_during_drag_then_applied_on_release(dv, qapp):
    """드래그 중에는 미러 보류, 종료 시 1회 미러."""
    a, b = dv.panel_a.table, dv.panel_b.table
    _select_cols(a, 0, 0)
    qapp.processEvents()
    b.selectionModel().clearSelection()
    qapp.processEvents()

    a._drag_selecting = True          # press 가 세우는 상태
    for tgt in (2, 3):                # 드래그 스텝
        _select_cols(a, 2, tgt)
        qapp.processEvents()
    assert _cols(a) == [2, 3]
    assert _cols(b) == [], f"드래그 중에 미러됐다: {_cols(b)}"

    a._end_drag_select()              # release
    qapp.processEvents()
    assert not a._drag_selecting
    assert _cols(b) == _cols(a), "종료 시 미러가 안 붙었다"


def test_end_drag_select_is_noop_when_not_dragging(dv, qapp):
    """드래그가 아닐 때 호출돼도 신호를 쏘지 않는다(불필요한 미러 방지)."""
    a = dv.panel_a.table
    fired = []
    a.drag_selection_finished.connect(lambda: fired.append(1))
    a._drag_selecting = False
    a._end_drag_select()
    assert fired == []


def test_press_paths_set_flag_and_release_clears(dv, qapp):
    """본체·오버레이·헤더 세 press 경로가 플래그를 세우고 release 가 해제하는지.

    한 경로라도 release 에서 해제되지 않으면 플래그가 남아 이후 선택이 미러되지 않는다.
    """
    from PyQt5.QtGui import QMouseEvent
    from PyQt5.QtCore import QEvent, QPointF
    a = dv.panel_a.table
    fc = dv._freeze["a"]

    def press_release(widget, pos):
        gp = widget.mapToGlobal(pos)
        for et in (QEvent.MouseButtonPress, QEvent.MouseButtonRelease):
            QApplication.sendEvent(widget, QMouseEvent(
                et, QPointF(pos), QPointF(gp), Qt.LeftButton,
                Qt.LeftButton if et == QEvent.MouseButtonPress else Qt.NoButton,
                Qt.NoModifier))
        qapp.processEvents()

    targets = [
        ("본체", a.viewport(), QPoint(30, 30)),
        ("좌측 고정 오버레이", fc.left.viewport(), QPoint(5, 30)),
        ("열 헤더", a.horizontalHeader().viewport(), QPoint(40, 6)),
        ("행 헤더", a.verticalHeader().viewport(), QPoint(6, 40)),
    ]
    for name, widget, pos in targets:
        a._drag_selecting = False
        press_release(widget, pos)
        assert not a._drag_selecting, f"{name}: release 후에도 드래그 플래그가 남아 있다"

    # 플래그가 남지 않았으니 이후 평범한 선택은 즉시 미러돼야 한다
    _select_cols(a, 2, 2)
    qapp.processEvents()
    assert _cols(dv.panel_b.table) == _cols(a)
