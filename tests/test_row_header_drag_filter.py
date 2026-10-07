# -*- coding: utf-8 -*-
"""'변경점만 보기' ON 에서 행 헤더 드래그 선택 회귀 테스트.

증상: 필터 ON 이면 행 헤더를 드래그해도 선택이 커서를 따라오지 않았다.

원인: 필터 ON 이면 보이는 행 대부분이 EXTRA(빈) 행인데(변경 행 몇 개 + EXTRA 20개),
_drag_row_at 이 목표를 데이터 범위(data_rows-1)로 clamp 해서 커서가 어느 EXTRA 행 위에
있어도 항상 마지막 데이터 행으로 접혔다. 그 결과
  (1) 선택이 커서를 따라가지 않고,
  (2) 드래그하지도 않은 위쪽 행들이 선택됐다.

부수 원인: _near_section_boundary 가 고정 6px 여백을 써서 22px 행의 59% 를 '리사이즈
그립'으로 오판정 → 드래그 선택이 절반 이상 먹지 않았다.
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtWidgets import QApplication


def _wait(win, ms=5000):
    w = getattr(win, "_diff_worker", None)
    if w is not None:
        w.wait(ms)
    app = QApplication.instance()
    if app is not None:
        app.processEvents()


@pytest.fixture
def view(qapp):
    """변경 행이 드물어 EXTRA 행이 보이는 목록을 지배하는 상황을 만든다."""
    from excelmerge.main_window import MainWindow
    win = MainWindow()
    win.show()
    try:
        dv = win.tabs.currentWidget()
        dv.panel_a.set_path("a.xlsx"); dv.panel_b.set_path("b.xlsx")
        dv._key_row = 0; dv._key_col = 0
        for p in (dv.panel_a, dv.panel_b):
            p.table.set_key_row(0); p.table.set_key_col(0)
        a = [["ID", "V", "W"]] + [[str(i), "a%d" % i, "p"] for i in range(30)]
        b = [row[:] for row in a]
        b[3][1] = "CH"          # 변경 행 1개만 → 나머지는 필터 ON 에서 숨김
        dv._on_loaded(a, b); _wait(win)
        qapp.processEvents()
        yield dv
    finally:
        win.close(); win.deleteLater(); qapp.processEvents()


def _rows(table):
    return sorted({i.row() for i in table.selectionModel().selectedIndexes()})


def test_extra_rows_are_visible_when_filter_on(view, qapp):
    """전제 확인 — 필터 ON 이면 EXTRA(빈) 행이 보이는 목록의 다수를 차지한다."""
    v = view.panel_a.table
    view.diff_only_btn.setChecked(True)
    qapp.processEvents()
    vis = [r for r in range(v.rowCount()) if not v.isRowHidden(r)]
    extra = [r for r in vis if r >= v._model.data_rows]
    assert extra, "EXTRA 행이 보이지 않으면 이 회귀 시나리오가 성립하지 않는다"
    assert len(extra) > len(vis) - len(extra), "EXTRA 행이 다수여야 재현 조건"


def test_drag_target_follows_cursor_over_extra_rows(view, qapp):
    """EXTRA 행 위에서도 드래그 목표가 커서 아래 행을 그대로 가리켜야 한다.

    예전에는 data_rows-1 로 clamp 돼 EXTRA 행 어디를 가리켜도 같은 값이 나왔다.
    """
    v = view.panel_a.table
    view.diff_only_btn.setChecked(True)
    qapp.processEvents()
    vh = v.verticalHeader()
    extra = [r for r in range(v.rowCount())
             if not v.isRowHidden(r) and r >= v._model.data_rows]
    assert len(extra) >= 3

    seen = []
    for row in extra[:3]:
        y = vh.sectionViewportPosition(row) + vh.sectionSize(row) // 2
        gp = v.viewport().mapToGlobal(QPoint(5, y))
        seen.append(v._row_under_global(gp))
    assert seen == extra[:3], f"목표가 커서를 따라오지 않는다: {seen} != {extra[:3]}"
    assert len(set(seen)) == len(seen), "EXTRA 행들이 같은 값으로 접혔다(clamp 회귀)"


def test_drag_from_extra_row_does_not_select_rows_above(view, qapp):
    """EXTRA 행에서 아래로 드래그하면 그 범위만 선택된다(위쪽 행이 끼지 않아야)."""
    v = view.panel_a.table
    view.diff_only_btn.setChecked(True)
    qapp.processEvents()
    extra = [r for r in range(v.rowCount())
             if not v.isRowHidden(r) and r >= v._model.data_rows]
    assert len(extra) >= 4
    anchor, target = extra[1], extra[3]

    v._header_anchor_row = anchor
    v._select_row_range(anchor, target)
    qapp.processEvents()
    got = _rows(v)
    assert got, "선택이 비었다"
    assert min(got) >= anchor, f"드래그하지 않은 위쪽 행이 선택됐다: {got[:6]}"
    assert max(got) == target, f"목표 행까지 선택되지 않았다: {got[-3:]}"


def test_row_header_grip_zone_is_not_majority(view, qapp):
    """행 헤더의 대부분이 '리사이즈 그립'으로 판정되면 드래그 선택이 먹지 않는다.

    22px 행에서 고정 6px 여백은 59% 를 삼켰다. Qt 실제 그립 폭 기준으로 좁혀
    과반은 드래그 선택이 가능해야 한다.
    """
    v = view.panel_a.table
    vh = v.verticalHeader()
    qapp.processEvents()
    h = min(vh.viewport().height(), 300)
    dead = total = 0
    for y in range(h):
        if vh.logicalIndexAt(y) < 0:
            continue
        total += 1
        if v._near_section_boundary(vh, "row", QPoint(3, y)):
            dead += 1
    assert total > 0
    ratio = dead / total
    assert ratio < 0.5, f"헤더의 {100*ratio:.0f}% 가 그립으로 판정돼 드래그 선택이 막힌다"


def test_section_boundary_still_detected_at_the_grip(view, qapp):
    """그립 판정을 좁혀도 실제 경계 바로 위는 여전히 리사이즈로 인식돼야 한다."""
    v = view.panel_a.table
    vh = v.verticalHeader()
    qapp.processEvents()
    rows = [r for r in range(v.rowCount()) if not v.isRowHidden(r)]
    r0 = rows[0]
    start, size = vh.sectionViewportPosition(r0), vh.sectionSize(r0)
    # 섹션 끝 경계 직전 1px 는 그립
    assert v._near_section_boundary(vh, "row", QPoint(3, start + size - 1))
    # 섹션 중앙은 그립이 아니어야 한다(드래그 선택 가능)
    assert not v._near_section_boundary(vh, "row", QPoint(3, start + size // 2))
