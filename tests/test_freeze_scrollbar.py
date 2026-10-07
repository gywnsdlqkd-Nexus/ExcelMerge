# -*- coding: utf-8 -*-
"""틀 고정 상태에서 스크롤바가 사라지던 버그 회귀 방지.

원인: ExcelTableView.updateGeometries() 가 super() **뒤에** setViewportMargins 로
고정 밴드 자리를 예약했다. QTableView.updateGeometries() 는 ① 헤더 sizeHint 로
여백을 잡고 ② 그 시점 뷰포트 크기로 스크롤바 range 를 계산하므로, 밴드를 나중에
더하면 range 가 '밴드를 안 뺀 더 넓은 뷰포트' 기준이 된다. 그래서

    0 < (콘텐츠 크기 - 실제 뷰포트 크기) <= 밴드 크기

인 구간에서 range 가 0 이 되어 **스크롤바가 아예 안 뜨고** 마지막 열/행을 볼 수
없었다. 밴드를 헤더 sizeHint 에 미리 포함(_BandHeaderView)해 Qt 가 처음부터 올바른
여백·range 를 계산하게 고쳤다.

두 테스트 모두 '넘치는 양 <= 밴드 크기' 경계를 직접 만들어 검사한다 — 수정 전에는
maximum() == 0 이라 정확히 이 회귀를 잡는다.
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from PyQt5.QtWidgets import QApplication


def _wait_diff(win, timeout_ms=5000):
    w = getattr(win, "_diff_worker", None)
    if w is not None:
        w.wait(timeout_ms)
    app = QApplication.instance()
    if app is not None:
        app.processEvents()


@pytest.fixture
def frozen_view(qapp):
    """key_row=1, key_col=1 로 채워진 활성 DiffView (틀 고정 ON)."""
    from excelmerge.main_window import MainWindow
    win = MainWindow()
    win.resize(1400, 800)
    win.show()
    try:
        dv = win.tabs.currentWidget()
        dv.panel_a.set_path("a.xlsx"); dv.panel_b.set_path("b.xlsx")
        dv._key_row = 1; dv._key_col = 1
        for p in (dv.panel_a, dv.panel_b):
            p.table.set_key_row(1); p.table.set_key_col(1)
        a = [["pre", "x", "y", "z"], ["ID", "V", "W", "U"]] + \
            [[str(i), "a%d" % i, "p", "q"] for i in range(30)]
        b = [row[:] for row in a]
        b[5][2] = "CH"
        dv._on_loaded(a, b); _wait_diff(win)
        qapp.processEvents(); dv._freeze["a"].refresh(); qapp.processEvents()
        assert dv._freeze["a"].active
        dv.diff_only_btn.setChecked(False)   # 전체 행 표시
        qapp.processEvents()
        yield dv
    finally:
        win.close(); win.deleteLater(); qapp.processEvents()


def _visible_cols(t):
    return [c for c in range(t.model().columnCount()) if not t.isColumnHidden(c)]


def _visible_rows(t):
    return [r for r in range(t.model().rowCount()) if not t.isRowHidden(r)]


def test_hscrollbar_reaches_last_column(frozen_view, qapp):
    """가로: 넘치는 폭이 고정 밴드 폭 이하일 때도 스크롤바가 뜨고, 끝까지 밀면
    마지막 열이 뷰포트 안에 완전히 들어와야 한다."""
    dv = frozen_view
    t = dv.panel_a.table
    fc = dv._freeze["a"]
    cols = _visible_cols(t)
    assert cols and fc._fw > 0, "테스트 전제(스크롤 열 + 고정 열) 불충족"

    # 마지막 스크롤 열을 줄여 '콘텐츠 폭 = 뷰포트 폭 + 1px' 로 만든다 —
    # 넘치는 양(1px)이 고정 밴드 폭보다 훨씬 작은 최악의 경계 조건.
    over = sum(t.columnWidth(c) for c in cols) - t.viewport().width()
    t.setColumnWidth(cols[-1], t.columnWidth(cols[-1]) - over + 1)
    qapp.processEvents()

    vp_w = t.viewport().width()
    content = sum(t.columnWidth(c) for c in cols)
    assert content > vp_w, "전제 실패: 콘텐츠가 뷰포트를 안 넘침"
    bar = t.horizontalScrollBar()
    assert bar.maximum() > 0, (
        f"콘텐츠 {content}px > 뷰포트 {vp_w}px 인데 가로 스크롤바 range 가 0 "
        f"— 오른쪽 열을 볼 수 없다")
    bar.setValue(bar.maximum())
    qapp.processEvents()
    right = t.visualRect(t.model().index(_visible_rows(t)[0], cols[-1])).right()
    assert 0 < right < vp_w, (
        f"끝까지 스크롤해도 마지막 열 우측이 뷰포트({vp_w}px) 밖({right}px) — 잘려 보인다")


def test_vscrollbar_reaches_last_row(frozen_view, qapp):
    """세로: 가로의 대칭 — 끝까지 밀면 마지막 행이 뷰포트 안에 완전히 들어와야 한다.
    (고정 행이 본체에서 숨겨져 range 를 부풀리므로 range>0 만으론 회귀를 못 잡는다.)"""
    dv = frozen_view
    t = dv.panel_a.table
    fc = dv._freeze["a"]
    rows = _visible_rows(t)
    assert rows and fc._fh > 0, "테스트 전제(스크롤 행 + 고정 행) 불충족"

    # 가로와 동일하게 '콘텐츠 높이 = 뷰포트 높이 + 1px' 경계를 만든다. 행은 한 개를
    # 줄여서는 부족하므로(높이가 음수가 된다) 전부 낮춘 뒤 마지막 행만 키운다.
    # 기본 minimumSectionSize(폰트 기준 ~21px)가 축소를 막으므로 함께 낮춘다.
    t.verticalHeader().setMinimumSectionSize(2)
    for r in rows:
        t.setRowHeight(r, 8)
    qapp.processEvents()
    over = sum(t.rowHeight(r) for r in rows) - t.viewport().height()
    t.setRowHeight(rows[-1], t.rowHeight(rows[-1]) - over + 1)
    qapp.processEvents()

    vp_h = t.viewport().height()
    content = sum(t.rowHeight(r) for r in rows)
    assert content > vp_h, "전제 실패: 콘텐츠가 뷰포트를 안 넘침"
    bar = t.verticalScrollBar()
    assert bar.maximum() > 0, (
        f"콘텐츠 {content}px > 뷰포트 {vp_h}px 인데 세로 스크롤바 range 가 0")
    bar.setValue(bar.maximum())
    qapp.processEvents()
    bottom = t.visualRect(t.model().index(rows[-1], _visible_cols(t)[0])).bottom()
    assert 0 < bottom < vp_h, (
        f"끝까지 스크롤해도 마지막 행 하단이 뷰포트({vp_h}px) 밖({bottom}px) — 잘려 보인다")


def test_frozen_band_stays_aligned_with_body(frozen_view, qapp):
    """밴드 예약 방식을 바꿔도 본체 뷰포트/헤더/고정 밴드 정렬은 그대로여야 한다."""
    dv = frozen_view
    t = dv.panel_a.table
    fc = dv._freeze["a"]
    x = t.frameWidth() + t.verticalHeader().width() + fc._fw
    y = t.frameWidth() + t.horizontalHeader().height() + fc._fh
    assert t.viewport().geometry().x() == x, "본체 뷰포트가 고정 열 밴드 옆에 없음"
    assert t.viewport().geometry().y() == y, "본체 뷰포트가 고정 행 밴드 아래에 없음"
    assert t.horizontalHeader().geometry().x() == x, "본체 가로 헤더 정렬 어긋남"
    assert t.verticalHeader().geometry().y() == y, "본체 세로 헤더 정렬 어긋남"
    assert fc.top.geometry().x() == x, "고정 상단 밴드 정렬 어긋남"
    assert fc.left.geometry().y() == y, "고정 좌측 밴드 정렬 어긋남"
