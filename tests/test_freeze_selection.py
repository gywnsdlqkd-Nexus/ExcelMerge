# -*- coding: utf-8 -*-
"""틀 고정 관련 두 버그 회귀 방지:

1) 데이터 열/행 리사이즈 시 반대 패널의 고정 밴드가 본체와 어긋남
   — 크로스패널 크기 동기(apply_column_width/apply_row_height)가 _applying_sizes 가드로
     FreezeController 리사이즈 핸들러를 억제해, 고정 top/left 밴드가 새 크기를 반영 못 함.
   → apply_*가 명시적으로 밴드를 재동기(_resync_freeze_sizes)해야 한다.

2) 헤더/Shift+방향키 다중 선택 시 키 열/행이 함께 선택되지 않음
   — 키 열/행은 본체에서 '틀 고정'으로 숨겨져 러버밴드/범위 선택에 안 잡힌다.
   → 다중 셀 선택 시 공유 선택 모델에 키 열(0..key_col)·키 행(0..key_row)을 보충해야 한다.
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from PyQt5.QtCore import Qt, QEvent, QItemSelection, QItemSelectionModel, QPoint
from PyQt5.QtGui import QMouseEvent
from PyQt5.QtWidgets import QApplication, QMenu


def _wait_diff(win, timeout_ms=5000):
    w = getattr(win, "_diff_worker", None)
    if w is not None:
        w.wait(timeout_ms)
    app = QApplication.instance()
    if app is not None:
        app.processEvents()


@pytest.fixture
def frozen_view(qapp):
    """key_row=1, key_col=1 (고정 행 0~1, 고정 열 0~1) 로 채워진 활성 DiffView."""
    from excelmerge.main_window import MainWindow
    win = MainWindow()
    win.show()
    try:
        dv = win.tabs.currentWidget()
        dv.panel_a.set_path("a.xlsx"); dv.panel_b.set_path("b.xlsx")
        dv._key_row = 1; dv._key_col = 1
        for p in (dv.panel_a, dv.panel_b):
            p.table.set_key_row(1); p.table.set_key_col(1)
        # 4 데이터 열(0,1 = 키 / 2,3 = 데이터), 넉넉한 데이터 행.
        a = [["pre", "x", "y", "z"], ["ID", "V", "W", "U"]] + \
            [[str(i), "a%d" % i, "p", "q"] for i in range(15)]
        b = [row[:] for row in a]
        b[5][2] = "CH"
        dv._on_loaded(a, b); _wait_diff(win)
        qapp.processEvents(); dv._freeze["a"].refresh(); qapp.processEvents()
        assert dv._freeze["a"].active and dv._freeze["b"].active
        yield dv
    finally:
        win.close(); win.deleteLater(); qapp.processEvents()


# ── 버그 1: 리사이즈 시 고정 밴드 재정렬 ──────────────────────────────────────
def test_apply_column_width_resyncs_freeze_band(frozen_view):
    """반대 패널로 열 너비가 동기될 때 그 패널의 고정 상단 밴드(top)도 새 폭을 반영해야 한다."""
    dv = frozen_view
    tb = dv.panel_b.table
    fc = dv._freeze["b"]
    W = tb.columnWidth(2) + 47          # 기존과 다른 폭
    tb.apply_column_width(2, W)
    QApplication.instance().processEvents()
    assert tb.columnWidth(2) == W                       # 본체 반영
    assert fc.top.columnWidth(2) == W, "고정 상단 밴드가 새 열 폭을 반영 못 함(어긋남)"


def test_apply_row_height_resyncs_freeze_band(frozen_view):
    """반대 패널로 행 높이가 동기될 때 그 패널의 고정 좌측 밴드(left)도 새 높이를 반영해야 한다."""
    dv = frozen_view
    tb = dv.panel_b.table
    fc = dv._freeze["b"]
    H = tb.rowHeight(5) + 13
    tb.apply_row_height(5, H)
    QApplication.instance().processEvents()
    assert tb.rowHeight(5) == H
    assert fc.left.rowHeight(5) == H, "고정 좌측 밴드가 새 행 높이를 반영 못 함(어긋남)"


# ── 버그 1b: 고정(키) 열/행을 corner 오버레이에서 리사이즈 ────────────────────
def _aligned(fc, t):
    x = t.frameWidth() + t.verticalHeader().width() + fc._fw
    return (t.viewport().geometry().x() == t.horizontalHeader().geometry().x()
            == fc.top.geometry().x() == x)


def test_corner_col_resize_updates_fw_mirrors_and_aligns(frozen_view):
    """키 열은 본체에서 숨겨져 corner 헤더에서만 조절 가능 — 드래그 시 _fw·본체 여백·
    반대 패널이 함께 갱신되고 고정 밴드가 본체와 정렬을 유지해야 한다."""
    dv = frozen_view
    fa, fb = dv._freeze["a"], dv._freeze["b"]
    ta, tb = dv.panel_a.table, dv.panel_b.table
    fw0 = fa._fw
    new = fa.corner.columnWidth(0) + 70
    fa.corner.horizontalHeader().resizeSection(0, new)
    QApplication.instance().processEvents()
    assert fa._fw == fw0 + 70, f"_fw 미갱신: {fa._fw} (기대 {fw0+70})"
    assert ta._user_col_widths.get(0) == new, "호스트에 키 열 폭 기록 안 됨"
    assert fb._fw == fa._fw, f"반대 패널 미러 안 됨: A={fa._fw} B={fb._fw}"
    assert _aligned(fa, ta), "리사이즈 후 A 고정 밴드가 본체와 어긋남"
    assert _aligned(fb, tb), "리사이즈 후 B(미러) 고정 밴드가 본체와 어긋남"


def test_corner_row_resize_updates_fh_and_mirrors(frozen_view):
    """키 행 높이도 corner 수직 헤더에서만 조절 가능 — _fh·반대 패널 동기."""
    dv = frozen_view
    fa, fb = dv._freeze["a"], dv._freeze["b"]
    ta = dv.panel_a.table
    fh0 = fa._fh
    new = fa.corner.rowHeight(0) + 25
    fa.corner.verticalHeader().resizeSection(0, new)
    QApplication.instance().processEvents()
    assert fa._fh == fh0 + 25, f"_fh 미갱신: {fa._fh} (기대 {fh0+25})"
    assert ta._user_row_heights.get(0) == new, "호스트에 키 행 높이 기록 안 됨"
    assert fb._fh == fa._fh, f"반대 패널 미러 안 됨: A={fa._fh} B={fb._fh}"


def test_data_column_resize_still_aligns(frozen_view):
    """회귀 방지: 데이터 열(본체)을 리사이즈해도 양 패널 고정 밴드 정렬 유지."""
    dv = frozen_view
    fa, fb = dv._freeze["a"], dv._freeze["b"]
    ta, tb = dv.panel_a.table, dv.panel_b.table
    ta.horizontalHeader().resizeSection(2, ta.columnWidth(2) + 60)
    QApplication.instance().processEvents()
    assert tb.columnWidth(2) == ta.columnWidth(2), "데이터 열 폭 미러 안 됨"
    assert _aligned(fa, ta) and _aligned(fb, tb), "데이터 열 리사이즈 후 밴드 어긋남"


# ── 엑셀식 선택: '터치한 셀만' 선택 — 키 열/행 자동 보충 없음 (key_col=1, key_row=1) ──
def test_cell_block_no_key_supplement(frozen_view):
    """임의의 셀 블록(행 5~7 × 열 2~3)을 드래그하면 딱 그 블록만 선택된다 —
    키 열(0,1)·키 행(0,1)은 손대지 않는 게 엑셀식(터치한 셀만)."""
    dv = frozen_view
    dv.diff_only_btn.setChecked(False)   # 전체 행 표시 → 5,6,7 모두 보임
    QApplication.instance().processEvents()
    host = dv.panel_a.table
    sm = host.selectionModel(); m = host.model()
    sm.select(QItemSelection(m.index(5, 2), m.index(7, 3)),
              QItemSelectionModel.ClearAndSelect)
    QApplication.instance().processEvents()
    assert host.get_selected_cells() == {(r, c) for r in (5, 6, 7) for c in (2, 3)}, \
        host.get_selected_cells()


def test_partial_row_strip_no_key_supplement(frozen_view):
    """행 일부(가로 스트립 5행 × 열 2~3)도 딱 그 스트립만 선택 — 키 열/행 보충 없음."""
    dv = frozen_view
    dv.diff_only_btn.setChecked(False)
    QApplication.instance().processEvents()
    host = dv.panel_a.table
    sm = host.selectionModel(); m = host.model()
    sm.select(QItemSelection(m.index(5, 2), m.index(5, 3)),
              QItemSelectionModel.ClearAndSelect)
    QApplication.instance().processEvents()
    assert host.get_selected_cells() == {(5, 2), (5, 3)}, host.get_selected_cells()


def test_header_column_selection_no_cross_key_col(frozen_view):
    """데이터 열 전체 선택(_select_col(2))은 그 열만 선택한다 — 다른 키 열(0,1)은 선택 안 됨.
    (열 2 자체의 키 행 셀(행0,1 × 열2)은 열의 일부라 당연히 선택된다.)"""
    dv = frozen_view
    host = dv.panel_a.table
    host._select_col(2)
    QApplication.instance().processEvents()
    sm = host.selectionModel()
    assert not sm.isSelected(host.model().index(5, 0)), "데이터 열 선택인데 키 열0 선택됨"
    assert not sm.isSelected(host.model().index(5, 1)), "데이터 열 선택인데 키 열1 선택됨"
    assert host._full_columns_selected() == [2], host._full_columns_selected()


def test_header_row_selection_no_cross_key_row(frozen_view):
    """데이터 행 전체 선택(_select_rows)은 그 행만 선택 — 다른 키 행(0,1)은 선택 안 됨.
    (행 자체의 키 열 셀은 행의 일부라 당연히 선택된다.)"""
    dv = frozen_view
    dv.diff_only_btn.setChecked(False)   # 전체 행 표시 → 5,6 모두 보임
    QApplication.instance().processEvents()
    host = dv.panel_a.table
    host._select_rows([5, 6])
    QApplication.instance().processEvents()
    sm = host.selectionModel()
    assert not sm.isSelected(host.model().index(0, 3)), "데이터 행 선택인데 키 행0 선택됨"
    assert not sm.isSelected(host.model().index(1, 3)), "데이터 행 선택인데 키 행1 선택됨"
    assert host._full_rows_selected() == [5, 6], host._full_rows_selected()


def test_shift_band_prunes_hidden_filtered_rows(frozen_view):
    """회귀(사용자 보고): '변경 행만 보기' ON에서 보이는 변경 행 사이를 Shift로 잡으면,
    사이에 숨겨진(볼 수 없는) 미변경 행은 선택에서 제외돼야 한다 — 보이는 행만 선택.
    키 행 보충은 하지 않으므로 키 행(0,1)은 선택되지 않는다(엑셀식)."""
    dv = frozen_view
    host = dv.panel_a.table
    sm = host.selectionModel(); m = host.model()
    # 픽스처: 변경 행은 display 5 하나만 보임. 숨겨진 데이터 행(예: 6)이 밴드에 끼면 제외.
    assert host.isRowHidden(6) and not host.isRowHidden(5)
    host._select_row_range(5, 8)   # 5(보임)~8 전폭 밴드 — 6,7,8 은 숨김
    QApplication.instance().processEvents()
    sel_rows = sorted({r for r, c in host.get_selected_cells()})
    assert 5 in sel_rows                         # 보이는 변경 행은 유지
    assert 0 not in sel_rows and 1 not in sel_rows, "키 행이 보충됨(엑셀식 아님)"
    assert not any(host.isRowHidden(r) for r in sel_rows), \
        f"숨겨진 행이 선택에 남음: {sel_rows}"
    assert host._full_rows_selected() == [5], host._full_rows_selected()


def test_cell_block_keeps_hidden_rows(frozen_view):
    """전폭이 아닌 셀 블록은 프룬 대상이 아니다 — 드래그한 직사각형 그대로(숨은 행 포함) 유지.
    키 열/행 보충도 없다 → 딱 블록만."""
    dv = frozen_view
    host = dv.panel_a.table
    sm = host.selectionModel(); m = host.model()
    assert host.isRowHidden(6)   # 픽스처: 6~8 은 숨김(미변경)
    sm.select(QItemSelection(m.index(5, 2), m.index(8, 3)),
              QItemSelectionModel.ClearAndSelect)
    QApplication.instance().processEvents()
    assert host.get_selected_cells() == {(r, c) for r in (5, 6, 7, 8) for c in (2, 3)}, \
        host.get_selected_cells()


def test_diff_only_full_row_selection_no_key_supplement(frozen_view):
    """'변경 행만 보기' ON에서 변경 행 전체를 선택해도 키 행(0,1)은 보충되지 않는다(엑셀식).
    선택된 행(5)은 전폭이라 그 행의 키 열 셀은 행의 일부로 포함된다."""
    dv = frozen_view
    dv.diff_only_btn.setChecked(True)   # 변경 행만 보기 ON
    QApplication.instance().processEvents()
    host = dv.panel_a.table
    sm = host.selectionModel()
    m = host.model()
    host._select_rows([5])   # 변경 행(5) 전체(행 헤더 선택 상당)
    QApplication.instance().processEvents()
    for c in (2, 3):        # 키 행(0,1)은 선택되지 않아야 함
        assert not sm.isSelected(m.index(0, c)) and not sm.isSelected(m.index(1, c)), \
            f"키 행이 보충됨 (열 {c})"
    assert host._full_rows_selected() == [5]


def test_single_cell_selection_not_supplemented(frozen_view):
    """단일 셀(내비게이션 착지)은 딱 그 셀만 — 키 열/행 보충 없음."""
    dv = frozen_view
    host = dv.panel_a.table
    sm = host.selectionModel()
    m = host.model()
    sm.select(QItemSelection(m.index(2, 2), m.index(2, 2)),
              QItemSelectionModel.ClearAndSelect)
    QApplication.instance().processEvents()
    assert sm.isSelected(m.index(2, 2))
    assert not sm.isSelected(m.index(2, 0)), "단일 셀인데 키 열이 보충됨"
    assert not sm.isSelected(m.index(0, 2)), "단일 셀인데 키 행이 보충됨"


def test_select_all_not_supplemented_or_fragmented(frozen_view):
    """selectAll 은 range 조각화 없이 전 열을 그대로 보고한다(키 밴드 포함).

    선택에 셀을 더하지 않으므로(엑셀식) selectAll 의 결과는 통짜 range 하나다.
    조각화되면 _full_columns_selected 가 일부 열을 놓친다."""
    dv = frozen_view
    host = dv.panel_a.table
    host.selectAll()
    QApplication.instance().processEvents()
    # 키 밴드(0,1)도 실제로 선택돼 있으므로 그대로 보고돼야 한다.
    expected = list(range(host.columnCount()))
    assert host._full_columns_selected() == expected, host._full_columns_selected()


def test_selection_mirrors_to_other_panel(frozen_view):
    """A 패널 선택이 반대 패널로 그대로 미러링된다 — 선택한 열만(키 열 보충 없음)."""
    dv = frozen_view
    ha = dv.panel_a.table
    hb = dv.panel_b.table
    ha._select_col(2)   # 데이터 열 전체 선택
    QApplication.instance().processEvents()
    smb = hb.selectionModel()
    assert smb.isSelected(hb.model().index(5, 2)), "반대 패널에 선택 미러 안 됨"
    assert not smb.isSelected(hb.model().index(5, 0)), "미러에 키 열이 붙음"


def test_normalize_no_infinite_loop(frozen_view):
    """선택 정규화(프룬)가 selectionChanged 재귀로 폭주하지 않는다(_supplementing 복원 확인)."""
    dv = frozen_view
    host = dv.panel_a.table
    host._select_col(2)
    QApplication.instance().processEvents()
    assert host._supplementing is False   # 항상 정상 복원(try/finally)



# ── 헤더 좌클릭 선택 (헤더 교체 회귀 방지) ──────────────────────────────────
def _click(widget, pos):
    """실제 좌클릭 시퀀스(press → release)를 위젯에 보낸다."""
    gp = widget.mapToGlobal(pos)
    for et, buttons in ((QEvent.MouseButtonPress, Qt.LeftButton),
                        (QEvent.MouseButtonRelease, Qt.NoButton)):
        QApplication.sendEvent(widget, QMouseEvent(
            et, pos, gp, Qt.LeftButton, buttons, Qt.NoModifier))
        QApplication.instance().processEvents()


def test_headers_are_clickable(frozen_view):
    """ExcelTableView 는 헤더를 _BandHeaderView 로 교체한다. QTableView 가 **자기가 만든
    기본 헤더에만** 켜 주는 sectionsClickable/highlightSections 를 교체본에 다시 켜지
    않으면 헤더 좌클릭이 sectionPressed 를 아예 안 쏜다."""
    t = frozen_view.panel_a.table
    for hdr, name in ((t.horizontalHeader(), "가로"), (t.verticalHeader(), "세로")):
        assert hdr.sectionsClickable(), f"{name} 헤더가 클릭 불가 — 헤더 선택이 죽는다"
        assert hdr.highlightSections(), f"{name} 헤더 선택 하이라이트 꺼짐"


def test_col_header_click_selects_that_column(frozen_view):
    """데이터 열 헤더를 좌클릭하면 그 열만 선택되고 앵커가 그 열로 갱신된다.

    앵커 갱신은 sectionPressed 로만 일어난다 — 안 쏘면 _header_anchor_col 이 낡은 값에
    머물러, 살짝만 끌어도 그 낡은 앵커부터 통째로 선택된다(F 를 눌렀는데 A~F 선택).
    """
    dv = frozen_view
    t = dv.panel_a.table
    hh = t.horizontalHeader()
    body_cols = [c for c in range(t.model().data_cols) if not t.isColumnHidden(c)]
    assert len(body_cols) >= 2, "전제: 고정 밴드 밖 데이터 열이 2개 이상"
    for c in body_cols[:2]:
        t.clearSelection()
        x = hh.sectionViewportPosition(c) + hh.sectionSize(c) // 2
        _click(hh.viewport(), QPoint(x, hh.viewport().height() // 2))
        sel = sorted({i.column() for i in t.selectionModel().selectedIndexes()})
        assert sel == [c], f"{c}열 헤더 클릭 → 선택 {sel}"
        assert t._header_anchor_col == c, "헤더 앵커가 갱신 안 됨(낡은 앵커 오염)"


def test_row_header_click_selects_that_row(frozen_view):
    """가로의 세로 대칭 — 행 헤더 좌클릭도 동일하게 동작해야 한다."""
    t = frozen_view.panel_a.table
    vh = t.verticalHeader()
    r = next(r for r in range(t.model().data_rows) if not t.isRowHidden(r))
    t.clearSelection()
    y = vh.sectionViewportPosition(r) + vh.sectionSize(r) // 2
    _click(vh.viewport(), QPoint(vh.viewport().width() // 2, y))
    rows = sorted({i.row() for i in t.selectionModel().selectedIndexes()})
    assert rows == [r], f"{r}행 헤더 클릭 → 선택 {rows}"
    assert t._header_anchor_row == r


# ── 고정 밴드 셀 우클릭 (병합 메뉴 배선) ────────────────────────────────────
def test_frozen_overlays_have_cell_context_menu(frozen_view):
    """키 열·그 좌측 열·키 행 셀은 본체에서 숨겨져 오버레이가 그린다. 오버레이에 셀
    우클릭이 배선돼 있지 않으면 그 셀들은 우클릭해도 아무 메뉴가 안 뜬다."""
    fc = frozen_view._freeze["a"]
    for name, v in (("left", fc.left), ("top", fc.top), ("corner", fc.corner)):
        assert v.contextMenuPolicy() == Qt.CustomContextMenu, f"{name} 오버레이 우클릭 미배선"
        assert v.receivers(v.customContextMenuRequested) > 0, f"{name} 오버레이 핸들러 없음"


def test_frozen_band_cell_right_click_opens_merge_menu(frozen_view):
    """고정 밴드의 '변경' 셀을 우클릭하면 본체와 동일한 병합 준비 메뉴가 떠야 한다."""
    dv = frozen_view
    app = QApplication.instance()
    t = dv.panel_a.table
    fc = dv._freeze["a"]
    r = 5   # 데이터 행(키 행=1 이므로 2행부터 본문)
    dv._diff_matrix[r][0] = ("modified", "old", "new")   # 고정 열 0 을 변경 상태로
    dv._notify_cells({(r, 0)})
    app.processEvents()
    assert t.model().cell_kind(r, 0) == "changed"

    rect = fc.left.visualRect(fc.left.model().index(r, 0))
    assert rect.isValid() and rect.width() > 0, "left 오버레이가 고정 셀을 그리지 않음"
    t._select_range(r, 0, r, 0)
    app.processEvents()

    before = {w for w in app.topLevelWidgets() if isinstance(w, QMenu)}
    fc._overlay_cell_menu(fc.left, rect.center())
    app.processEvents()
    menus = [w for w in app.topLevelWidgets()
             if isinstance(w, QMenu) and w.isVisible() and w not in before]
    try:
        assert menus, "고정 밴드 셀 우클릭에 병합 준비 메뉴가 안 뜬다"
        texts = [a.text() for a in menus[0].actions() if a.text()]
        assert any("A → B" in x for x in texts), texts
    finally:
        for m in menus:
            m.close()
        app.processEvents()


# ── 헤더 다중 선택 → 병합 대상 (키 밴드 누락 회귀) ──────────────────────────
def test_header_multiselect_spanning_key_band_keeps_all_columns(frozen_view):
    """키 밴드를 가로지르는 헤더 다중 선택은 그 밴드도 병합 대상에 남아야 한다.

    회귀: _full_columns_selected 가 키 밴드(0..key_col)를 기본으로 걸러내, A~G 헤더를
    잡고 병합 준비해도 A~D 가 조용히 빠졌다(E~G 만 준비됨). 그 필터는 선택에 키 열을
    자동 보충하던 시절의 잔재이고, 지금은 '터치한 셀만' 선택하므로 걸러낼 이유가 없다.
    """
    host = frozen_view.panel_a.table
    kc = host._key_col
    last = min(kc + 3, host.model().data_cols - 1)
    assert last > kc, "전제: 키 밴드 오른쪽에 데이터 열이 있어야 한다"
    host._select_column_range(0, last)          # A..last 헤더 다중 선택
    QApplication.instance().processEvents()

    want = list(range(0, last + 1))
    assert host._full_columns_selected() == want, host._full_columns_selected()
    # 우클릭 앵커가 데이터 열이든 키 밴드든 대상 집합은 선택 전체
    assert host._selected_header_cols(last) == want, "키 밴드가 병합 대상에서 빠짐"
    assert host._selected_header_cols(0) == want, "키 밴드 앵커에서도 선택 전체여야"


def test_single_data_column_header_does_not_pull_key_band(frozen_view):
    """반대 방향 회귀 — 데이터 열 하나만 선택하면 키 밴드가 딸려오면 안 된다."""
    host = frozen_view.panel_a.table
    data_col = host._key_col + 1
    host._select_col(data_col)
    QApplication.instance().processEvents()
    assert host._full_columns_selected() == [data_col]
    assert host._selected_header_cols(data_col) == [data_col]


def test_header_multiselect_spanning_key_row_band_keeps_all_rows(frozen_view):
    """행 축 대칭 — 키 행 밴드를 가로지르는 행 헤더 다중 선택도 그대로 유지."""
    dv = frozen_view
    dv.diff_only_btn.setChecked(False)   # 전체 행 표시
    QApplication.instance().processEvents()
    host = dv.panel_a.table
    kr = host._key_row
    last = min(kr + 3, host.model().data_rows - 1)
    assert last > kr, "전제: 키 행 밴드 아래에 데이터 행이 있어야 한다"
    host._select_row_range(0, last)
    QApplication.instance().processEvents()
    want = list(range(0, last + 1))
    assert host._full_rows_selected() == want, host._full_rows_selected()
    assert host._selected_header_rows(last) == want, "키 행 밴드가 병합 대상에서 빠짐"
