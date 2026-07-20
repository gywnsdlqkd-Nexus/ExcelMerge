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
import tempfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ["APPDATA"] = tempfile.mkdtemp(prefix="em_test_appdata_")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from PyQt5.QtCore import QItemSelection, QItemSelectionModel
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


# ── 버그 2: 다중 선택 시 키 열/행 항상 함께 선택 (key_col=1, key_row=1) ────────
def test_block_always_supplements_key_col_and_row(frozen_view):
    """중간 블록(행 5~7 × 열 2~3, 어느 경계에도 안 닿음)이라도 키 열(0~1)·키 행(0~1)이 함께 선택."""
    dv = frozen_view
    host = dv.panel_a.table
    sm = host.selectionModel()
    m = host.model()
    sm.select(QItemSelection(m.index(5, 2), m.index(7, 3)),
              QItemSelectionModel.ClearAndSelect)
    QApplication.instance().processEvents()
    for r in (5, 6, 7):   # 선택 행에 키 열(0,1)
        assert sm.isSelected(m.index(r, 0)) and sm.isSelected(m.index(r, 1)), \
            f"키 열 미보충 (행 {r})"
    for c in (2, 3):      # 선택 열에 키 행(0,1)
        assert sm.isSelected(m.index(0, c)) and sm.isSelected(m.index(1, c)), \
            f"키 행 미보충 (열 {c})"
    assert sm.isSelected(m.index(6, 2))   # 원래 선택 유지


def test_header_column_selection_supplements_key_col(frozen_view):
    """헤더로 데이터 열 전체 선택(_select_col) 시 키 열(0,1)도 함께 선택되고,
    _full_columns_selected 는 키 열을 제외해 데이터 열만 보고한다(헤더 확장·메뉴 시맨틱 보존)."""
    dv = frozen_view
    host = dv.panel_a.table
    host._select_col(2)
    QApplication.instance().processEvents()
    sm = host.selectionModel()
    assert sm.isSelected(host.model().index(5, 0)), "헤더 열 선택 시 키 열 미보충"
    assert sm.isSelected(host.model().index(5, 1))
    assert host._full_columns_selected() == [2], host._full_columns_selected()


def test_header_row_selection_supplements_key_row(frozen_view):
    """헤더로 데이터 행 전체 선택(_select_rows) 시 키 행(0,1)도 함께 선택되고,
    _full_rows_selected 는 키 행을 제외해 데이터 행만 보고한다."""
    dv = frozen_view
    host = dv.panel_a.table
    host._select_rows([5, 6])
    QApplication.instance().processEvents()
    sm = host.selectionModel()
    assert sm.isSelected(host.model().index(0, 3)), "헤더 행 선택 시 키 행 미보충"
    assert sm.isSelected(host.model().index(1, 3))
    assert host._full_rows_selected() == [5, 6], host._full_rows_selected()


def test_diff_only_row_selection_supplements_key_row(frozen_view):
    """회귀(사용자 보고): '변경 행만 보기' ON에서 변경 행 셀을 선택하면 숨겨진 키 행이
    함께 선택돼야 한다. (키 행은 본체에서 항상 숨김 → 격자 인접이 아닌 구조 조건으로 판정)"""
    dv = frozen_view
    dv.diff_only_btn.setChecked(True)   # 변경 행만 보기 ON
    QApplication.instance().processEvents()
    host = dv.panel_a.table
    sm = host.selectionModel()
    m = host.model()
    # 변경 행(5)의 데이터 셀(열 2~3) 선택 — diff-only라 위쪽 데이터 행들은 숨김 상태.
    sm.select(QItemSelection(m.index(5, 2), m.index(5, 3)),
              QItemSelectionModel.ClearAndSelect)
    QApplication.instance().processEvents()
    for c in (2, 3):      # 키 행(0,1) 보충 — 숨겨져 있어도 상단 밴드에 표시
        assert sm.isSelected(m.index(0, c)) and sm.isSelected(m.index(1, c)), \
            f"변경 행만 보기 ON: 키 행 미보충 (열 {c})"
    assert sm.isSelected(m.index(5, 0)) and sm.isSelected(m.index(5, 1))  # 키 열도 함께


def test_single_cell_selection_not_supplemented(frozen_view):
    """단일 셀(내비게이션 착지)은 보충하지 않는다 — 결정론적 단일 선택 유지."""
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
    """selectAll 은 이미 키를 포함하므로 보충 안 함 → range 조각화 없이 데이터 열 전체 보고."""
    dv = frozen_view
    host = dv.panel_a.table
    host.selectAll()
    QApplication.instance().processEvents()
    # 조각화되면 _full_columns_selected 가 일부 열을 놓친다 → 키 열(0,1) 제외 전 열이 연속으로.
    expected = list(range(2, host.columnCount()))   # key_col=1 → 첫 데이터 열=2
    assert host._full_columns_selected() == expected, host._full_columns_selected()


def test_supplement_mirrors_to_other_panel(frozen_view):
    """보충된 키 셀을 포함한 선택이 반대 패널로도 미러링된다."""
    dv = frozen_view
    ha = dv.panel_a.table
    hb = dv.panel_b.table
    m = ha.model()
    ha.selectionModel().select(QItemSelection(m.index(5, 2), m.index(7, 3)),
                               QItemSelectionModel.ClearAndSelect)
    QApplication.instance().processEvents()
    smb = hb.selectionModel()
    assert smb.isSelected(hb.model().index(5, 0)), "반대 패널에 키 셀 미러 안 됨"
    assert smb.isSelected(hb.model().index(0, 2))


def test_supplement_no_infinite_loop(frozen_view):
    """보충이 selectionChanged 재귀로 폭주하지 않는다(_supplementing 가드 복원 확인)."""
    dv = frozen_view
    host = dv.panel_a.table
    m = host.model()
    host.selectionModel().select(QItemSelection(m.index(2, 2), m.index(9, 3)),
                                 QItemSelectionModel.ClearAndSelect)
    QApplication.instance().processEvents()
    assert host._supplementing is False   # 항상 정상 복원(try/finally)
