# -*- coding: utf-8 -*-
"""병합 준비는 **고른 셀만** 준비한다 — 키 열/행 자동 보충 없음.

예전에는 선택 행에 키 열 0..key_col, 선택 열에 키 행 0..key_row 를 자동으로 끼워 넣었다
(틀 고정으로 숨겨진 키 셀이 러버밴드 선택에 안 잡히는 걸 메우려는 보충).

문제: key_col **왼쪽** 열(#Description 등)은 키가 아니라 단지 고정 밴드에 있을 뿐이다.
신규(A 에만 있는) 행은 전 열이 added 라 status 필터로도 안 걸러져서, E 열 셀 하나를
병합 준비했을 뿐인데 B~D 까지 함께 준비됐다. 자동 보충은 사용자가 고르지도 않은 셀을
말없이 끼워 넣는 자리라 제거했다.

행/열 전체를 준비하려면 헤더로 전체를 선택하면 된다 — 그때는 키 열/행이 '실제 선택'에
들어오므로 그대로 포함된다(아래 두 번째 테스트).
"""
import pytest
from PyQt5.QtWidgets import QApplication

from excelmerge.constants import DIR_A2B

HEADER = ["UniqueID", "#Description", "#Description2", "TextID", "ko", "en"]
KEY_COL, KEY_ROW = 3, 0          # TextID(D) 를 키 열로, 0행이 헤더


def _wait_diff(dv, timeout_ms=5000):
    w = getattr(dv, "_diff_worker", None)
    if w is not None:
        w.wait(timeout_ms)
    QApplication.instance().processEvents()


@pytest.fixture
def staged_view(qapp):
    """A 에만 있는 신규 행(전 열 added)을 가진 활성 DiffView."""
    from excelmerge.main_window import MainWindow
    win = MainWindow()
    win.show()
    try:
        dv = win.tabs.currentWidget()
        dv.panel_a.set_path("a.xlsx"); dv.panel_b.set_path("b.xlsx")
        dv._key_col, dv._key_row = KEY_COL, KEY_ROW
        for p in (dv.panel_a, dv.panel_b):
            p.table.set_key_col(KEY_COL); p.table.set_key_row(KEY_ROW)
        a = [HEADER,
             ["1", "desc", "desc2", "KEY_BOTH", "ko1", "en1"],
             ["2", "새설명", "새설명2", "KEY_A_ONLY", "새 ko", "새 en"]]
        b = [HEADER,
             ["1", "desc", "desc2", "KEY_BOTH", "ko1", "en1"]]
        dv._on_loaded(a, b); _wait_diff(dv)
        # A 전용 행(= B 에 없는 신규 행) 찾기
        dv._new_row = next(
            i for i, (a_orig, b_orig) in enumerate(dv._diff_row_meta) if b_orig is None)
        assert all(dv._diff_matrix[dv._new_row][c][0] == "added" for c in range(len(HEADER))), \
            "전제: 신규 행은 전 열이 added"
        yield dv
    finally:
        win.close(); win.deleteLater(); qapp.processEvents()


def test_single_cell_stages_only_that_cell(staged_view):
    """신규 행의 E(4) 셀 하나만 준비 → 그 셀만. 키 열(D)·그 왼쪽(B·C)이 따라오면 안 된다."""
    dv = staged_view
    r = dv._new_row
    dv.panel_a.table._select_range(r, 4, r, 4)
    QApplication.instance().processEvents()
    dv._stage_selected(DIR_A2B)
    QApplication.instance().processEvents()
    assert set(dv._staged) == {(r, 4)}, (
        f"고른 셀 외에 {sorted(set(dv._staged) - {(r, 4)})} 까지 병합 준비됨")


def test_full_row_selection_still_stages_key_band(staged_view):
    """행 헤더로 전체 선택하면 키 열/그 왼쪽 열도 '실제 선택'이라 그대로 준비된다."""
    dv = staged_view
    r = dv._new_row
    dv.panel_a.table._select_rows([r])
    QApplication.instance().processEvents()
    dv._stage_selected(DIR_A2B)
    QApplication.instance().processEvents()
    assert set(dv._staged) == {(r, c) for c in range(len(HEADER))}, sorted(dv._staged)
