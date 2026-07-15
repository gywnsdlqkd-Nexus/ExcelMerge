# -*- coding: utf-8 -*-
"""다중 선택 스테이징 시 키 열/행 셀 보충 — 틀 고정으로 숨겨진 키 셀이 병합에서 누락되지 않도록.

버그: 키 열/행은 본체에서 '틀 고정'으로 숨겨져 러버밴드 셀 선택 range 에 안 잡혀,
신규(added) 행을 복사할 때 키 값(UniqueID 등)이 빠졌다.
"""
import pytest


@pytest.fixture
def diff_view(qapp):
    from excelmerge.main_window import MainWindow
    w = MainWindow()
    try:
        yield w.tabs.currentWidget()
    finally:
        w.close(); w.deleteLater(); qapp.processEvents()


def test_key_col_supplemented_for_selected_rows(diff_view):
    diff_view._key_col = 0
    diff_view._key_row = 0
    extra = diff_view._key_cells_for_selection({(5, 3), (6, 4)})
    assert (5, 0) in extra and (6, 0) in extra   # 선택 행에 키 열(0) 보충
    assert (0, 3) in extra and (0, 4) in extra   # 선택 열에 키 행(0) 보충


def test_key_col_multi_column_key(diff_view):
    diff_view._key_col = 2          # 키 열 = C → 0,1,2 고정
    diff_view._key_row = 0
    extra = diff_view._key_cells_for_selection({(5, 4)})
    assert {(5, 0), (5, 1), (5, 2)} <= extra


def test_no_key_when_row_order_mode(diff_view):
    diff_view._key_col = -1          # ROW 순서 비교(키 열 없음)
    diff_view._key_row = 0
    extra = diff_view._key_cells_for_selection({(5, 4)})
    # 키 열 보충 없음, 키 행(0)만(선택 열 4)
    assert not any(c < 0 or (r, c) == (5, 0) for (r, c) in extra if c != 4)
    assert (0, 4) in extra


def test_empty_selection(diff_view):
    assert diff_view._key_cells_for_selection(set()) == set()
