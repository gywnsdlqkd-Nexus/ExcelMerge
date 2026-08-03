# -*- coding: utf-8 -*-
"""병합 준비(스테이징) 판단 규칙 단위 테스트 — **QApplication 없이** 실행된다.

이 규칙들은 원래 DiffView 안에서 UI 효과와 얽혀 있어, 검증하려면 MainWindow 를 띄워야
했다(그래서 사실상 미검증이었다). staging 모듈로 분리한 뒤 순수 함수로 직접 검증한다.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excelmerge import staging
from excelmerge.constants import DIR_A2B, DIR_B2A


def _cell(status, a="", b=""):
    return (status, a, b)


def _matrix():
    """3행 × 3열. (1,1) 변경, (2,2) 추가, 나머지 same."""
    return [
        [_cell("same", "ID", "ID"), _cell("same", "H1", "H1"), _cell("same", "H2", "H2")],
        [_cell("same", "1", "1"), _cell("modified", "old", "new"), _cell("same", "x", "x")],
        [_cell("same", "2", "2"), _cell("same", "y", "y"), _cell("added", "", "plus")],
    ]


# ── key_cells_for_selection ─────────────────────────────────────────────────
def test_key_cells_empty_selection():
    assert staging.key_cells_for_selection(set(), 0, 0) == set()


def test_key_cells_adds_key_col_for_selected_rows():
    got = staging.key_cells_for_selection({(3, 5)}, key_row=-1, key_col=1)
    assert got == {(3, 0), (3, 1)}


def test_key_cells_adds_key_row_for_selected_cols():
    got = staging.key_cells_for_selection({(3, 5)}, key_row=1, key_col=-1)
    assert got == {(0, 5), (1, 5)}


def test_key_cells_both_axes():
    got = staging.key_cells_for_selection({(2, 2)}, key_row=0, key_col=0)
    assert got == {(2, 0), (0, 2)}


def test_key_cells_none_and_negative_are_ignored():
    assert staging.key_cells_for_selection({(1, 1)}, None, None) == set()
    assert staging.key_cells_for_selection({(1, 1)}, -1, -1) == set()


# ── stageable_cells ─────────────────────────────────────────────────────────
def test_stageable_keeps_only_changed():
    m = _matrix()
    cells = {(0, 0), (1, 1), (2, 2), (1, 2)}
    assert staging.stageable_cells(m, cells, set()) == {(1, 1), (2, 2)}


def test_stageable_drops_excluded_cols():
    m = _matrix()
    assert staging.stageable_cells(m, {(1, 1), (2, 2)}, {1}) == {(2, 2)}


def test_stageable_drops_out_of_range():
    m = _matrix()
    cells = {(99, 1), (1, 99), (-1, 1), (1, -1), (1, 1)}
    assert staging.stageable_cells(m, cells, set()) == {(1, 1)}


def test_stageable_empty_matrix():
    assert staging.stageable_cells([], {(0, 0)}, set()) == set()


def test_stageable_ragged_row():
    """행 길이가 들쭉날쭉해도 그 행의 실제 길이로 판정해야 한다."""
    m = [[_cell("modified", "a", "b")], [_cell("modified", "c", "d"), _cell("modified", "e", "f")]]
    assert staging.stageable_cells(m, {(0, 1), (1, 1)}, set()) == {(1, 1)}


# ── staged_display_value ────────────────────────────────────────────────────
def test_display_value_direction():
    m = _matrix()
    assert staging.staged_display_value(m, 1, 1, DIR_A2B) == "old"   # A→B 는 A 값
    assert staging.staged_display_value(m, 1, 1, DIR_B2A) == "new"   # B→A 는 B 값


def test_display_value_out_of_range_is_blank():
    m = _matrix()
    assert staging.staged_display_value(m, 99, 1, DIR_A2B) == ""
    assert staging.staged_display_value(m, 1, 99, DIR_A2B) == ""


def test_display_value_malformed_cell_is_blank():
    assert staging.staged_display_value([[None]], 0, 0, DIR_A2B) == ""
    assert staging.staged_display_value([["oops"]], 0, 0, DIR_A2B) == ""


# ── staged_keys_in_cols / excludable_cols ───────────────────────────────────
def test_staged_keys_in_cols():
    staged = {(1, 1): DIR_A2B, (2, 2): DIR_A2B, (3, 1): DIR_B2A}
    assert sorted(staging.staged_keys_in_cols(staged, [1])) == [(1, 1), (3, 1)]
    assert staging.staged_keys_in_cols(staged, []) == []


def test_excludable_cols_filters_key_and_already_excluded():
    # 키 열(0)은 제외 불가, 이미 제외된 2는 중복 제외 대상 아님
    assert staging.excludable_cols([0, 1, 2, 3], key_col=0, excluded_cols={2}) == [1, 3]


def test_excludable_cols_all_filtered():
    assert staging.excludable_cols([0], key_col=0, excluded_cols=set()) == []
