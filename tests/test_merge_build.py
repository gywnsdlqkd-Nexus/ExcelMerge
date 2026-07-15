# -*- coding: utf-8 -*-
"""merge_build.build_side_patches — staged 셀 → 파일 패치 구성(순수, Qt/파일 없음)."""
from excelmerge.merge_build import build_side_patches
from excelmerge.constants import DIR_A2B, DIR_B2A


# diff_matrix[r][c] = (status, a_val, b_val), row_meta[r] = (a_orig, b_orig)
def _matrix():
    return [
        [("same", "ID", "ID"), ("modified", "A2", "B2")],   # r0
        [("added", "onlyA", ""), ("added", "x", "")],       # r1 (B에 행 없음)
    ]


def test_a2b_overwrite_and_style_src():
    m = _matrix()
    row_meta = [(0, 0), (1, 1)]                 # 두 행 모두 A·B에 존재
    staged = {(0, 1): DIR_A2B}                  # r0c1 을 A→B
    patches, insert_rows, style_src = build_side_patches(DIR_A2B, m, row_meta, staged)
    # 대상 B의 원본행=0, 열=1 → cell_ref(0,1)="B1" 에 A값("A2") 기록
    assert patches == {"B1": "A2"}, patches
    assert style_src == {"B1": "B1"}, style_src     # 소스 A좌표 cell_ref(0,1)="B1"
    assert insert_rows == {}


def test_a2b_insert_when_target_row_missing():
    m = _matrix()
    row_meta = [(0, 0), (1, None)]              # r1 은 B에 행 없음(b_orig=None)
    staged = {(1, 0): DIR_A2B, (1, 1): DIR_A2B}
    patches, insert_rows, style_src = build_side_patches(DIR_A2B, m, row_meta, staged)
    assert patches == {}                        # 덮어쓸 대상 행이 없음
    assert 1 in insert_rows and len(insert_rows[1]) == 2
    cols = sorted(c for (c, _v, _s) in insert_rows[1])
    assert cols == [0, 1]


def test_b2a_uses_b_value_and_targets_a():
    m = _matrix()
    row_meta = [(0, 0), (1, 1)]
    staged = {(0, 1): DIR_B2A}                  # r0c1 을 B→A
    patches, insert_rows, style_src = build_side_patches(DIR_B2A, m, row_meta, staged)
    # 대상 A의 원본행=0, 열=1 → cell_ref(0,1)="B1" 에 B값("B2") 기록
    assert patches == {"B1": "B2"}, patches
    assert insert_rows == {}


def test_direction_filter_isolates_cells():
    m = _matrix()
    row_meta = [(0, 0), (1, 1)]
    staged = {(0, 0): DIR_A2B, (0, 1): DIR_B2A}
    a2b_patches, _, _ = build_side_patches(DIR_A2B, m, row_meta, staged)
    b2a_patches, _, _ = build_side_patches(DIR_B2A, m, row_meta, staged)
    assert set(a2b_patches) == {"A1"}           # r0c0 → cell_ref(0,0)=="A1"
    assert set(b2a_patches) == {"B1"}           # r0c1 → cell_ref(0,1)=="B1"
