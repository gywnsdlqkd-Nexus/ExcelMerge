# -*- coding: utf-8 -*-
"""3파일 비교 엔진 — 가운데(기준) 대비 왼쪽·오른쪽 변경을 맞게 나누는가.

판정 규칙을 표로 못 박고, 2파일 엔진과의 등가성(한쪽이 기준과 같으면 다른 쪽 변경 = 2파일
비교 결과)을 확인한다. Qt 없이 도는 순수 함수 테스트다.
"""
import random

import pytest

from excelmerge.diff3_engine import (
    S3_BOTH, S3_CONFLICT, S3_LEFT, S3_RIGHT, S3_SAME, SRC_LEFT, SRC_RIGHT,
    auto_merge_plan, classify, compute_diff3, conflict_masks, count_dropped3, match_columns3,
)
from excelmerge.diff_engine import compute_diff, match_columns, row_change_masks

H = ["ID", "V", "W"]


def _grid(m):
    """{키: [상태...]} — 키 열(0)로 행을 찾는다."""
    return {row[0][2] or row[0][1] or row[0][3]: [c[0] for c in row] for row in m}


def _by_key(m, meta):
    out = {}
    for row, mm in zip(m, meta):
        k = next(v for v in row[0][1:] if v)
        out[k] = (row, mm)
    return out


# ── 셀 판정 ─────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("l,b,r,want", [
    ("x", "x", "x", S3_SAME),
    ("y", "x", "x", S3_LEFT),
    ("x", "x", "y", S3_RIGHT),
    ("y", "x", "y", S3_BOTH),
    ("y", "x", "z", S3_CONFLICT),
    ("", "x", "x", S3_LEFT),       # 왼쪽이 지움
    ("y", "", "", S3_LEFT),        # 왼쪽이 채움
    ("y", "", "y", S3_BOTH),
    ("", "x", "", S3_BOTH),        # 양쪽이 지움
])
def test_cell_classification(l, b, r, want):
    assert classify(l, b, r) == want


# ── 키 기준 행 맞춤 ─────────────────────────────────────────────────────────

def test_changes_are_split_by_side():
    base = [H, ["1", "a", "p"], ["2", "b", "q"], ["3", "c", "r"]]
    left = [H, ["1", "A", "p"], ["2", "b", "q"], ["3", "C", "r"]]
    right = [H, ["1", "a", "P"], ["2", "B", "q"], ["3", "X", "r"]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    g = _grid(m)
    assert g["1"] == [S3_SAME, S3_LEFT, S3_RIGHT]
    assert g["2"] == [S3_SAME, S3_RIGHT, S3_SAME]
    assert g["3"] == [S3_SAME, S3_CONFLICT, S3_SAME]
    assert meta[1:] == [(1, 1, 1), (2, 2, 2), (3, 3, 3)]


def test_new_rows_keep_their_place():
    """새 행은 그 파일에서의 이웃 사이에 보인다 — 맨 아래로 밀리지 않는다."""
    base = [H, ["1", "a", ""], ["3", "c", ""]]
    left = [H, ["1", "a", ""], ["2", "L2", ""], ["3", "c", ""]]
    right = [H, ["1", "a", ""], ["3", "c", ""], ["4", "R4", ""]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert [next(v for v in row[0][1:] if v) for row in m[1:]] == ["1", "2", "3", "4"]
    rows = _by_key(m, meta)
    assert rows["2"][1] == (2, None, None) and rows["2"][0][1][0] == S3_LEFT
    assert rows["4"][1] == (None, None, 3) and rows["4"][0][1][0] == S3_RIGHT


def test_the_same_new_row_on_both_sides():
    base = [H, ["1", "a", ""]]
    left = [H, ["1", "a", ""], ["9", "same", "L"]]
    right = [H, ["1", "a", ""], ["9", "same", "R"]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    row, mm = _by_key(m, meta)["9"]
    assert mm == (2, None, 2)
    assert [c[0] for c in row] == [S3_BOTH, S3_BOTH, S3_CONFLICT]


def test_a_row_deleted_on_one_side_and_untouched_on_the_other():
    base = [H, ["1", "a", "p"], ["2", "b", "q"]]
    left = [H, ["1", "a", "p"]]
    right = [H, ["1", "a", "p"], ["2", "b", "q"]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    row, mm = _by_key(m, meta)["2"]
    assert mm == (None, 2, 2)
    assert [c[0] for c in row] == [S3_LEFT, S3_LEFT, S3_LEFT]


def test_delete_versus_edit_is_a_whole_row_conflict():
    """한쪽은 지우고 다른 쪽은 고쳤다 — 셀 단위로 갈라 반쪽만 병합되지 않게 행 전체를 충돌로."""
    base = [H, ["1", "a", "p"], ["2", "b", "q"]]
    left = [H, ["1", "a", "p"]]
    right = [H, ["1", "a", "p"], ["2", "b", "Q"]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    row, _mm = _by_key(m, meta)["2"]
    assert [c[0] for c in row] == [S3_CONFLICT, S3_CONFLICT, S3_CONFLICT]
    assert not any(k[0] == len(m) - 1 for k in auto_merge_plan(m)), "충돌 행을 자동 병합하려 한다"


def test_both_sides_deleted_the_row():
    base = [H, ["1", "a", ""], ["2", "b", ""]]
    left = [H, ["1", "a", ""]]
    right = [H, ["1", "a", ""]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    row, mm = _by_key(m, meta)["2"]
    assert mm == (None, 2, None)
    assert [c[0] for c in row] == [S3_BOTH, S3_BOTH, S3_SAME]


# ── 열 맞춤 ─────────────────────────────────────────────────────────────────

def test_columns_added_on_each_side_are_placed():
    base = [["ID", "V"], ["1", "a"]]
    left = [["ID", "NEW_L", "V"], ["1", "n", "a"]]
    right = [["ID", "V", "NEW_R"], ["1", "a", "r"]]
    cm = match_columns3(left, base, right, 0)
    assert cm == [(0, 0, 0), (1, None, None), (2, 1, 1), (None, None, 2)]
    m, _meta = compute_diff3(left, base, right, 0, 0, cm)
    assert [c[0] for c in m[1]] == [S3_SAME, S3_LEFT, S3_SAME, S3_RIGHT]


def test_the_same_new_column_on_both_sides_shows_once():
    base = [["ID", "V"], ["1", "a"]]
    left = [["ID", "V", "TAG"], ["1", "a", "x"]]
    right = [["ID", "V", "TAG"], ["1", "a", "x"]]
    cm = match_columns3(left, base, right, 0)
    assert cm == [(0, 0, 0), (1, 1, 1), (2, None, 2)]
    m, _ = compute_diff3(left, base, right, 0, 0, cm)
    assert m[1][2][0] == S3_BOTH


def test_unnamed_headers_fall_back_to_positions():
    """가운데 빈 헤더는 이름으로 맞출 수 없다(끝에 붙은 빈 칸은 열로 보지 않는다 — 2파일과 같음)."""
    base = [["ID", "", "W"], ["1", "a", "w"]]
    assert match_columns3(base, base, base, 0) is None


# ── 2파일 엔진과의 등가성 ───────────────────────────────────────────────────

def _random_table(rng, keys, cols):
    rows = [["ID"] + [f"C{i}" for i in range(1, cols)]]
    for k in keys:
        rows.append([k] + [rng.choice(["", "x", "y", "z"]) for _ in range(1, cols)])
    return rows


@pytest.mark.parametrize("seed", range(20))
def test_one_side_equal_to_base_matches_the_two_file_engine(seed):
    """왼쪽이 기준과 같으면, 3파일의 '오른쪽만 바뀜' 칸 = 2파일 비교(기준, 오른쪽)의 변경 칸."""
    rng = random.Random(seed)
    keys = [str(i) for i in range(30)]
    base = _random_table(rng, rng.sample(keys, 20), 5)
    right = _random_table(rng, rng.sample(keys, 20), 5)
    m3, meta3 = compute_diff3(base, base, right, 0, 0)
    m2, meta2 = compute_diff(base, right, 0, 0)
    assert [(b, r) for _l, b, r in meta3] == meta2, "행 순서·대응이 2파일 엔진과 다르다"
    for row3, row2 in zip(m3, m2):
        for c3, c2 in zip(row3, row2):
            assert (c3[0] == S3_RIGHT) == (c2[0] != "same"), (c3, c2)
            assert c3[0] in (S3_SAME, S3_RIGHT)


# ── 행 순서 모드 · 보조 함수 ────────────────────────────────────────────────

def test_row_order_mode_is_positional():
    base = [["a"], ["b"]]
    left = [["a"], ["B"], ["new"]]
    right = [["A"], ["b"]]
    m, meta = compute_diff3(left, base, right, -1, 0)
    assert [c[0][0] for c in m] == [S3_RIGHT, S3_LEFT, S3_LEFT]
    assert meta == [(0, 0, 0), (1, 1, 1), (2, None, None)]


def test_masks_and_auto_merge_plan():
    base = [H, ["1", "a", "p"], ["2", "b", "q"]]
    left = [H, ["1", "A", "p"], ["2", "B", "q"]]
    right = [H, ["1", "a", "P"], ["2", "X", "q"]]
    m, _ = compute_diff3(left, base, right, 0, 0)
    assert row_change_masks(m) == [0, 0b110, 0b010]
    assert conflict_masks(m) == [0, 0, 0b010]
    assert auto_merge_plan(m) == {(1, 1): SRC_LEFT, (1, 2): SRC_RIGHT}


def test_dropped_rows_are_counted_across_the_three_files():
    base = [H, ["1", "a", ""], ["1", "dup", ""], ["", "nokey", ""]]
    left = [H, ["1", "a", ""]]
    right = [H, ["1", "a", ""], ["1", "dup", ""]]
    assert count_dropped3(left, base, right, None, 0, 0) == 3


def test_large_sheet_is_fast_enough():
    """2만 행 × 40열 — 2파일 엔진(약 0.35초/2만×71)과 같은 자릿수여야 한다."""
    import time
    rng = random.Random(1)
    n, cols = 20000, 40
    base = [["ID"] + [f"C{i}" for i in range(1, cols)]] + \
        [[str(i)] + [f"v{i}_{c}" for c in range(1, cols)] for i in range(n)]
    left = [row[:] for row in base]
    right = [row[:] for row in base]
    for i in rng.sample(range(1, n), 500):
        left[i][3] = "L"
    for i in rng.sample(range(1, n), 500):
        right[i][5] = "R"
    t0 = time.perf_counter()
    m, _ = compute_diff3(left, base, right, 0, 0, match_columns3(left, base, right, 0))
    took = time.perf_counter() - t0
    assert len(m) == n + 1
    assert took < 3.0, f"{took:.2f}초"
