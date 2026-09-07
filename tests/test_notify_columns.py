# -*- coding: utf-8 -*-
"""notify_columns 의 dataChanged 접기 — 커버 범위는 그대로, emit 횟수만 줄어든다.

'변경 검사에서 제외' 토글은 열마다 dataChanged 를 따로 쏘고 있었다. 큰 시트에서는
emit 자체가 지배적이다(실측 9,856행: 열 3개 제외/해제 1회에 emit 만 ~41ms).
바로 위 notify_cells 는 이미 대량 변경을 단일 bounding rect 로 접는데 열 쪽에는 없었다.

접기가 **덜 칠하는 일이 없어야** 하므로, 방출된 rect 들이 덮는 (행, 열) 집합이 예전
'열마다 하나씩' 구현과 정확히 같은지 오라클로 대조한다.
"""
import pytest
from PyQt5.QtCore import Qt

from excelmerge.diff_model import DiffTableModel
from excelmerge.diff_engine import compute_diff

HEADER = ["A", "B", "C", "D", "E", "F"]
A_DATA = [HEADER] + [[f"k{i}", "1", "2", "3", "4", "5"] for i in range(1, 9)]
B_DATA = [HEADER] + [[f"k{i}", "1", "9", "3", "4", "5"] for i in range(1, 9)]


@pytest.fixture
def model(qapp):
    matrix, row_meta = compute_diff(A_DATA, B_DATA, key_col=0, key_row=0)
    m = DiffTableModel("a")
    m.set_diff_data(matrix, row_meta, {}, set(), set())
    assert m.data_rows == 9 and m.data_cols == 6
    return m


def _capture(model, cols):
    """notify_columns 가 방출한 dataChanged rect 목록과 헤더 갱신 섹션 목록.
    한 모델을 여러 번 계측하므로 반드시 연결을 끊는다(안 끊으면 다음 호출이
    앞선 리스트에도 쌓여 비교가 무의미해진다)."""
    rects, headers = [], []

    def on_data(tl, br, roles=None):
        rects.append((tl.row(), tl.column(), br.row(), br.column()))

    def on_header(orient, first, last):
        headers.append((orient, first, last))

    model.dataChanged.connect(on_data)
    model.headerDataChanged.connect(on_header)
    try:
        model.notify_columns(cols)
    finally:
        model.dataChanged.disconnect(on_data)
        model.headerDataChanged.disconnect(on_header)
    return rects, headers


def _covered(rects):
    out = set()
    for r1, c1, r2, c2 in rects:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                out.add((r, c))
    return out


def _oracle_covered(model, cols):
    """접기 이전 구현 — 열마다 (0, c)~(last_row, c) 하나씩."""
    last_row = max(0, model.data_rows - 1)
    out = set()
    for c in cols:
        if 0 <= c < model.data_cols:
            for r in range(0, last_row + 1):
                out.add((r, c))
    return out


@pytest.mark.parametrize("cols", [
    [0], [1, 2, 3], [1, 2, 5], [5, 1, 2], [0, 5], list(range(6)),
    [2, 3, 99, -1], [], [99], [3, 4, 3],
])
def test_covered_cells_match_old_per_column_emits(model, cols):
    rects, _ = _capture(model, cols)
    assert _covered(rects) == _oracle_covered(model, cols), f"cols={cols}"


def test_contiguous_run_is_a_single_emit(model):
    rects, _ = _capture(model, [1, 2, 3])
    assert rects == [(0, 1, 8, 3)], rects


def test_gaps_split_into_separate_emits(model):
    rects, _ = _capture(model, [1, 2, 5])
    assert rects == [(0, 1, 8, 2), (0, 5, 8, 5)], rects


def test_unsorted_input_is_folded_the_same(model):
    assert _capture(model, [5, 2, 1])[0] == _capture(model, [1, 2, 5])[0]


def test_out_of_range_columns_emit_nothing(model):
    rects, _ = _capture(model, [99, -3])
    assert rects == []


def test_header_refresh_is_still_per_column(model):
    """헤더는 열 수만큼이라 저렴 — 접지 않고 그대로 둔다(아이콘/색 갱신 누락 방지)."""
    _, headers = _capture(model, [1, 2, 5])
    assert headers == [(Qt.Horizontal, 1, 1), (Qt.Horizontal, 2, 2),
                       (Qt.Horizontal, 5, 5)]
