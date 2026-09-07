# -*- coding: utf-8 -*-
"""행별 '변경 열' 비트마스크 회귀 — 매트릭스와 갈라지면 안 된다.

'변경점만 보기' 필터 / 미니맵 / 변경점 이동(Alt+화살표) / 상태바 변경 셀 수는 모두
"이 행에 (제외 열을 뺀) 변경이 있나"를 묻는다. 예전엔 넷 다 행마다
`any(st != SAME for c, (st, *_) in enumerate(row) if c not in excl)` 로 O(C) 튜플
언패킹을 되풀이했다(6328행 x 71열에서 미니맵 122ms, 개수 47ms).

지금은 DiffWorker 가 행별 비트마스크를 함께 만들어 보내고, 판정은 `masks[r] & keep`
정수 연산 하나로 끝난다(94~126배).

대가는 **파생 상태가 하나 늘었다는 것**이다. 마스크가 매트릭스와 어긋나면 변경이
있는 행이 숨겨지거나 상태바 개수가 틀린다 — 병합 툴에서는 '변경을 못 보고 지나침'을
뜻한다. 그래서 매트릭스 변형은 _set_matrix_cell() 한 곳으로 묶었고, 아래 테스트가
그 계약을 고정한다.
"""
import random

import pytest
from PyQt5.QtWidgets import QApplication

from excelmerge.constants import (
    STATUS_SAME, STATUS_ADDED, STATUS_MODIFIED, DIR_A2B,
)
from excelmerge.diff_engine import (
    count_changed, count_changed_masked, keep_mask, row_change_masks,
)

STATUSES = [STATUS_SAME, STATUS_ADDED, STATUS_MODIFIED]


def _naive_row_changed(row, excl):
    """마스크 도입 전 판정 — 정본 오라클."""
    return any(st != STATUS_SAME for c, (st, *_) in enumerate(row) if c not in excl)


# ── 순수 함수: property 테스트 ────────────────────────────────────────────────

@pytest.mark.parametrize("seed", range(25))
def test_masks_match_naive_scan(seed):
    """무작위 매트릭스 x 무작위 제외 열에서 마스크 판정이 기존 스캔과 같아야 한다."""
    rnd = random.Random(seed)
    rows, cols = rnd.randint(1, 40), rnd.randint(1, 40)
    matrix = [[(rnd.choice(STATUSES), "a", "b") for _ in range(cols)]
              for _ in range(rows)]
    excl = {c for c in range(cols) if rnd.random() < 0.3}

    masks = row_change_masks(matrix)
    keep = keep_mask(cols, excl)
    assert len(masks) == rows

    for r, row in enumerate(matrix):
        assert bool(masks[r] & keep) == _naive_row_changed(row, excl), \
            f"seed={seed} r={r} excl={sorted(excl)}"
    assert count_changed_masked(masks, keep) == count_changed(matrix, excl), \
        f"seed={seed} excl={sorted(excl)}"


def test_masks_empty_matrix():
    assert row_change_masks([]) == []
    assert count_changed_masked([], keep_mask(0)) == 0


def test_keep_mask_ignores_out_of_range_columns():
    """범위 밖 제외 열 지정이 마스크를 오염시키면 안 된다."""
    assert keep_mask(3, {5, -1}) == 0b111
    assert keep_mask(3, {1}) == 0b101
    assert keep_mask(3, None) == 0b111


def test_mask_bit_positions_are_column_indices():
    matrix = [[(STATUS_SAME, "", ""), (STATUS_MODIFIED, "x", "y"),
               (STATUS_SAME, "", ""), (STATUS_ADDED, "z", "")]]
    assert row_change_masks(matrix) == [0b1010]


# ── 뷰 통합: 저장 확정이 마스크를 함께 갱신하는지 ─────────────────────────────

HEADER = ["ID", "V", "W"]
A_DATA = [HEADER, ["1", "x", "s"], ["2", "s2", "s3"]]
B_DATA = [HEADER, ["1", "y", "s"], ["2", "s2", "s3"]]
CELL = (1, 1)   # display 행 1(키 "1") 의 V 열 = ("modified", "x", "y")


def _wait_diff(dv, timeout_ms=5000):
    w = getattr(dv, "_diff_worker", None)
    if w is not None:
        w.wait(timeout_ms)
    QApplication.instance().processEvents()


@pytest.fixture
def dv(qapp, monkeypatch):
    """V 열 한 칸만 다른 활성 DiffView. 저장 완료 모달은 막는다."""
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    win = MainWindow()
    win.show()
    try:
        view = win.tabs.currentWidget()
        view.panel_a.set_path("a.xlsx")
        view.panel_b.set_path("b.xlsx")
        view._key_col, view._key_row = 0, 0
        for p in (view.panel_a, view.panel_b):
            p.table.set_key_col(0)
            p.table.set_key_row(0)
        view._on_loaded(A_DATA, B_DATA)
        _wait_diff(view)
        assert view._diff_matrix[CELL[0]][CELL[1]][0] == STATUS_MODIFIED
        yield view
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def test_worker_supplies_masks(dv):
    """DiffWorker 결과에 마스크가 매트릭스와 같은 길이로 실려 온다."""
    assert len(dv._row_masks) == len(dv._diff_matrix)
    assert dv._row_masks == row_change_masks(dv._diff_matrix)


def test_save_confirm_keeps_masks_in_sync(dv):
    """저장 확정(status -> same)이 마스크를 함께 갱신해야 한다.

    _set_matrix_cell 을 우회해 _diff_matrix 에 직접 대입하면 마스크가 낡아
    상태바 변경 셀 수와 Alt+화살표 탐색이 이미 병합된 셀을 계속 가리킨다.
    """
    assert dv._count_changed() == 1, dv._count_changed()
    assert CELL in set(dv._iter_changed_cells())

    dv._staged[CELL] = DIR_A2B
    dv._saving_side = "b"          # A->B staged 를 확정하는 쪽
    dv._on_staged_saved(1)
    QApplication.instance().processEvents()

    assert dv._diff_matrix[CELL[0]][CELL[1]] == (STATUS_SAME, "x", "x")
    assert dv._row_masks == row_change_masks(dv._diff_matrix), \
        "마스크가 매트릭스와 갈라졌다 — _set_matrix_cell 을 우회했다"
    assert dv._count_changed() == 0, \
        f"상태바 변경 셀 수가 낡았다: {dv._count_changed()}"
    assert CELL not in set(dv._iter_changed_cells()), \
        "이미 병합된 셀이 변경점 이동 대상에 남았다"


def test_row_masks_checked_rebuilds_on_length_mismatch(dv):
    """마스크 길이가 어긋나면(있어서는 안 되는 상태) 재구축해 맞는 답을 낸다."""
    dv._row_masks = []                       # 강제로 갈라놓기
    assert dv._row_masks_checked() == row_change_masks(dv._diff_matrix)
    assert dv._count_changed() == 1


def test_masks_cleared_with_matrix(dv):
    """매트릭스를 비우는 경로는 마스크도 함께 비운다."""
    dv._reset_compare_state()
    assert dv._diff_matrix == [] and dv._row_masks == []
