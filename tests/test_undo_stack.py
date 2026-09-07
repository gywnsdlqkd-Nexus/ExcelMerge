# -*- coding: utf-8 -*-
"""Ctrl+Z 취소 스택 회귀 — 항목은 '지금 매트릭스의 좌표'일 때만 유효하다.

_undo_stack 은 (r, c) 좌표를 담는다. 그런데 새 비교 / 새로고침 / 시트 전환 / 키 열·행
변경은 매트릭스를 통째로 재계산하고, 저장 확정은 staged 를 소비한다 — 그 뒤 남아 있는
항목은 **다른 셀을 가리킨다.** 예전엔 스택을 비우는 곳이 한 군데도 없어서 Ctrl+Z 가
사용자가 건드린 적 없는 셀의 준비/표시값을 조용히 지울 수 있었다.

또 병합 준비 '취소'는 스택에 쌓이지 않아 되살릴 수 없었다(stage 만 되돌아가는 비대칭).

여기 테스트는 그 두 계약을 고정한다.
"""
import pytest
from PyQt5.QtWidgets import QApplication

from excelmerge.constants import STATUS_MODIFIED, DIR_A2B, DIR_B2A

HEADER = ["ID", "V", "W"]
A_DATA = [HEADER, ["1", "x", "p"], ["2", "m", "q"]]
B_DATA = [HEADER, ["1", "y", "p"], ["2", "n", "q"]]
C1 = (1, 1)   # display 행 1(키 "1") V 열 — ("modified", "x", "y")
C2 = (2, 1)   # display 행 2(키 "2") V 열 — ("modified", "m", "n")


def _wait_diff(dv, timeout_ms=5000):
    w = getattr(dv, "_diff_worker", None)
    if w is not None:
        w.wait(timeout_ms)
    QApplication.instance().processEvents()


@pytest.fixture
def dv(qapp, monkeypatch):
    """V 열 두 칸이 다른 활성 DiffView. 모달(저장 완료/알림)은 막는다."""
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
        assert view._diff_matrix[C1[0]][C1[1]][0] == STATUS_MODIFIED
        assert view._diff_matrix[C2[0]][C2[1]][0] == STATUS_MODIFIED
        yield view
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def _stage(dv, cell, direction=DIR_A2B):
    """_stage_selected 와 같은 상태 변화 — 선택 UI 를 거치지 않는 직접 경로."""
    dv._undo_stack.append(("stage", [cell], direction))
    dv._staged[cell] = direction
    display = dv._diff_matrix[cell[0]][cell[1]][1 if direction == DIR_A2B else 2]
    dv.panel_a._staged_display[cell] = display
    dv.panel_b._staged_display[cell] = display


# ── 스택 정리 계약 ────────────────────────────────────────────────────────────

def test_stack_cleared_on_recompute(dv):
    """키 변경은 행 순서를 바꾼다 — 옛 좌표 항목이 남으면 안 된다."""
    _stage(dv, C1)
    assert dv._undo_stack
    dv._recompute_diff()
    _wait_diff(dv)
    assert dv._undo_stack == []


def test_stack_cleared_on_new_compare(dv):
    _stage(dv, C1)
    dv._run_compare()
    assert dv._undo_stack == []


def test_stack_cleared_on_reset(dv):
    _stage(dv, C1)
    dv._reset_compare_state()
    assert dv._undo_stack == []


def test_stack_cleared_on_save(dv):
    """확정 저장은 되돌릴 수 없다 — 항목을 남기면 Ctrl+Z 가 무반응으로 소모된다."""
    _stage(dv, C1)
    dv._saving_side = "b"
    dv._on_staged_saved(1)
    QApplication.instance().processEvents()
    assert dv._undo_stack == []


def test_undo_after_save_leaves_merged_cell_display_intact(dv):
    """저장 뒤 남은 옛 항목이 **병합 완료 셀의 셀값란 표시값**을 지우면 안 된다.

    _staged_display 는 저장 뒤에도 남아 merged 셀의 셀값란(병합된 값)을 공급한다
    (panels._refresh_cell_edit_from_selection). 스택을 비우지 않으면 나중의 Ctrl+Z 가
    그 항목을 pop 하면서 이 표시값까지 함께 날린다 — 사용자가 건드린 적 없는 셀이다.
    """
    _stage(dv, C1)
    dv._saving_side = "b"
    dv._on_staged_saved(1)
    QApplication.instance().processEvents()
    assert C1 in dv._merged_cells
    assert dv.panel_a._staged_display.get(C1) == "x"

    _stage(dv, C2)          # 저장과 무관한 다른 셀을 새로 준비
    dv._undo()              # 방금 준비한 C2 만 취소돼야 한다
    assert C2 not in dv._staged
    dv._undo()              # 스택이 비어 있어야 하므로 아무 일도 없어야 한다

    assert dv.panel_a._staged_display.get(C1) == "x", \
        "저장된 셀의 셀값란 표시값이 옛 취소 항목에 의해 지워졌다"
    assert dv.panel_b._staged_display.get(C1) == "x"


def test_undo_on_empty_stack_is_reported(dv):
    """무음 no-op 이면 사용자가 '먹혔는지' 알 수 없다 — 상태바에 남긴다."""
    dv._undo_stack.clear()
    dv._undo()
    assert "되돌릴" in dv.status.currentMessage()


# ── 취소(unstage)도 되돌릴 수 있어야 한다 ─────────────────────────────────────

def test_unstage_is_undoable(dv):
    """준비 → 취소 → Ctrl+Z 는 취소를 되돌린다(그 이전 준비를 지우는 게 아니다)."""
    _stage(dv, C1, DIR_A2B)
    _stage(dv, C2, DIR_B2A)
    dv.panel_a.table._select_range(C2[0], C2[1], C2[0], C2[1])
    dv._unstage_selected()
    assert C2 not in dv._staged and C1 in dv._staged

    dv._undo()
    assert dv._staged.get(C2) == DIR_B2A, "취소가 되돌아가지 않았다"
    assert dv._staged.get(C1) == DIR_A2B, "관계없는 준비가 함께 사라졌다"
    # 셀값란 표시값(병합될 값)도 함께 복원 — b_to_a 이므로 B 쪽 값.
    assert dv.panel_a._staged_display.get(C2) == "n"
    assert dv.panel_b._staged_display.get(C2) == "n"


def test_undo_of_undone_unstage_removes_again(dv):
    """복원한 뒤 한 번 더 Ctrl+Z 하면 그 앞 단계(준비)로 계속 거슬러 간다."""
    _stage(dv, C1)
    dv.panel_a.table._select_range(C1[0], C1[1], C1[0], C1[1])
    dv._unstage_selected()
    dv._undo()                      # 취소를 되돌림 → 다시 staged
    assert C1 in dv._staged
    dv._undo()                      # 그 앞의 stage 를 되돌림
    assert C1 not in dv._staged
    assert dv._undo_stack == []


# ── _merged_cells 공유 참조 ───────────────────────────────────────────────────

def test_merged_stays_shared_with_models_during_recompute(dv):
    """_merged_cells 는 두 모델과 **같은 객체**여야 한다 — 비동기 재계산 창 포함.

    populate 가 이 set 을 참조로 넘기므로(set_diff_data), 리셋에서 새 객체로
    갈아치우면 다음 populate 까지 뷰와 모델이 서로 다른 set 을 본다.
    _recompute_diff 는 워커를 띄우고 곧장 반환하므로 그 창이 실제로 열린다.
    """
    model_a = dv.panel_a.table.model()
    model_b = dv.panel_b.table.model()
    assert dv._merged_cells is model_a._merged
    assert dv._merged_cells is model_b._merged

    dv._recompute_diff()            # populate 이전 — 여기서 단언해야 의미가 있다
    assert dv._merged_cells is model_a._merged, \
        "_merged_cells 를 새 set 으로 갈아치워 모델과 갈라졌다"
    assert dv._merged_cells is model_b._merged
    _wait_diff(dv)
