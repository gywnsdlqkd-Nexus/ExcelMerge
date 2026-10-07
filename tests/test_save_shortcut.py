# -*- coding: utf-8 -*-
"""Ctrl+S — 병합 준비된 쪽을 저장한다.

어느 파일을 쓸지는 **준비된 셀의 방향**이 정한다(B→A 준비면 A 파일, A→B 면 B 파일).
저장은 파일을 실제로 바꾸는 일이므로, 이 단축키가 **어느 쪽도 임의로 고르지 않는다**는
것을 특히 고정한다 — 양쪽에 준비가 걸려 있으면 아무 쪽도 쓰지 않는다.

전부 **실제 키 입력**(QTest.keyClick)으로 돈다. 핸들러를 직접 부르면 단축키 등록과
컨텍스트가 빠지는데, 그게 바로 Ctrl+C 가 오래 죽어 있던 이유였다.
"""
import pytest
from PyQt5.QtCore import Qt
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QApplication

from excelmerge.constants import DIR_A2B, DIR_B2A

A_DATA = [["TID", "NAME"], ["k1", "칼"], ["k2", "방패"]]
B_DATA = [["TID", "NAME"], ["k1", "검"], ["k2", "방패"]]
CELL = (1, 1)          # 값이 다른 칸


@pytest.fixture
def view(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    win = MainWindow()
    win.show()
    try:
        v = win.tabs.currentWidget()
        v.panel_a.set_path(str(tmp_path / "a.xlsx"))
        v.panel_b.set_path(str(tmp_path / "b.xlsx"))
        v._on_loaded(A_DATA, B_DATA)
        w = getattr(v, "_diff_worker", None)
        if w is not None:
            w.wait(8000)
        for _ in range(20):
            qapp.processEvents()
        yield v
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def _press_save(view):
    """실제 Ctrl+S. Ctrl 을 반드시 떼 준다 — 전역 수정자 상태가 눌린 채 남으면
    뒤따르는 테스트의 헤더 클릭이 Ctrl+클릭으로 오인된다(test_copy_cells 의 교훈)."""
    w = view.panel_a.table
    w.setFocus()
    QApplication.instance().processEvents()
    QTest.keyClick(w, Qt.Key_S, Qt.ControlModifier)
    QTest.keyRelease(w, Qt.Key_Control, Qt.NoModifier)
    QApplication.instance().processEvents()


def _record_saves(view, monkeypatch):
    """_save_staged 를 가로채 '어느 쪽을 저장하려 했는가'만 받아 둔다(파일은 안 쓴다)."""
    saved = []
    monkeypatch.setattr(view, "_save_staged", saved.append)
    return saved


# ── 준비된 방향이 저장할 쪽을 정한다 ─────────────────────────────────────────

def test_ctrl_s_saves_the_side_the_cells_are_staged_for_a(view, monkeypatch):
    """B→A 로 준비했으면 A 파일을 쓴다."""
    saved = _record_saves(view, monkeypatch)
    view._staged[CELL] = DIR_B2A
    _press_save(view)
    assert saved == ["a"]


def test_ctrl_s_saves_the_side_the_cells_are_staged_for_b(view, monkeypatch):
    """A→B 로 준비했으면 B 파일을 쓴다."""
    saved = _record_saves(view, monkeypatch)
    view._staged[CELL] = DIR_A2B
    _press_save(view)
    assert saved == ["b"]


def test_ctrl_s_agrees_with_the_save_button(view, monkeypatch):
    """단축키가 저장하는 쪽은 버튼이 켜져 있는 쪽과 같아야 한다.

    둘이 따로 판정하면 '버튼은 꺼져 있는데 Ctrl+S 는 저장한다' 가 생긴다.
    """
    saved = _record_saves(view, monkeypatch)
    view._staged[CELL] = DIR_A2B
    view._set_save_btn_state()
    assert view.panel_b.save_btn.isEnabled()
    assert not view.panel_a.save_btn.isEnabled()
    _press_save(view)
    assert saved == ["b"]


# ── 쓰지 말아야 할 때 ────────────────────────────────────────────────────────

def test_ctrl_s_does_nothing_without_staged_cells(view, monkeypatch):
    """준비된 셀이 없으면 아무 파일도 건드리지 않는다."""
    saved = _record_saves(view, monkeypatch)
    assert not view._staged
    _press_save(view)
    assert saved == []


def test_ctrl_s_refuses_to_pick_a_side_when_both_are_staged(view, monkeypatch):
    """양쪽에 준비가 걸려 있으면 아무 쪽도 쓰지 않는다 — 단축키가 고를 일이 아니다.

    둘 다 쓰면 Ctrl+S 한 번이 파일 두 개를 바꾼다. 조용히 넘어가지 말고 알린다.
    """
    saved = _record_saves(view, monkeypatch)
    view._staged[(1, 1)] = DIR_A2B
    view._staged[(2, 1)] = DIR_B2A
    seen = []
    view.status.messageChanged.connect(seen.append)
    _press_save(view)
    assert saved == []
    assert any("양쪽" in m for m in seen), seen


def test_ctrl_s_does_not_start_a_second_save_while_one_is_running(view, monkeypatch):
    """저장이 도는 중에 또 누르면 워커가 덮어써져 앞 저장을 놓친다.

    가짜 워커는 **진짜 QThread** 여야 한다 — DiffView.shutdown 이 blockSignals/wait 를
    부르는데, 그게 없는 객체를 꽂아 두면 closeEvent 안에서 AttributeError 가 나고
    프로세스가 트레이스백 없이 즉사한다(종료 코드만 남아 원인을 못 찾는다).
    """
    from PyQt5.QtCore import QThread

    class _Running(QThread):
        def isRunning(self):
            return True

    saved = _record_saves(view, monkeypatch)
    view._staged[CELL] = DIR_A2B
    view._staged_merge_worker = _Running()
    try:
        _press_save(view)
        assert saved == []
    finally:
        view._staged_merge_worker = None   # 정리는 뒷정리에 떠넘기지 않는다


def test_ctrl_s_after_a_finished_save_does_not_blow_up(view, monkeypatch):
    """끝난 저장 워커는 C++ 객체가 사라진다 — 그 참조로 isRunning() 을 부르면 터진다.

    v220 실기에서 **저장을 한 번 한 뒤** Ctrl+S 를 다시 누르자 오류 창이 떴다:
    'wrapped C/C++ object of type StagedMergeWorker has been deleted'. 유닛 테스트는
    진짜 저장을 돌리지 않아 워커가 죽은 상태를 한 번도 만들지 않았고, 그래서 못 잡았다.

    sip.delete 로 C++ 쪽만 없애면 deleteLater 가 끝난 그 상태를 그대로 재현한다.
    """
    from PyQt5 import sip
    from PyQt5.QtCore import QThread

    saved = _record_saves(view, monkeypatch)
    dead = QThread()
    sip.delete(dead)                       # deleteLater 가 끝난 뒤와 같은 상태
    view._staged_merge_worker = dead
    view._staged[CELL] = DIR_A2B

    _press_save(view)

    assert saved == ["b"], "죽은 워커 참조 때문에 저장이 막혔다"
    assert view._staged_merge_worker is None, "죽은 참조를 끊지 않았다"


# ── 안내는 _save_staged 가 한다 — 단축키가 미리 걸러 삼키지 않는다 ───────────

def test_ctrl_s_on_a_non_excel_file_still_explains_why(view, monkeypatch, tmp_path):
    """비엑셀 파일은 저장을 못 한다 — 그 안내가 단축키 경로에서도 떠야 한다.

    버튼의 활성 상태로 미리 걸러 내면 Ctrl+S 가 아무 반응 없는 것처럼 보인다.
    """
    from excelmerge import diff_view as dv_mod
    view.panel_b.set_path(str(tmp_path / "b.json"))
    seen = []
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda parent, title, text, *a, **k:
                                     seen.append((title, text))))
    view._staged[CELL] = DIR_A2B
    _press_save(view)
    assert seen, "비엑셀 안내가 뜨지 않았다"
    assert "저장" in seen[-1][0]
