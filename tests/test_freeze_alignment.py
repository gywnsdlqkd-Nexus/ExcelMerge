# -*- coding: utf-8 -*-
"""틀 고정 키 열이 본체와 **같은 줄에** 그려지는가 — 모든 스크롤 위치에서.

실기에서 나왔다(36↔40 빌드 Data_MailBox_CS.xlsx). 아래로 스크롤하면 키 열(MailBoxID)만
반 행쯤 내려가 그려져, 행 번호 252 자리에 키 251 이 보였다. 행 번호와 본문(#Desc 등)은
서로 맞아서 "키 값만 아래로 밀린다" 로 보였다.

왜 오래 안 보였나. 고정 열은 **별도 위젯**이고, 세로 동기화는 본체의 픽셀 오프셋을
그 위젯의 스크롤바에 복사하는 방식이다. 그 오프셋은 **양쪽이 모두 제자리에 있을 때만**
같은 줄을 뜻한다. 밴드의 y 는 reposition() 이 `fr + hdrH + fh` 로 정하는데 그건 창 크기가
바뀔 때만 돈다 — 그 뒤 헤더 높이나 고정 행 높이가 바뀌면 밴드만 어긋난 채 남고,
오프셋은 서로 같으니 동기화는 "맞다"고 본다. 그래서 **스크롤을 아무리 해도 안 고쳐지고**
창 크기를 바꿔야 돌아왔다.

그러니 스크롤바 값이 아니라 **실제로 그려지는 화면 위치**를 재야 잡힌다. 아래 테스트는
전부 그렇게 한다.
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from PyQt5.QtCore import QPoint
from PyQt5.QtWidgets import QApplication


def _wait_diff(win, timeout_ms=8000):
    w = getattr(win, "_diff_worker", None)
    if w is not None:
        w.wait(timeout_ms)
    app = QApplication.instance()
    if app is not None:
        app.processEvents()


@pytest.fixture
def frozen(qapp):
    """키 앵커 A1, 스크롤이 생길 만큼 긴 표. 틀 고정 ON."""
    from excelmerge.main_window import MainWindow
    win = MainWindow()
    win.resize(1200, 700)
    win.show()
    try:
        dv = win.tabs.currentWidget()
        dv.panel_a.set_path("a.xlsx"); dv.panel_b.set_path("b.xlsx")
        a = [["ID", "V", "W", "U"]] + [[str(i), "a%d" % i, "p", "q"] for i in range(200)]
        b = [row[:] for row in a]
        b[7][1] = "CH"
        dv._on_loaded(a, b); _wait_diff(win)
        qapp.processEvents()
        dv._freeze["a"].refresh(); qapp.processEvents()
        assert dv._freeze["a"].active, "전제: 틀 고정이 켜져 있어야 한다"
        dv.diff_only_btn.setChecked(False)      # 전체 행 — 스크롤 범위 확보
        qapp.processEvents()
        yield dv
    finally:
        win.close(); win.deleteLater(); qapp.processEvents()


def _row_y(view, r, col):
    """그 행이 **화면에서** 어느 높이에 그려지는지(전역 y). 스크롤바 값이 아니라 결과를 본다."""
    if view.isRowHidden(r):
        return None
    rect = view.visualRect(view.model().index(r, col))
    if rect.isNull() or rect.height() == 0:
        return None
    return view.viewport().mapToGlobal(rect.topLeft()).y()


def _drift(dv, side="a"):
    """본체와 고정 밴드가 같은 행을 다른 높이에 그리면 그 목록을 돌려준다."""
    fz = dv._freeze[side]
    host, left = fz.host, fz.left
    body_col = dv._key_col + 1
    out = []
    for r in range(host.model().rowCount()):
        hy = _row_y(host, r, body_col)
        ly = _row_y(left, r, 0)
        if hy is None or ly is None:
            continue
        if hy != ly:
            out.append((r, hy, ly, ly - hy))
    return out


# ── 모든 스크롤 위치에서 ─────────────────────────────────────────────────────

def test_aligned_at_every_scroll_position(frozen, qapp):
    """맨 위부터 맨 아래까지 한 칸씩 — 어디서도 어긋나면 안 된다."""
    dv = frozen
    sb = dv.panel_a.table.verticalScrollBar()
    assert sb.maximum() > 0, "전제: 세로 스크롤이 생겨야 한다"
    bad = []
    for val in range(0, sb.maximum() + 1):
        sb.setValue(val)
        qapp.processEvents()
        d = _drift(dv)
        if d:
            bad.append((val, d[0]))
    assert not bad, f"어긋난 스크롤 위치 {len(bad)}곳: {bad[:5]}"


def test_aligned_at_the_very_bottom(frozen, qapp):
    """맨 아래가 특히 위험하다 — 본체는 ScrollPerItem, 밴드는 ScrollPerPixel 이라
    끝에서 클램프가 서로 다르게 걸릴 수 있다."""
    dv = frozen
    sb = dv.panel_a.table.verticalScrollBar()
    sb.setValue(sb.maximum())
    qapp.processEvents()
    assert _drift(dv) == []


def test_both_panels_stay_aligned(frozen, qapp):
    """A/B 는 스크롤이 묶여 있다 — 한쪽만 맞고 끝나면 안 된다."""
    dv = frozen
    sb = dv.panel_a.table.verticalScrollBar()
    sb.setValue(sb.maximum() // 2)
    qapp.processEvents()
    assert _drift(dv, "a") == []
    assert _drift(dv, "b") == []


# ── 어긋난 상태에서 스스로 돌아오는가 ────────────────────────────────────────

def test_a_misplaced_band_heals_on_the_next_scroll(frozen, qapp):
    """실기 증상의 핵심 — 한 번 어긋나면 **스크롤로는 안 고쳐졌다.**

    밴드를 강제로 몇 픽셀 내려 그 상태를 만든 뒤, 스크롤 한 번에 제자리로 돌아오는지
    본다. 고치기 전에는 창 크기를 바꿔야만 돌아왔다.
    """
    dv = frozen
    fz = dv._freeze["a"]
    sb = dv.panel_a.table.verticalScrollBar()
    sb.setValue(max(1, sb.maximum() // 3))
    qapp.processEvents()
    assert _drift(dv) == [], "전제: 흔들기 전에는 맞아야 한다"

    fz.left.move(fz.left.x(), fz.left.y() + 11)   # 반 행쯤 내려 놓는다
    qapp.processEvents()
    assert _drift(dv), "전제: 흔든 뒤에는 어긋나야 한다"

    sb.setValue(sb.value() + 1)                   # 스크롤 한 번
    qapp.processEvents()
    assert _drift(dv) == [], "스크롤해도 어긋남이 남아 있다"


def test_healing_does_not_loop(frozen, qapp):
    """보정이 reposition 을 부르고 그게 다시 동기화를 부른다 — 재진입을 막아야 한다."""
    dv = frozen
    fz = dv._freeze["a"]
    calls = []
    orig = fz.reposition

    def counted(*a, **k):
        calls.append(1)
        return orig(*a, **k)

    fz.reposition = counted
    try:
        fz.left.move(fz.left.x(), fz.left.y() + 11)
        qapp.processEvents()
        dv.panel_a.table.verticalScrollBar().setValue(2)
        qapp.processEvents()
    finally:
        fz.reposition = orig
    assert len(calls) <= 2, f"reposition 이 {len(calls)}번 — 재귀로 돈다"
