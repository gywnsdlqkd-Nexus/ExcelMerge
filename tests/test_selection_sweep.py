# -*- coding: utf-8 -*-
"""틀 고정 경계를 넘나드는 선택 — 전수 검사(437조합). 화면 없이 돈다.

키 열/행은 본체에서 숨긴 채 오버레이 세 개(키 열 · 키 행 · 모서리)가 그린다. 예전엔 네 뷰가
각자 Qt 기본 처리로 선택했고, 경계를 넘는 경우만 수정키별로 땜질해 와서, 제보가 하나씩
들어왔다(v228 Shift+클릭 → 그다음 Ctrl+드래그). 하나씩 고치지 않고 **조합 전체**를 돌린다.

  셀     시작 뷰 × 끝 뷰(본체 · 키 열 · 키 행 · 모서리) × 8가지 동작
         (Shift+클릭 · Ctrl+클릭 · 드래그 · Ctrl+드래그 · Shift+드래그 · Ctrl+Shift+클릭 ·
          Ctrl 로 잡은 영역 + Shift+클릭 · Shift 범위 위 Ctrl+Shift+클릭) — 스크롤 없음 / 세로 스크롤 / 가로 스크롤
  헤더   A열·1행 머리(모서리)와 본체 머리 사이 × 6가지 동작
  키보드 Shift+방향키로 경계 넘기, 마우스로 잡은 기준을 키보드가 이어받기
  우클릭 선택 안 → 유지, 밖 → 그 칸

기대값은 엑셀 규칙이다. 선택이 B 패널에 똑같이 미러되는지도 함께 본다.
고치기 전(v228): 389조합 중 168개 어긋남 — 셀 153 · 헤더 15.

이벤트 생성 주의(둘 다 하네스에서 실제로 틀렸다):
  · 마우스 이벤트는 로컬·창·전역 좌표를 모두 채운다. 위치 하나만 주면 전역 좌표가 로컬
    값으로 채워져, 전역 좌표로 경계를 판정하는 코드가 엉뚱한 칸을 가리킨다.
  · 헤더 코드는 앱 전역 수정키 상태(keyboardModifiers)를 읽는다. 그 상태는 '실제 입력'
    표시가 붙은 이벤트(QTest)만 갱신하므로 헤더는 QTest 로 누르고 뗀다.
"""
import itertools
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from PyQt5.QtCore import QEvent, QPoint, QPointF, Qt
from PyQt5.QtGui import QMouseEvent
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QApplication

NO, CT, SH = Qt.NoModifier, Qt.ControlModifier, Qt.ShiftModifier
R = []   # (영역, 이름, ok, 설명)


def build(nrows, ncols):
    from excelmerge.main_window import MainWindow
    win = MainWindow(); win.resize(1200, 700); win.show()
    dv = win.tabs.currentWidget(); dv.panel_a.set_path("a.xlsx"); dv.panel_b.set_path("b.xlsx")
    a = [["ID"] + [f"C{c}" for c in range(1, ncols)]] + \
        [[str(r)] + [f"v{r}_{c}" for c in range(1, ncols)] for r in range(nrows)]
    b = [x[:] for x in a]; b[7][1] = "CH"
    dv._on_loaded(a, b); w = getattr(dv, "_diff_worker", None); w and w.wait(8000); QApplication.instance().processEvents()
    dv._freeze["a"].refresh(); dv.diff_only_btn.setChecked(False); QApplication.instance().processEvents()
    return win, dv


class H:
    def __init__(self, dv):
        self.dv, self.t, self.fc = dv, dv.panel_a.table, dv._freeze["a"]
        self.V = {"B": self.t, "L": self.fc.left, "T": self.fc.top, "C": self.fc.corner}

    def vname(self, cell):
        r, c = cell
        fr, fcn = self.fc._n_rows, self.fc._n_cols
        if r < fr and c < fcn: return "C"
        if r < fr: return "T"
        if c < fcn: return "L"
        return "B"

    def send(self, vp, typ, lp, mods, button=Qt.LeftButton, buttons=Qt.LeftButton):
        lp = QPoint(lp)
        ev = QMouseEvent(typ, QPointF(lp), QPointF(vp.mapTo(vp.window(), lp)),
                         QPointF(vp.mapToGlobal(lp)), button, buttons, mods)
        QApplication.sendEvent(vp, ev); QApplication.instance().processEvents()

    def cpos(self, cell):
        v = self.V[self.vname(cell)]
        rect = v.visualRect(self.t._model.index(*cell))
        assert not rect.isEmpty(), ("안 보임", cell)
        return v, rect.center()

    def click(self, cell, mods=NO, button=Qt.LeftButton):
        v, p = self.cpos(cell)
        self.send(v.viewport(), QEvent.MouseButtonPress, p, mods, button, button)
        self.send(v.viewport(), QEvent.MouseButtonRelease, p, mods, button, Qt.NoButton)

    def drag_pts(self, vp, g0, g1, mods):
        self.send(vp, QEvent.MouseButtonPress, vp.mapFromGlobal(g0), mods)
        for i in range(1, 6):
            self.send(vp, QEvent.MouseMove, vp.mapFromGlobal(g0 + (g1 - g0) * i / 5), mods, button=Qt.NoButton)
        self.send(vp, QEvent.MouseButtonRelease, vp.mapFromGlobal(g1), mods, buttons=Qt.NoButton)

    def drag(self, s, e, mods=NO):
        v0, p0 = self.cpos(s); v1, p1 = self.cpos(e)
        self.drag_pts(v0.viewport(), v0.viewport().mapToGlobal(p0), v1.viewport().mapToGlobal(p1), mods)

    # 헤더: 열 c 는 고정이면 corner 가로 헤더, 아니면 본체 가로 헤더
    def hdr(self, axis, i):
        if axis == "col":
            h = self.fc.corner.horizontalHeader() if i < self.fc._n_cols else self.t.horizontalHeader()
            x = h.sectionViewportPosition(i) + h.sectionSize(i) // 2
            assert h.isVisible() and 0 <= x < h.viewport().width(), ("열 머리 안 보임", i)
            return h, QPoint(x, h.viewport().height() // 2)
        h = self.fc.corner.verticalHeader() if i < self.fc._n_rows else self.t.verticalHeader()
        y = h.sectionViewportPosition(i) + h.sectionSize(i) // 2
        assert h.isVisible() and 0 <= y < h.viewport().height(), ("행 머리 안 보임", i)
        return h, QPoint(h.viewport().width() // 2, y)

    # 헤더 코드는 이벤트가 아니라 앱 전역 수정키 상태(keyboardModifiers)를 읽는다. 그 상태는
    # '실제 입력' 표시(spontaneous)가 붙은 이벤트만 갱신하므로 QTest 로 누르고 뗀다.
    def hclick(self, axis, i, mods=NO):
        h, p = self.hdr(axis, i)
        QTest.mousePress(h.viewport(), Qt.LeftButton, mods, p); QApplication.instance().processEvents()
        QTest.mouseRelease(h.viewport(), Qt.LeftButton, mods, p); QApplication.instance().processEvents()

    def hdrag(self, axis, i, j, mods=NO):
        h0, p0 = self.hdr(axis, i); h1, p1 = self.hdr(axis, j)
        vp = h0.viewport(); g0 = vp.mapToGlobal(p0); g1 = h1.viewport().mapToGlobal(p1)
        QTest.mousePress(vp, Qt.LeftButton, mods, p0); QApplication.instance().processEvents()
        for k in range(1, 6):
            self.send(vp, QEvent.MouseMove, vp.mapFromGlobal(g0 + (g1 - g0) * k / 5), mods, button=Qt.NoButton)
        QTest.mouseRelease(vp, Qt.LeftButton, mods, vp.mapFromGlobal(g1)); QApplication.instance().processEvents()

    def key(self, k, mods=NO):
        QTest.keyClick(self.t, k, mods)
        if mods & CT: QTest.keyRelease(self.t, Qt.Key_Control, NO)
        if mods & SH: QTest.keyRelease(self.t, Qt.Key_Shift, NO)
        QApplication.instance().processEvents()

    def cells(self, table=None):
        out = set()
        for rg in (table or self.t).selectionModel().selection():
            for r in range(rg.top(), rg.bottom() + 1):
                for c in range(rg.left(), rg.right() + 1):
                    out.add((r, c))
        return out

    def cur(self):
        i = self.t.currentIndex(); return (i.row(), i.column())

    def reset(self, p):
        self.t.clearSelection(); QApplication.instance().processEvents(); self.click(p)


def rect(x, y):
    return {(r, c) for r in range(min(x[0], y[0]), max(x[0], y[0]) + 1)
            for c in range(min(x[1], y[1]), max(x[1], y[1]) + 1)}


def check(area, name, h, exp, exp_cur=None):
    got = h.cells()
    ok = got == exp and (exp_cur is None or h.cur() == exp_cur) and h.cells(h.dv.panel_b.table) == got
    why = ""
    if got != exp:
        why = f"빠짐 {sorted(exp - got)[:5]}({len(exp - got)}) 더함 {sorted(got - exp)[:5]}({len(got - exp)})"
    elif exp_cur is not None and h.cur() != exp_cur:
        why = f"현재셀 {h.cur()}≠{exp_cur}"
    elif h.cells(h.dv.panel_b.table) != got:
        why = "B 미러 다름"
    R.append((area, name, ok, why))


def cell_matrix(area, h, CELLS, P):
    fams = {
        "Shift+클릭": lambda s, e: (h.reset(P), h.click(s), h.click(e, SH), rect(s, e))[-1],
        "Ctrl+클릭": lambda s, e: (h.reset(P), h.click(s, CT), h.click(e, CT), {P} if s == e else {P, s, e})[-1],
        "드래그": lambda s, e: (h.reset(P), h.drag(s, e), rect(s, e))[-1],
        "Ctrl+드래그": lambda s, e: (h.reset(P), h.drag(s, e, CT), {P} | rect(s, e))[-1],
        "Shift+드래그": lambda s, e: (h.reset(P), h.click(s), h.drag(CELLS[h.vname(s)][1] if CELLS[h.vname(s)][1] != s else CELLS[h.vname(s)][0], e, SH), rect(s, e))[-1],
        "Ctrl+Shift+클릭": lambda s, e: (h.reset(P), h.click(s, CT), h.click(e, CT | SH), {P} | rect(s, e))[-1],
        "Ctrl영역+Shift+클릭": lambda s, e: (h.reset(P), h.click(s, CT), h.click(e, SH), {P} | rect(s, e))[-1],
        # Shift 로 넓게 잡은 범위 위의 Ctrl+Shift — 먼저 잡은 범위가 남아야 한다(엑셀·Qt 동일)
        "Shift범위+Ctrl+Shift+클릭": lambda s, e: (h.reset(P), h.click(s), h.click(P, SH), h.click(e, CT | SH),
                                              rect(s, P) | rect(s, e))[-1],
    }
    for fam, f in fams.items():
        for sv, ev in itertools.product("BLTC", repeat=2):
            s = CELLS[sv][0]; e = CELLS[ev][1] if sv == ev else CELLS[ev][0]
            exp = f(s, e)
            check(area, f"{fam} {sv}→{ev}", h, exp, e)


def header_matrix(h, P):
    fr_cols, fr_rows = h.fc._n_cols, h.fc._n_rows
    nrow, ncol = h.t.rowCount(), h.t.columnCount()   # 헤더 클릭은 표 끝 여분 행·열까지 그 줄 전체
    one = {}
    for axis, idxs in (("col", [0, 2, 3]), ("row", [0, 5, 9])):
        for i in idxs:
            h.reset(P); h.hclick(axis, i); one[(axis, i)] = h.cells()
    def span(axis, i, j):
        out = set()
        for k in range(min(i, j), max(i, j) + 1):
            out |= ({(r, k) for r in range(nrow)} if axis == "col" else {(k, c) for c in range(ncol)})
        return out
    # 단일 클릭이 그 줄 전체인지(이 앱의 헤더 클릭 규칙) 먼저 본다
    for (axis, i), got in one.items():
        R.append(("헤더", f"{axis} {i} 클릭 = 그 줄 전체", got == span(axis, i, i),
                  "" if got == span(axis, i, i) else f"{len(got)}칸"))
    for axis, (a, b2, c2) in (("col", (0, 2, 3)), ("row", (0, 5, 9))):
        for x, y in ((a, b2), (b2, a), (b2, c2)):
            tag = f"{axis} {x}→{y}"
            h.reset(P); h.hclick(axis, x); h.hclick(axis, y, SH); check("헤더", f"Shift+클릭 {tag}", h, span(axis, x, y))
            h.reset(P); h.hclick(axis, x, CT); h.hclick(axis, y, CT); check("헤더", f"Ctrl+클릭 {tag}", h, {P} | span(axis, x, x) | span(axis, y, y))
            h.reset(P); h.hdrag(axis, x, y); check("헤더", f"드래그 {tag}", h, span(axis, x, y))
            h.reset(P); h.hdrag(axis, x, y, CT); check("헤더", f"Ctrl+드래그 {tag}", h, {P} | span(axis, x, y))
            h.reset(P); h.hclick(axis, x); h.hdrag(axis, y, y, SH); check("헤더", f"Shift+드래그 {tag}", h, span(axis, x, y))
            h.reset(P); h.hclick(axis, x, CT); h.hclick(axis, y, CT | SH); check("헤더", f"Ctrl+Shift+클릭 {tag}", h, {P} | span(axis, x, y))


def keyboard(h):
    last = h.t._model.data_cols - 1
    cases = [
        ("키 열→오른쪽 Shift+→×2", (5, 0), [(Qt.Key_Right, SH)] * 2, rect((5, 0), (5, 2))),
        ("본체→키 열 Shift+←×2", (5, 2), [(Qt.Key_Left, SH)] * 2, rect((5, 0), (5, 2))),
        ("본체→키 행 Shift+↑×3", (3, 2), [(Qt.Key_Up, SH)] * 3, rect((0, 2), (3, 2))),
        ("키 행→아래 Shift+↓×2", (0, 2), [(Qt.Key_Down, SH)] * 2, rect((0, 2), (2, 2))),
        ("모서리→오른쪽아래 Shift+→↓", (0, 0), [(Qt.Key_Right, SH), (Qt.Key_Down, SH)], rect((0, 0), (1, 1))),
        ("키 열 Ctrl+Shift+→", (5, 0), [(Qt.Key_Right, CT | SH)], rect((5, 0), (5, last))),
    ]
    for name, start, keys, exp in cases:
        h.reset(start)
        for k, m in keys: h.key(k, m)
        check("키보드", name, h, exp)
    # 마우스로 잡은 기준을 키보드가 이어받는가
    h.reset((5, 0)); h.click((7, 2), SH); h.key(Qt.Key_Right, SH)
    check("키보드", "키 열 클릭→Shift+클릭→Shift+→", h, rect((5, 0), (7, 3)))


def rightclick(h):
    for name, inside, exp_keep in (("선택 안 키 열", (7, 0), True), ("선택 안 본체", (6, 1), True),
                                   ("선택 밖 키 열", (12, 0), False), ("선택 밖 본체", (12, 2), False)):
        h.reset((20, 1)); h.drag((5, 0), (9, 2)); before = h.cells()
        h.click(inside, button=Qt.RightButton)
        check("우클릭", name, h, before if exp_keep else {inside})


def visible_rows(t, lo=1):
    return [r for r in range(lo, t.rowCount()) if not t.isRowHidden(r)
            and 0 <= t.rowViewportPosition(r) < t.viewport().height() - 40]


def visible_cols(t, lo=1):
    return [c for c in range(lo, t.columnCount()) if not t.isColumnHidden(c)
            and 0 <= t.columnViewportPosition(c)
            and t.columnViewportPosition(c) + t.columnWidth(c) <= t.viewport().width()]




def _report():
    bad = [(area, name, why) for area, name, ok, why in R if not ok]
    total = len(R)
    R.clear()
    assert total, "검사한 조합이 없다"
    assert not bad, f"{len(bad)}/{total} 어긋남:\n" + "\n".join(f"  [{a}] {n}: {w}" for a, n, w in bad)


def _module_window(qapp, tmp_path_factory, name, nrows, ncols):
    """모듈 단위 창 — **스스로 APPDATA 를 격리한다.**

    conftest 의 격리(isolated_appdata)는 함수 단위라, 모듈 단위 픽스처는 그보다 먼저
    만들어져 **사용자의 실제 설정**(키 위치 등)을 읽는다. 실제로 그래서 키 열이 B 까지
    고정된 표가 만들어졌고, 결과가 사용자 환경에 따라 달라졌다.
    """
    with pytest.MonkeyPatch.context() as mp:
        mp.setenv("APPDATA", str(tmp_path_factory.mktemp(name)))
        win, dv = build(nrows, ncols)
        h = H(dv)
        assert (h.fc._n_rows, h.fc._n_cols) == (1, 1), (
            f"전제: 키 행 1개·키 열 1개 — {(h.fc._n_rows, h.fc._n_cols)} (설정이 섞였다)")
        yield h
        win.close(); win.deleteLater(); qapp.processEvents()


@pytest.fixture(scope="module")
def tall(qapp, tmp_path_factory):
    yield from _module_window(qapp, tmp_path_factory, "appdata_tall", 200, 4)


@pytest.fixture(scope="module")
def wide(qapp, tmp_path_factory):
    yield from _module_window(qapp, tmp_path_factory, "appdata_wide", 60, 40)


TOP_CELLS = {"B": [(9, 2), (7, 3)], "L": [(5, 0), (11, 0)], "T": [(0, 2), (0, 3)], "C": [(0, 0), (0, 0)]}


def _scroll(h, v=None, hz=None):
    if v is not None:
        h.t.verticalScrollBar().setValue(v)
    if hz is not None:
        h.t.horizontalScrollBar().setValue(hz)
    QApplication.instance().processEvents()
    h.fc.refresh()
    QApplication.instance().processEvents()


def test_cells_without_scrolling(tall):
    _scroll(tall, v=0)
    cell_matrix("셀", tall, TOP_CELLS, (15, 1))
    _report()


def test_cells_scrolled_down(tall):
    _scroll(tall, v=40)
    vr = visible_rows(tall.t)
    cell_matrix("셀(세로 스크롤)", tall,
                {"B": [(vr[4], 2), (vr[2], 3)], "L": [(vr[1], 0), (vr[6], 0)],
                 "T": [(0, 2), (0, 3)], "C": [(0, 0), (0, 0)]}, (vr[9], 1))
    _scroll(tall, v=0)
    _report()


def test_cells_scrolled_right(wide):
    _scroll(wide, hz=10)
    vc = visible_cols(wide.t)
    # 표 영역 폭은 환경(테스트 순서·설정)에 따라 달라진다 — 보이는 열 3개만 쓴다
    assert len(vc) >= 3 and vc[0] > 1, f"전제: 오른쪽으로 스크롤돼 열이 3개 이상 보여야 한다 {vc}"
    cell_matrix("셀(가로 스크롤)", wide,
                {"B": [(9, vc[1]), (7, vc[2])], "L": [(5, 0), (11, 0)],
                 "T": [(0, vc[1]), (0, vc[2])], "C": [(0, 0), (0, 0)]}, (15, vc[0]))
    _report()


def test_headers(tall):
    _scroll(tall, v=0)
    header_matrix(tall, (15, 1))
    _report()


def test_keyboard(tall):
    _scroll(tall, v=0)
    keyboard(tall)
    _report()


def test_right_click(tall):
    _scroll(tall, v=0)
    rightclick(tall)
    _report()


def test_dragging_below_the_table_keeps_scrolling(tall):
    """선택 엔진이 드래그를 직접 처리하므로 Qt 기본 드래그의 자동 스크롤을 대신해야 한다 —
    표 아래로 끌어 놓고 마우스를 멈춰도 계속 내려가며 선택이 따라온다. 떼면 멈춘다."""
    _scroll(tall, v=0)
    t = tall.t
    vp = t.viewport()
    tall.reset((15, 1))
    v0, p0 = tall.cpos((5, 1))
    g0 = v0.viewport().mapToGlobal(p0)
    below = vp.mapToGlobal(QPoint(p0.x(), vp.height() + 15))
    tall.send(vp, QEvent.MouseButtonPress, vp.mapFromGlobal(g0), NO)
    tall.send(vp, QEvent.MouseMove, vp.mapFromGlobal(below), NO, button=Qt.NoButton)
    bar0, row0 = t.verticalScrollBar().value(), tall.cur()[0]
    QTest.qWait(600)                                   # 마우스는 그대로 — 타이머만 돈다
    bar1, row1 = t.verticalScrollBar().value(), tall.cur()[0]
    tall.send(vp, QEvent.MouseButtonRelease, vp.mapFromGlobal(below), NO, buttons=Qt.NoButton)
    QTest.qWait(200)
    bar2 = t.verticalScrollBar().value()
    assert bar1 > bar0 and row1 > row0, f"멈춘 채로 스크롤이 안 된다: {bar0}->{bar1}, 행 {row0}->{row1}"
    QTest.qWait(300)
    assert t.verticalScrollBar().value() == bar2, "뗀 뒤에도 스크롤이 계속된다"
    rows = {r for r, _ in tall.cells()}
    assert min(rows) == 5 and max(rows) == row1, f"선택이 커서를 따라오지 않았다: {min(rows)}~{max(rows)}"
    _scroll(tall, v=0)
