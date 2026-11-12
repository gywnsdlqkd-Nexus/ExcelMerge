# -*- coding: utf-8 -*-
"""탭이 하나뿐이면 ✕ 를 숨긴다 — 마지막 탭을 닫으면 빈 탭을 새로 만들 뿐이고, 누르면 엉뚱한 창이
잠깐 떴다(사용자 제보). Ctrl+W 도 마지막 탭은 닫지 않는다."""
from PyQt5.QtWidgets import QTabBar

from excelmerge.main_window import MainWindow


def _buttons(win):
    """보이는 ✕ 만 — 숨긴 버튼은 None 으로."""
    bar = win.tabs.tabBar()
    out = []
    for i in range(win.tabs.count()):
        b = bar.tabButton(i, QTabBar.RightSide)
        out.append(b if b is not None and not b.isHidden() else None)
    return out


def test_single_tab_has_no_close_button(qapp):
    win = MainWindow()
    try:
        assert win.tabs.count() == 1
        assert _buttons(win) == [None]
    finally:
        win.close()


def test_close_buttons_follow_the_tab_count(qapp):
    win = MainWindow()
    try:
        win._new_tab()
        assert win.tabs.count() == 2 and all(b is not None for b in _buttons(win))
        win._new_tab()
        assert win.tabs.count() == 3 and all(b is not None for b in _buttons(win))
        _buttons(win)[0].click()
        qapp.processEvents()
        assert win.tabs.count() == 2 and all(b is not None for b in _buttons(win))
        _buttons(win)[1].click()
        qapp.processEvents()
        assert win.tabs.count() == 1 and _buttons(win) == [None]
        # 숨긴 ✕ 는 탭 폭에 빈자리를 남기지 않는다
        hidden = win.tabs.tabBar().tabButton(0, QTabBar.RightSide)
        assert hidden is None or hidden.width() == 0
    finally:
        win.close()


def test_ctrl_w_keeps_the_last_tab(qapp):
    win = MainWindow()
    try:
        only = win.tabs.widget(0)
        win._close_current_tab()
        qapp.processEvents()
        assert win.tabs.count() == 1 and win.tabs.widget(0) is only
    finally:
        win.close()


def test_corner_buttons_keep_their_height_with_one_tab(qapp):
    """✕ 를 숨겨 탭이 낮아져도 우상단 코너 버튼('업데이트 확인'·'＋ 새 비교'·'＋ 3파일 비교')이
    눌리지 않는다 — QTabWidget 은 코너 높이를 탭 높이로 자른다(v231 에서 글자가 잘림, 사용자 제보)."""
    from PyQt5.QtWidgets import QPushButton
    win = MainWindow()
    try:
        win.show()
        corner = win.tabs.cornerWidget()

        def heights():
            qapp.processEvents()
            return [b.height() for b in corner.findChildren(QPushButton)]

        assert win.tabs.tabBar().tabSizeHint(0).height() >= corner.sizeHint().height()
        one = heights()
        win._new_tab()
        two = heights()
        win._close_tab(1)
        assert one == two == heights()
        hint = corner.findChildren(QPushButton)[0].sizeHint().height()
        assert min(one) >= hint - 1, (one, hint)
    finally:
        win.close()


def test_merge3_replacing_the_only_empty_tab_has_no_close_button(qapp, tmp_path):
    """빈 비교 탭 자리를 3파일 탭으로 바꿔 끼워도 탭은 하나 — ✕ 없음."""
    import openpyxl
    paths = []
    for name in ("l", "b", "r"):
        wb = openpyxl.Workbook()
        wb.active.append(["ID", "V"])
        wb.active.append(["1", name])
        p = tmp_path / f"{name}.xlsx"
        wb.save(str(p))
        paths.append(str(p))
    win = MainWindow()
    try:
        win.open_merge3(*paths)
        qapp.processEvents()
        assert win.tabs.count() == 1 and _buttons(win) == [None]
        win._new_tab()
        assert win.tabs.count() == 2 and all(b is not None for b in _buttons(win))
    finally:
        for i in range(win.tabs.count()):
            w = win.tabs.widget(i)
            if hasattr(w, "shutdown"):
                w.shutdown()
        win.close()
