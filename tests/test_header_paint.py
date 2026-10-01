# -*- coding: utf-8 -*-
"""헤더 색이 **실제로 칠해지는지** — 모델이 뭘 반환하는지가 아니라 그려진 픽셀을 본다.

실기에서 드러난 일: 열 매칭을 넣으며 '한쪽에만 있는 열 헤더는 연초록' 이라고 적고
모델이 그 색을 반환하는지만 테스트했다. 통과했다. 그런데 화면에는 **한 픽셀도** 칠해지지
않았다. 더 파 보니 키 열 노랑(HEADER_KEY_BG)도 같은 이유로 **한 번도** 칠해진 적이 없었다.

원인 둘:
  1. 앱에 스타일시트가 걸려 있으면 Qt 가 헤더 섹션을 직접 그리며 BackgroundRole 을 무시한다.
  2. QHeaderView.paintSection 은 painter 에 **클립을 남기고** 돌아온다. 그대로 덧칠하면
     전부 잘려 나간다(실측: 불투명 빨강조차 0픽셀). super() 를 save/restore 로 감싸야 한다.

그래서 이 파일의 테스트는 전부 grab() 으로 그려 보고 픽셀을 읽는다. 모델 반환값만 보는
테스트는 이 사고를 또 놓친다.
"""
import pytest
from PyQt5.QtCore import Qt

from excelmerge.theme import (HEADER_EXCL_BG, HEADER_KEY_BG, HEADER_NORMAL_BG,
                              HEADER_ONESIDE_BG, HEADER_TINT_ALPHA)

A_DATA = [["ID", "NAME", "VALUE", "MEMO"], ["k1", "칼", "10", "m"], ["k2", "방패", "20", "m"]]
B_DATA = [["ID", "GRADE", "NAME", "VALUE", "MEMO"],
          ["k1", "A", "검", "10", "m"], ["k2", "S", "방패", "20", "m"]]


@pytest.fixture
def view(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    monkeypatch.setenv("APPDATA", str(tmp_path))
    win = MainWindow()
    win.resize(1200, 700)
    win.show()
    try:
        v = win.tabs.currentWidget()
        v.panel_a.set_path(str(tmp_path / "a.xlsx"))
        v.panel_b.set_path(str(tmp_path / "b.xlsx"))
        v._on_loaded(A_DATA, B_DATA)
        w = getattr(v, "_diff_worker", None)
        if w is not None:
            w.wait(8000)
        for _ in range(40):
            qapp.processEvents()
        yield v
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def _painted(header, section):
    """그 섹션에 실제로 그려진 색. 글자를 피해 섹션 위쪽을 읽는다."""
    img = header.grab().toImage()
    x = header.sectionPosition(section) + header.sectionSize(section) // 2
    x = min(max(x, 0), img.width() - 1)
    c = img.pixelColor(x, 3)
    return (c.red(), c.green(), c.blue())


def _expected_tint(bg, base=HEADER_NORMAL_BG):
    """base 위에 bg 를 HEADER_TINT_ALPHA 로 얹었을 때 나와야 하는 색(±2)."""
    a = HEADER_TINT_ALPHA / 255.0
    return tuple(round(bg_c * a + base_c * (1 - a))
                 for bg_c, base_c in ((bg.red(), base.red()),
                                      (bg.green(), base.green()),
                                      (bg.blue(), base.blue())))


def _close(got, want, tol=3):
    return all(abs(g - w) <= tol for g, w in zip(got, want))


def _col(view, name):
    return view._display_header_names(view._diff_col_meta).index(name)


# ── 한쪽에만 있는 열 ─────────────────────────────────────────────────────────

def test_a_one_sided_column_header_is_actually_green(view):
    hdr = view.panel_b.table.horizontalHeader()
    got = _painted(hdr, _col(view, "GRADE"))
    want = _expected_tint(HEADER_ONESIDE_BG)
    assert _close(got, want), f"그려진 색 {got}, 기대 {want} — 연초록이 안 칠해졌다"


def test_a_normal_column_header_is_untouched(view):
    hdr = view.panel_b.table.horizontalHeader()
    got = _painted(hdr, _col(view, "NAME"))
    assert _close(got, (HEADER_NORMAL_BG.red(), HEADER_NORMAL_BG.green(),
                        HEADER_NORMAL_BG.blue())), got


def test_both_panels_show_the_tint(view):
    """A 패널에서는 그 열이 '-' 로 보이지만 색은 똑같이 칠해져야 한다."""
    want = _expected_tint(HEADER_ONESIDE_BG)
    c = _col(view, "GRADE")
    for panel in (view.panel_a, view.panel_b):
        hdr = panel.table.horizontalHeader()
        assert _close(_painted(hdr, c), want), f"{panel.side}: {_painted(hdr, c)}"


# ── 검사 제외 열 ─────────────────────────────────────────────────────────────

def test_an_excluded_column_header_is_actually_grey(view, qapp):
    c = _col(view, "MEMO")
    before = _painted(view.panel_b.table.horizontalHeader(), c)
    view._on_columns_exclude_set([c], True)
    for _ in range(20):
        qapp.processEvents()
    after = _painted(view.panel_b.table.horizontalHeader(), c)
    assert after != before, "제외했는데 헤더 색이 그대로다"
    assert _close(after, _expected_tint(HEADER_EXCL_BG)), after


# ── 키 열 — 고정 밴드 오버레이가 그린다 ────────────────────────────────────

def test_the_key_column_header_is_actually_yellow(view):
    """키 열은 본체에서 숨겨지고 고정 밴드 오버레이가 그린다.

    그 오버레이의 헤더까지 바꾸지 않으면 키 노랑은 거기서만 또 안 나온다 —
    이 색은 이 수정 전까지 **한 번도** 칠해진 적이 없었다.
    """
    fz = view.panel_b.table._freeze
    assert fz is not None and fz._active, "전제: 틀 고정이 켜져 있어야 한다"
    hdr = fz.corner.horizontalHeader()
    assert hdr.isSectionHidden(0) is False
    got = _painted(hdr, 0)
    assert _close(got, _expected_tint(HEADER_KEY_BG)), (
        f"그려진 색 {got}, 기대 {_expected_tint(HEADER_KEY_BG)} — 키 노랑이 안 칠해졌다")


def test_the_overlay_header_can_still_be_clicked(view):
    """헤더를 교체하면서 클릭 플래그를 빠뜨리면 키 열 헤더 클릭/메뉴가 죽는다."""
    fz = view.panel_b.table._freeze
    hdr = fz.corner.horizontalHeader()
    assert hdr.sectionsClickable() and hdr.highlightSections()


# ── 글자가 묻히지 않는다 ────────────────────────────────────────────────────
# 덮어 칠하면 열 문자가 사라진다. 그래서 반투명으로 얹는다.
# ※ 글자 자체는 여기서 확인할 수 없다 — 오프스크린 grab() 은 배경·테두리만 그리고
#   글자는 렌더링하지 않는다(실측: 어떤 섹션에서도 어두운 픽셀이 0개). 가독성은
#   실기 스크린샷으로 본다. 여기서는 '덮어 칠하지 않는다'는 규칙만 고정한다.

def test_the_tint_is_translucent():
    assert 0 < HEADER_TINT_ALPHA < 255, "불투명하게 칠하면 열 문자가 사라진다"


def test_an_opaque_tint_would_fail_the_colour_tests(view):
    """위 색 테스트들이 '불투명하게 칠해도 통과' 하면 안 된다 — 경계를 확인한다."""
    opaque = (HEADER_ONESIDE_BG.red(), HEADER_ONESIDE_BG.green(), HEADER_ONESIDE_BG.blue())
    assert not _close(_expected_tint(HEADER_ONESIDE_BG), opaque), (
        "기대색이 원색과 같다 — 반투명 여부를 구분하지 못한다")
