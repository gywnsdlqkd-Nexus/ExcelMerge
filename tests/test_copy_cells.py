# -*- coding: utf-8 -*-
"""Ctrl+C — 선택한 셀을 엑셀이 읽는 TSV 로 클립보드에 복사한다.

복사 코드 자체는 예전부터 있었지만 **한 번도 동작한 적이 없었다**: 단축키를 패널마다
기본 컨텍스트(WindowShortcut)로 등록해 A/B 두 패널이 같은 창에 같은 Ctrl+C 를 건 꼴이
됐고, Qt 는 이럴 때 activated 대신 activatedAmbiguously 를 쏜다(= 아무 일도 안 일어남).
그래서 이 파일의 테스트는 전부 **실제 키 입력**(QTest.keyClick)으로 단축키 경로를 탄다 —
메서드를 직접 부르면 그 버그를 영영 못 잡는다.
"""
import pytest
from PyQt5.QtCore import Qt
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QApplication

from excelmerge.widgets import AX_ROW

KEY_COL = 3     # D열 = 키 헤더 → 틀 고정 밴드는 A~D
N_COLS = 8

HEADER = [f"H{c}" for c in range(N_COLS)]


def _row(i, val):
    r = [f"a{i}_{c}" for c in range(N_COLS)]
    r[KEY_COL] = f"k{i}"
    r[5] = val
    return r


# 짝수 i 만 값이 달라 '변경점만 보기'에 남는다(홀수 행은 숨는다).
A_DATA = [HEADER] + [_row(i, f"v{i}") for i in range(1, 13)]
B_DATA = [HEADER] + [_row(i, f"v{i}" if i % 2 else f"CHANGED{i}") for i in range(1, 13)]
# 붙여넣기를 깨뜨리는 값들 — 탭/줄바꿈/따옴표 (A쪽 표시값에 심는다)
A_DATA[2][6] = "탭\t포함"
A_DATA[4][6] = "두\n줄"
A_DATA[6][6] = '따옴표 "인용"'


def _wait_diff(dv, ms=8000):
    w = getattr(dv, "_diff_worker", None)
    if w is not None:
        w.wait(ms)
    QApplication.instance().processEvents()


@pytest.fixture
def dv(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    monkeypatch.setenv("APPDATA", str(tmp_path))   # 전역 키 앵커 저장 격리
    win = MainWindow()
    win.show()
    try:
        view = win.tabs.currentWidget()
        view.panel_a.set_path("a.xlsx")
        view.panel_b.set_path("b.xlsx")
        view._key_col, view._key_row = KEY_COL, 0
        for p in (view.panel_a, view.panel_b):
            p.table.set_key_col(KEY_COL)
            p.table.set_key_row(0)
        view._on_loaded(A_DATA, B_DATA)
        _wait_diff(view)
        assert view._freeze["a"].active
        yield view
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def _clip():
    return QApplication.instance().clipboard().text()


def _press_copy(w):
    """실제 Ctrl+C — 단축키 등록/컨텍스트까지 함께 검증한다.

    누른 뒤 Ctrl 을 반드시 떼 준다: QTest 의 키 릴리스도 modifiers 에 Ctrl 을 달고
    오기 때문에 Qt 가 캐시한 전역 수정자 상태가 'Ctrl 눌림'으로 남는다. 헤더 클릭
    처리(_on_h_section_pressed)는 이벤트가 아니라 QApplication.keyboardModifiers()
    를 읽으므로, 그대로 두면 **다음 테스트 파일**의 헤더 클릭이 Ctrl+클릭으로 오인된다
    (실제로 test_freeze_selection 의 헤더 앵커 테스트가 깨졌다).
    """
    w.setFocus()
    QApplication.instance().processEvents()
    QTest.keyClick(w, Qt.Key_C, Qt.ControlModifier)
    QTest.keyRelease(w, Qt.Key_Control, Qt.NoModifier)
    QApplication.instance().processEvents()


def _body_rows(dv, n):
    """본체에서 보이는 본문 행 n개."""
    tbl = dv.panel_a.table
    vis = [r for r in range(1, len(dv._diff_matrix)) if not tbl.isRowHidden(r)]
    assert len(vis) >= n, vis
    return vis[:n]


def _select(dv, r1, c1, r2, c2, side="a"):
    tbl = dv.panels[side].table
    tbl._select_range(r1, c1, r2, c2)
    tbl._set_current_cell_no_update(r2, c2)
    QApplication.instance().processEvents()
    return tbl


# ── 본체: 고른 사각형이 그대로 TSV 로 ──────────────────────────────────────────

def test_ctrl_c_copies_the_selected_block(dv):
    """2행 × 2열 사각형 — 칸은 탭, 행은 CRLF(엑셀이 기대하는 구분자)."""
    r1, r2 = _body_rows(dv, 2)
    tbl = _select(dv, r1, 4, r2, 5)
    QApplication.instance().clipboard().setText("<이전 값>")
    _press_copy(tbl)
    m = tbl.model()
    expect = "\r\n".join(
        "\t".join(m.display_text(r, c) for c in (4, 5)) for r in (r1, r2))
    assert _clip() == expect
    assert len(_clip().split("\r\n")) == 2 and _clip().count("\t") == 2


def test_copy_reads_the_panel_that_has_focus(dv):
    """A/B 두 패널이 같은 단축키를 걸어도 포커스 쪽 값이 나와야 한다(모호함 회귀)."""
    r = _body_rows(dv, 1)[0]
    a = _select(dv, r, 5, r, 5, "a")
    b = _select(dv, r, 5, r, 5, "b")
    _press_copy(a)
    assert _clip() == a.model().display_text(r, 5)
    _press_copy(b)
    assert _clip() == b.model().display_text(r, 5)
    assert a.model().display_text(r, 5) != b.model().display_text(r, 5), \
        "전제: 이 칸은 A/B 값이 달라야 구분이 된다"


def test_single_cell_copies_without_separators(dv):
    r = _body_rows(dv, 1)[0]
    tbl = _select(dv, r, 4, r, 4)
    _press_copy(tbl)
    assert _clip() == tbl.model().display_text(r, 4)


# ── 화면에서 볼 수 없는 것은 복사하지 않는다 ─────────────────────────────────

def test_filtered_rows_are_not_copied(dv):
    """'변경점만 보기'로 숨은 행은 빠진다 — 엑셀도 필터가 걸리면 보이는 행만 복사한다."""
    tbl = dv.panel_a.table
    r1, r2 = _body_rows(dv, 2)
    assert any(tbl.isRowHidden(r) for r in range(r1 + 1, r2)), \
        "전제: 두 행 사이에 필터로 숨은 행이 있어야 의미 있는 테스트"
    _select(dv, r1, 5, r2, 5)
    _press_copy(tbl)
    assert _clip() == "\r\n".join(
        tbl.model().display_text(r, 5) for r in (r1, r2))


def test_select_all_copies_only_visible_rows(dv):
    """Ctrl+A 는 숨은 행까지 잡지만, 복사되는 줄 수는 보이는 행 수와 같아야 한다."""
    tbl = dv.panel_a.table
    tbl.selectAll()
    QApplication.instance().processEvents()
    _press_copy(tbl)
    visible = [r for r in range(tbl.rowCount()) if tbl._navigable(AX_ROW, r)]
    assert len(visible) < tbl.rowCount(), "전제: 필터로 숨은 행이 있어야 한다"
    assert len(_clip().split("\r\n")) == len(visible)


def test_frozen_band_columns_are_copied(dv):
    """틀 고정 밴드(A~D)는 본체에서 숨김이지만 화면엔 보이므로 복사에 포함된다."""
    tbl = dv.panel_a.table
    assert tbl.isColumnHidden(0) and tbl.isColumnHidden(KEY_COL), "전제: 밴드는 본체 숨김"
    r = _body_rows(dv, 1)[0]
    _select(dv, r, 0, r, 5)
    _press_copy(tbl)
    cells = _clip().split("\t")
    assert len(cells) == 6
    assert cells[0] == tbl.model().display_text(r, 0)
    assert cells[KEY_COL] == tbl.model().display_text(r, KEY_COL)


def test_key_row_is_copied_too(dv):
    """키 행(0)도 상단 밴드에 보이므로 함께 복사된다."""
    tbl = dv.panel_a.table
    assert tbl.isRowHidden(0), "전제: 키 행은 본체에서 숨김"
    r = _body_rows(dv, 1)[0]
    _select(dv, 0, 5, r, 5)
    _press_copy(tbl)
    lines = _clip().split("\r\n")
    assert lines[0] == tbl.model().display_text(0, 5)


# ── 엑셀이 읽을 수 있는 형태로 ────────────────────────────────────────────────

@pytest.mark.parametrize("row_i, raw", [
    (2, "탭\t포함"),
    (4, "두\n줄"),
    (6, '따옴표 "인용"'),
])
def test_special_characters_are_quoted(dv, row_i, raw):
    """값 안의 탭·줄바꿈·따옴표는 큰따옴표로 감싸야 칸/행이 갈라지지 않는다."""
    tbl = dv.panel_a.table
    assert not tbl.isRowHidden(row_i), "전제: 이 행은 화면에 보여야 한다"
    assert tbl.model().display_text(row_i, 6) == raw
    _select(dv, row_i, 6, row_i, 6)
    _press_copy(tbl)
    assert _clip() == '"' + raw.replace('"', '""') + '"'


def test_plain_value_is_not_quoted(dv):
    r = _body_rows(dv, 1)[0]
    _select(dv, r, 5, r, 5)
    _press_copy(dv.panel_a.table)
    assert not _clip().startswith('"')


# ── 표 밖에서 누른 Ctrl+C ─────────────────────────────────────────────────────

def test_text_widgets_keep_their_own_copy(dv):
    """경로칸에서는 윈도우 표준대로 '고른 글자만' 복사된다.

    편집 위젯은 자기가 쓰는 키에 ShortcutOverride 를 돌려줘 앱 단축키를 끈다 —
    즉 우리 핸들러가 아니라 QLineEdit 이 복사한다. 표 복사를 붙이면서 이 표준
    동작을 뺏지 않았는지 고정한다.
    """
    panel = dv.panel_a
    panel.path_edit.setText("ABCDEF")
    panel.path_edit.setSelection(0, 3)
    _press_copy(panel.path_edit)
    assert _clip() == "ABC"


# ── 상태바 피드백 ─────────────────────────────────────────────────────────────

def test_status_message_reports_the_copied_size(dv):
    seen = []
    dv.panel_a.status_message.connect(seen.append)
    r1, r2 = _body_rows(dv, 2)
    tbl = _select(dv, r1, 4, r2, 6)
    _press_copy(tbl)
    assert seen and seen[-1] == "2행 × 3열 복사됨"


def test_empty_selection_leaves_the_clipboard_alone(dv):
    tbl = dv.panel_a.table
    seen = []
    dv.panel_a.status_message.connect(seen.append)
    tbl.clearSelection()
    QApplication.instance().processEvents()
    QApplication.instance().clipboard().setText("<이전 값>")
    _press_copy(tbl)
    assert _clip() == "<이전 값>"
    assert seen and "없습니다" in seen[-1]
