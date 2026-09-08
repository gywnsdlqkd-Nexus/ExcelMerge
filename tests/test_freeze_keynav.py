# -*- coding: utf-8 -*-
"""틀 고정 밴드로 들어가는 키보드 이동 — 숨겨져 있어도 오버레이가 보여주므로 갈 수 있어야 한다.

틀 고정이 켜지면 본체(host)가 키 밴드를 **실제로 숨기고**(열: _apply_col_hidden,
행: _apply_diff_filter 의 0..key_row) 오버레이가 그 자리를 그린다. Qt 기본 moveCursor 는
숨긴 섹션을 '갈 수 없는 곳'으로 취급하므로, 키 열이 D일 때 E에서 ← 를 눌러도 제자리였다.
← / Shift+← / Home / ↑ / Ctrl+↑ 가 모두 같은 이유로 막혀 있었다.

반대로 **'변경점만 보기'로 숨은 행은 계속 갈 수 없어야 한다**(정말로 볼 수 없으므로).
그 경계를 여기서 고정한다.
"""
import pytest
from PyQt5.QtCore import Qt, QEvent
from PyQt5.QtGui import QKeyEvent
from PyQt5.QtWidgets import QApplication

KEY_COL = 3     # D열 = 키 헤더
N_COLS = 8

# 행 0 = 키 행(헤더). 본문은 키 "k1".. — 짝수 번째만 값이 달라 필터에 남는다.
HEADER = [f"H{c}" for c in range(N_COLS)]
def _row(i, val):
    r = [f"a{i}_{c}" for c in range(N_COLS)]
    r[KEY_COL] = f"k{i}"
    r[5] = val
    return r
A_DATA = [HEADER] + [_row(i, f"v{i}") for i in range(1, 20)]
B_DATA = [HEADER] + [_row(i, f"v{i}" if i % 2 else f"CHANGED{i}") for i in range(1, 20)]


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
    # 키 앵커 전역 저장이 실제 %APPDATA% 를 건드리지 않도록 격리.
    monkeypatch.setenv("APPDATA", str(tmp_path))
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
        assert view._freeze["a"]._n_cols == KEY_COL + 1
        yield view
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def _press(tbl, key, mods=Qt.NoModifier):
    tbl.keyPressEvent(QKeyEvent(QEvent.KeyPress, key, mods))
    QApplication.instance().processEvents()


def _body_row(dv, nth=0):
    """본체에서 보이는 n번째 본문 행."""
    tbl = dv.panel_a.table
    vis = [r for r in range(len(dv._diff_matrix)) if not tbl.isRowHidden(r)]
    return vis[nth]


def _move(dv, r, c, key, mods=Qt.NoModifier):
    tbl = dv.panel_a.table
    tbl._move_current_cell(r, c)
    QApplication.instance().processEvents()
    _press(tbl, key, mods)
    return tbl._current_cell()


# ── 밴드로 들어가기 (버그 본체) ───────────────────────────────────────────────

def test_left_enters_frozen_band(dv):
    """E열에서 ← → D열(키 열). 보고된 버그."""
    r = _body_row(dv)
    assert dv.panel_a.table.isColumnHidden(KEY_COL), "전제: 키 열은 본체에서 숨겨져 있다"
    assert _move(dv, r, KEY_COL + 1, Qt.Key_Left) == (r, KEY_COL)


def test_left_moves_inside_band(dv):
    """밴드 안에서도 계속 왼쪽으로 (D → C)."""
    r = _body_row(dv)
    assert _move(dv, r, KEY_COL, Qt.Key_Left) == (r, KEY_COL - 1)


def test_left_stops_at_grid_edge(dv):
    """A열에서 ← 는 제자리 — 격자 끝이라 정상."""
    r = _body_row(dv)
    assert _move(dv, r, 0, Qt.Key_Left) == (r, 0)


def test_home_goes_to_first_column(dv):
    r = _body_row(dv)
    assert _move(dv, r, KEY_COL + 2, Qt.Key_Home) == (r, 0)


def test_up_reaches_key_row(dv):
    """첫 본문 행에서 ↑ → 키 행(0). 행 축도 같은 뿌리로 깨져 있었다."""
    r = _body_row(dv)
    assert dv.panel_a.table.isRowHidden(0), "전제: 키 행은 본체에서 숨겨져 있다"
    assert _move(dv, r, 5, Qt.Key_Up) == (0, 5)


def test_ctrl_up_reaches_key_row(dv):
    """Ctrl+↑ 도 키 행에 착지 — _next_visible_row 가 키 행을 건너뛰고 있었다."""
    r = _body_row(dv)
    assert _move(dv, r, 5, Qt.Key_Up, Qt.ControlModifier) == (0, 5)


def test_shift_left_extends_selection_into_band(dv):
    """Shift+← 로 밴드까지 선택이 늘어나고, 정규화가 되걷어내지 않는다."""
    r = _body_row(dv)
    tbl = dv.panel_a.table
    assert _move(dv, r, KEY_COL + 1, Qt.Key_Left, Qt.ShiftModifier) == (r, KEY_COL)
    cells = tbl.get_selected_cells()
    assert (r, KEY_COL) in cells, "키 열 셀이 선택되지 않았다"
    assert (r, KEY_COL + 1) in cells, "출발 셀이 선택에서 빠졌다"


def test_shift_left_accumulates_across_presses(dv):
    """Shift+← 를 연달아 누르면 확장이 누적된다(두 칸으로 붕괴하지 않는다)."""
    r = _body_row(dv)
    tbl = dv.panel_a.table
    tbl._move_current_cell(r, KEY_COL + 1)
    QApplication.instance().processEvents()
    for _ in range(3):
        _press(tbl, Qt.Key_Left, Qt.ShiftModifier)
    cells = tbl.get_selected_cells()
    assert {(r, c) for c in range(KEY_COL - 2, KEY_COL + 2)} <= cells, sorted(cells)


# ── 밴드 '안에서' 위/아래 (2차 보고 버그) ─────────────────────────────────────
# 첫 수정은 '목적지가 밴드일 때'만 Qt 를 우회했다. 그래서 커서가 이미 밴드 안(A~D)에
# 있을 때의 ↑/↓ 는 목적지 행이 밴드가 아니므로 그대로 Qt 에 넘어갔고, Qt 는 **출발 칸의
# 열이 숨김**이라 moveCursor 에서 무효 인덱스를 돌려줘 제자리였다.

@pytest.mark.parametrize("col", list(range(KEY_COL + 1)))
def test_down_inside_band(dv, col):
    """밴드 안 어느 열에서든 ↓ 는 다음 보이는 행으로. 보고된 2차 버그."""
    assert dv.panel_a.table.isColumnHidden(col), "전제: 밴드 열은 본체에서 숨김"
    assert _move(dv, _body_row(dv, 1), col, Qt.Key_Down) == (_body_row(dv, 2), col)


@pytest.mark.parametrize("col", list(range(KEY_COL + 1)))
def test_up_inside_band(dv, col):
    assert _move(dv, _body_row(dv, 2), col, Qt.Key_Up) == (_body_row(dv, 1), col)


def test_shift_down_inside_band_extends_selection(dv):
    r1, r2 = _body_row(dv, 1), _body_row(dv, 2)
    tbl = dv.panel_a.table
    assert _move(dv, r1, 0, Qt.Key_Down, Qt.ShiftModifier) == (r2, 0)
    cells = tbl.get_selected_cells()
    assert (r1, 0) in cells and (r2, 0) in cells, sorted(cells)


def test_up_inside_band_reaches_key_row(dv):
    """밴드 안에서 위로 끝까지 — 키 행(0)까지 도달하고 거기서 멈춘다."""
    assert _move(dv, _body_row(dv, 0), 0, Qt.Key_Up) == (0, 0)
    assert _move(dv, 0, 0, Qt.Key_Up) == (0, 0)


def test_home_inside_band_stays_put(dv):
    """A열에서 Home 은 제자리 — Qt 에 넘기면 숨은 열을 피해 첫 데이터 열로 튄다."""
    r = _body_row(dv)
    assert _move(dv, r, 0, Qt.Key_Home) == (r, 0)


def test_vertical_move_inside_band_skips_filtered_rows(dv):
    """밴드 안에서의 ↑/↓ 도 필터로 숨은 행은 밟지 않는다."""
    tbl = dv.panel_a.table
    assert [r for r in range(1, len(dv._diff_matrix)) if tbl.isRowHidden(r)]
    for key in (Qt.Key_Up, Qt.Key_Down):
        tbl._move_current_cell(_body_row(dv, 2), 0)
        QApplication.instance().processEvents()
        for _ in range(30):
            _press(tbl, key)
            r, c = tbl._current_cell()
            assert c == 0, "세로 이동인데 열이 바뀌었다"
            assert (not tbl.isRowHidden(r)) or r == 0, f"볼 수 없는 행 {r} 에 착지했다"


def test_band_cursor_drags_the_viewport_along(dv):
    """밴드 열 위에서 ↓ 를 눌러도 화면이 따라온다.

    QTableView.scrollTo 는 isIndexHidden 인 인덱스를 통째로 무시한다 — 밴드 열은 본체에서
    숨김이라, 커서만 내려가고 뷰포트는 그대로 있어 커서가 화면 밖으로 걸어나갔다.
    밴드 열과 일반 열의 스크롤 결과가 같아야 한다.
    """
    tbl = dv.panel_a.table
    tbl.setFixedHeight(70)          # 스크롤이 필요한 크기로 축소
    QApplication.instance().processEvents()
    assert tbl.verticalScrollBar().maximum() > 0, "전제: 스크롤이 필요한 상태"

    def walk(col):
        tbl.verticalScrollBar().setValue(0)
        tbl._move_current_cell(_body_row(dv, 0), col)
        QApplication.instance().processEvents()
        off = 0
        for _ in range(8):
            _press(tbl, Qt.Key_Down)
            r, _c = tbl._current_cell()
            bot = tbl.rowAt(tbl.viewport().height() - 1)
            if bot < 0:
                bot = tbl.rowCount() - 1
            if not (tbl.rowAt(0) <= r <= bot):
                off += 1
        return tbl.verticalScrollBar().value(), off

    band_v, band_off = walk(0)
    plain_v, plain_off = walk(KEY_COL + 1)
    assert band_off == 0, f"밴드 열에서 커서가 {band_off}번 화면 밖으로 나갔다"
    assert band_v == plain_v, f"스크롤 불일치: 밴드 {band_v} vs 일반 {plain_v}"


def test_scroll_assist_keeps_horizontal_position(dv):
    """세로 보조 스크롤이 가로 위치를 흔들지 않는다."""
    tbl = dv.panel_a.table
    tbl.setFixedWidth(120)
    QApplication.instance().processEvents()
    hbar = tbl.horizontalScrollBar()
    if hbar.maximum() == 0:
        pytest.skip("가로 스크롤이 필요 없는 폭")
    hbar.setValue(hbar.maximum())
    tbl._move_current_cell(_body_row(dv, 0), 0)
    QApplication.instance().processEvents()
    before = hbar.value()
    for _ in range(4):
        _press(tbl, Qt.Key_Down)
    assert hbar.value() == before


# ── Shift 확장의 앵커 (3차 보고 버그) ────────────────────────────────────────
# A1~G1 을 잡고 Shift+↓ 를 연타하면 두 번째 눌렀을 때 선택이 무너졌다 — 고정 밴드 열
# A~D 가 통째로 빠지고, 엉뚱한 아래쪽 두 행 × G열만 남았다.
#
# Qt 의 Shift+방향키 앵커는 QAbstractItemViewPrivate::pressedPosition, 즉 **뷰포트 픽셀
# 좌표**다. 확장은 그 픽셀에서 새 커서 칸 중심까지의 사각형을 indexAt 으로 되짚어
# 만들어진다. 이 좌표는 마우스 press 만 갱신하므로 (a) 우리가 _select_range 로 만든
# 선택(오버레이 드래그·헤더·찾기·Ctrl+점프) 뒤에는 낡아 있고, 무효로 판정되면 Qt 는
# 조용히 **직전 커서 칸으로 재앵커**해 선택을 붕괴시킨다. 게다가 (b) 픽셀→인덱스라
# 폭 0 으로 숨긴 밴드 열은 애초에 표현할 수 없다.
# 그래서 Shift 확장은 밴드 유무와 무관하게 항상 _shift_anchor 로 직접 만든다.

def _box(tbl):
    """선택 전체의 바운딩 박스 (top, left, bottom, right)."""
    rngs = list(tbl.selectionModel().selection())
    if not rngs:
        return None
    return (min(r.top() for r in rngs), min(r.left() for r in rngs),
            max(r.bottom() for r in rngs), max(r.right() for r in rngs))


def _select_a1_g1(dv):
    """A1~G1 (키 행 × 밴드 안팎에 걸친 7열) 선택 — 오버레이 드래그와 같은 경로."""
    tbl = dv.panel_a.table
    tbl._select_range(0, 0, 0, 6)
    tbl._set_current_cell_no_update(0, 6)
    QApplication.instance().processEvents()
    return tbl


def test_shift_down_from_key_row_keeps_anchor_and_band(dv):
    """보고된 버그 그대로 — A1:G1 에서 Shift+↓ 두 번. 두 번째에서 무너졌었다."""
    tbl = _select_a1_g1(dv)
    assert _box(tbl) == (0, 0, 0, 6)
    _press(tbl, Qt.Key_Down, Qt.ShiftModifier)
    assert _box(tbl) == (0, 0, _body_row(dv, 0), 6)
    _press(tbl, Qt.Key_Down, Qt.ShiftModifier)
    assert _box(tbl) == (0, 0, _body_row(dv, 1), 6),         "두 번째 Shift+↓ 에서 앵커가 풀리고 밴드 열이 빠졌다"


def test_shift_down_accumulates_over_many_presses(dv):
    """연타해도 앵커(0,0)가 고정되고 밴드 열 A~D 가 계속 선택에 남는다."""
    tbl = _select_a1_g1(dv)
    for n in range(4):
        _press(tbl, Qt.Key_Down, Qt.ShiftModifier)
        top, left, bot, right = _box(tbl)
        assert (top, left, right) == (0, 0, 6), f"{n + 1}번째에서 앵커/열이 어긋났다"
        assert bot == _body_row(dv, n)


def test_shift_down_keeps_selection_outside_band_too(dv):
    """밴드가 전혀 안 걸린 선택도 같은 앵커 버그를 겪었다 — 열이 커서 열로 붕괴했다."""
    tbl = dv.panel_a.table
    r = _body_row(dv, 0)
    tbl._select_range(r, KEY_COL + 1, r, KEY_COL + 3)
    tbl._set_current_cell_no_update(r, KEY_COL + 3)
    QApplication.instance().processEvents()
    _press(tbl, Qt.Key_Down, Qt.ShiftModifier)
    assert _box(tbl) == (r, KEY_COL + 1, _body_row(dv, 1), KEY_COL + 3)


def test_shift_up_shrinks_without_losing_anchor(dv):
    """↓ 로 늘렸다가 ↑ 로 줄여도 앵커는 그대로 — 커서가 늘 사각형 모서리에 있다."""
    tbl = _select_a1_g1(dv)
    for _ in range(3):
        _press(tbl, Qt.Key_Down, Qt.ShiftModifier)
    assert _box(tbl) == (0, 0, _body_row(dv, 2), 6)
    _press(tbl, Qt.Key_Up, Qt.ShiftModifier)
    assert _box(tbl) == (0, 0, _body_row(dv, 1), 6)


def test_ctrl_shift_down_keeps_shift_anchor(dv):
    """Ctrl+Shift+↓ 도 앵커를 이어받는다 — currentIndex 를 앵커로 쓰면 붕괴했다."""
    tbl = _select_a1_g1(dv)
    for _ in range(2):
        _press(tbl, Qt.Key_Down, Qt.ShiftModifier)
    _press(tbl, Qt.Key_Down, Qt.ControlModifier | Qt.ShiftModifier)
    top, left, _bot, right = _box(tbl)
    assert (top, left, right) == (0, 0, 6), "Ctrl+Shift+↓ 가 앵커를 커서로 붕괴시켰다"


def test_ctrl_shift_home_extends_to_the_anchor(dv):
    """앵커가 A1 이면 Ctrl+Shift+Home 은 A1 한 칸으로 접힌다(엑셀과 동일)."""
    tbl = _select_a1_g1(dv)
    for _ in range(2):
        _press(tbl, Qt.Key_Down, Qt.ShiftModifier)
    _press(tbl, Qt.Key_Home, Qt.ControlModifier | Qt.ShiftModifier)
    assert _box(tbl) == (0, 0, 0, 0)


def test_shift_extension_skips_filtered_rows(dv):
    """확장의 '끝'은 언제나 볼 수 있는 행 — 숨은 행에 커서를 세우지 않는다."""
    tbl = _select_a1_g1(dv)
    for _ in range(6):
        _press(tbl, Qt.Key_Down, Qt.ShiftModifier)
        r, _c = tbl._current_cell()
        assert not tbl.isRowHidden(r), f"숨은 행 {r} 에 확장 끝이 놓였다"


def test_plain_arrow_outside_band_stays_a_single_cell(dv):
    """비회귀: Shift 없는 평범한 이동은 여전히 단일 선택 — Qt 기본 경로 그대로."""
    tbl = dv.panel_a.table
    r = _body_row(dv, 0)
    tbl._move_current_cell(r, KEY_COL + 2)
    QApplication.instance().processEvents()
    _press(tbl, Qt.Key_Down)
    assert tbl._current_cell() == (_body_row(dv, 1), KEY_COL + 2)
    assert _box(tbl) == (_body_row(dv, 1), KEY_COL + 2,
                         _body_row(dv, 1), KEY_COL + 2)


# ── End / PageUp / PageDown ──────────────────────────────────────────────────
# Shift 확장은 위와 같은 픽셀 앵커 문제를 겪는다. 여기에 페이지 이동만의 문제가 하나 더
# 있다: Qt 의 PageUp/PageDown 은 위 가드 앞에서 일찍 return 하므로 밴드에서도 '제자리'가
# 되진 않지만, 걸음을 뷰포트 픽셀로 재는 탓에 밴드 열에서 일반 열과 **다른 칸**에
# 착지했다(실측: 같은 행에서 A열은 +1116, E열은 +1841). 화면 행을 직접 세어 통일한다.

def _shrink(dv, h=70):
    tbl = dv.panel_a.table
    tbl.setFixedHeight(h)
    QApplication.instance().processEvents()
    assert tbl.verticalScrollBar().maximum() > 0, "전제: 스크롤이 필요한 상태"
    return tbl


def _page_land(dv, tbl, start, col, key, vbar):
    """커서와 스크롤 오프셋을 **둘 다** 고정한 뒤 페이지 키를 눌러 착지 행을 얻는다.

    스크롤을 고정하지 않으면 밴드 열은 scrollTo 가 안 먹혀(숨긴 인덱스) 뷰포트가 커서와
    따로 놀고, 일반 열은 정렬된다 — 그 차이가 곧 버그라서 여기서 통제해야 드러난다.
    """
    tbl._move_current_cell(start, col)
    QApplication.instance().processEvents()
    tbl.verticalScrollBar().setValue(vbar)
    QApplication.instance().processEvents()
    _press(tbl, key)
    return tbl._current_cell()[0]


@pytest.mark.parametrize("vbar", [0, 1, 2])
def test_pagedown_step_is_same_inside_and_outside_band(dv, vbar):
    """같은 칸·같은 스크롤에서 PageDown 은 밴드 열이든 일반 열이든 같은 행에 착지한다."""
    tbl = _shrink(dv)
    start = _body_row(dv, 0)
    band = _page_land(dv, tbl, start, 0, Qt.Key_PageDown, vbar)
    plain = _page_land(dv, tbl, start, KEY_COL + 1, Qt.Key_PageDown, vbar)
    assert band == plain, f"vbar={vbar}: 밴드 {band} vs 일반 {plain}"


@pytest.mark.parametrize("vbar", [0, 1, 2])
def test_pageup_step_is_same_inside_and_outside_band(dv, vbar):
    tbl = _shrink(dv)
    start = _body_row(dv, -1)
    band = _page_land(dv, tbl, start, 0, Qt.Key_PageUp, vbar)
    plain = _page_land(dv, tbl, start, KEY_COL + 1, Qt.Key_PageUp, vbar)
    assert band == plain, f"vbar={vbar}: 밴드 {band} vs 일반 {plain}"


def test_page_move_never_lands_on_filtered_row(dv):
    """페이지 이동도 '변경점만 보기'로 숨은 행은 밟지 않는다(키 행만 예외)."""
    tbl = _shrink(dv)
    for key in (Qt.Key_PageDown, Qt.Key_PageUp):
        tbl._move_current_cell(_body_row(dv, 1), 0)
        QApplication.instance().processEvents()
        for _ in range(6):
            _press(tbl, key)
            r, _c = tbl._current_cell()
            assert (not tbl.isRowHidden(r)) or r == 0,                 f"{key}: 볼 수 없는 행 {r} 에 착지했다"


def test_shift_pagedown_keeps_anchor_and_band(dv):
    """Shift+PageDown 도 앵커(A1)와 밴드 열을 지킨다."""
    tbl = _shrink(dv)
    _select_a1_g1(dv)
    _press(tbl, Qt.Key_PageDown, Qt.ShiftModifier)
    top, left, bot, right = _box(tbl)
    assert (top, left, right) == (0, 0, 6), "앵커/밴드 열이 어긋났다"
    assert bot > 0, "아래로 확장되지 않았다"
    assert not tbl.isRowHidden(bot), "확장 끝이 숨은 행에 놓였다"


def test_shift_pageup_returns_to_the_anchor(dv):
    """PageDown 으로 늘린 뒤 PageUp 이면 앵커 쪽으로 되접힌다."""
    tbl = _shrink(dv)
    _select_a1_g1(dv)
    _press(tbl, Qt.Key_PageDown, Qt.ShiftModifier)
    grown = _box(tbl)[2]
    _press(tbl, Qt.Key_PageUp, Qt.ShiftModifier)
    top, left, bot, right = _box(tbl)
    assert (top, left, right) == (0, 0, 6)
    assert bot <= grown


def test_shift_end_extends_to_the_last_data_column(dv):
    """Shift+End — 선택은 마지막 **데이터** 열까지(빈 여유 열은 _select_range 가 잘라낸다).

    커서는 Shift 없는 End 와 같은 칸(격자 마지막 열)에 놓인다 — 두 경로가 어긋나지
    않도록 목적지를 Qt 와 같게 잡았다. Ctrl+Shift+End 도 예전부터 같은 모양이다.
    """
    tbl = dv.panel_a.table
    m = tbl._model
    _select_a1_g1(dv)
    _press(tbl, Qt.Key_End, Qt.ShiftModifier)
    assert _box(tbl) == (0, 0, 0, m.data_cols - 1)
    assert tbl._current_cell() == (0, tbl.columnCount() - 1)


def test_end_from_band_reaches_last_column(dv):
    """Shift 없는 End 도 밴드 열에서 나갈 수 있다(결과 열은 보이므로 원래도 됐다)."""
    r = _body_row(dv)
    assert _move(dv, r, 0, Qt.Key_End) == (r, dv.panel_a.table.columnCount() - 1)


# ── Enter/Return: 아래 칸으로 ─────────────────────────────────────────────────
# Enter 는 '한 행 아래'(cur_r + 1)로 옮기고 있었다. '변경점만 보기'가 켜진 기본 상태에선
# 그 칸이 대개 숨은 행이라 커서가 화면에서 사라지고, 그 자리에서 Alt+→ 를 누르면 보이지도
# 않는 행이 병합 준비된다. ↓ 와 같은 걸음(다음 '갈 수 있는' 행)이어야 한다.
# (실파일 실측, 필터 ON: 80행에서 Enter 6번 → 81~86 전부 숨은 행. ↓ 는 102·123·180….)

def _enter_walk(dv, start, col, key=Qt.Key_Return, mods=Qt.NoModifier, n=6):
    """start@col 에서 key 를 n 번 눌러 착지 칸을 순서대로 모은다."""
    tbl = dv.panel_a.table
    tbl._move_current_cell(start, col)
    QApplication.instance().processEvents()
    out = []
    for _ in range(n):
        _press(tbl, key, mods)
        out.append(tbl._current_cell())
    return out


@pytest.mark.parametrize("col", [0, KEY_COL, KEY_COL + 2])
def test_enter_never_lands_on_filtered_row(dv, col):
    """밴드 열이든 일반 열이든 Enter 는 볼 수 없는 행에 착지하지 않는다."""
    tbl = dv.panel_a.table
    assert [r for r in range(1, len(dv._diff_matrix)) if tbl.isRowHidden(r)],         "전제: 필터로 숨은 본문 행이 있어야 의미 있는 테스트"
    for r, c in _enter_walk(dv, _body_row(dv), col):
        assert c == col, "세로 이동인데 열이 바뀌었다"
        assert not tbl.isRowHidden(r), f"볼 수 없는 행 {r} 에 착지했다"


@pytest.mark.parametrize("col", [0, KEY_COL + 2])
def test_enter_walks_exactly_like_down_arrow(dv, col):
    """Enter 의 걸음은 ↓ 와 같다 — 두 경로가 갈라지면 그게 곧 버그였다."""
    assert (_enter_walk(dv, _body_row(dv), col)
            == _enter_walk(dv, _body_row(dv), col, Qt.Key_Down))


def test_shift_enter_also_skips_filtered_rows(dv):
    """Shift+Enter 도 같은 경로를 탄다(방향은 예전처럼 아래)."""
    tbl = dv.panel_a.table
    for r, _c in _enter_walk(dv, _body_row(dv), KEY_COL + 2, mods=Qt.ShiftModifier):
        assert not tbl.isRowHidden(r), f"볼 수 없는 행 {r} 에 착지했다"


def test_enter_from_key_row_reaches_first_body_row(dv):
    """고정 키 행(0)에서 Enter → 첫 보이는 본문 행."""
    assert _enter_walk(dv, 0, KEY_COL + 2, n=1) == [(_body_row(dv), KEY_COL + 2)]


def test_enter_stops_at_the_last_row(dv):
    """격자 마지막 행에서 Enter 는 제자리 — 갈 칸이 없다."""
    last = dv.panel_a.table.rowCount() - 1
    assert _enter_walk(dv, last, KEY_COL + 2, n=3) == [(last, KEY_COL + 2)] * 3


def test_enter_moves_one_row_without_the_filter(dv):
    """필터를 끄면 예전처럼 정확히 한 행 아래 — 걸음의 정의만 바뀌었을 뿐이다."""
    dv.diff_only_btn.setChecked(False)
    QApplication.instance().processEvents()
    r = _body_row(dv, 2)
    c = KEY_COL + 2
    assert _enter_walk(dv, r, c, n=3) == [(r + 1, c), (r + 2, c), (r + 3, c)]


# ── 밴드에서 나오기 — 원래 되던 동작 비회귀 ───────────────────────────────────

def test_right_leaves_band(dv):
    r = _body_row(dv)
    assert _move(dv, r, KEY_COL, Qt.Key_Right) == (r, KEY_COL + 1)


def test_down_leaves_key_row(dv):
    first = _body_row(dv)
    assert _move(dv, 0, 5, Qt.Key_Down) == (first, 5)


def test_ctrl_left_still_jumps_to_first_column(dv):
    r = _body_row(dv)
    assert _move(dv, r, KEY_COL + 1, Qt.Key_Left, Qt.ControlModifier) == (r, 0)


# ── 핵심 비회귀: 필터로 숨은 행은 여전히 밟을 수 없다 ─────────────────────────

def test_filtered_rows_are_never_stepped_on(dv):
    """'변경점만 보기'로 숨은 행에는 절대 착지하지 않는다(키 행만 예외).

    밴드를 '보이는 것'으로 치는 수정이 필터 숨김까지 뚫어버리면, 사용자가 볼 수 없는
    행에 커서가 놓인다. 위/아래로 끝까지 훑으며 모든 착지 지점을 검사한다.
    """
    tbl = dv.panel_a.table
    assert dv._diff_only, "전제: 비교 직후 '변경점만 보기'가 켜져 있다"
    hidden_body = [r for r in range(1, len(dv._diff_matrix)) if tbl.isRowHidden(r)]
    assert hidden_body, "전제: 필터로 숨은 본문 행이 있어야 의미 있는 테스트"

    for key in (Qt.Key_Up, Qt.Key_Down):
        tbl._move_current_cell(_body_row(dv, 2), 5)
        QApplication.instance().processEvents()
        for _ in range(30):
            _press(tbl, key)
            r, _c = tbl._current_cell()
            assert (not tbl.isRowHidden(r)) or r == 0, \
                f"{key}: 볼 수 없는 행 {r} 에 착지했다"


def test_excluded_columns_are_not_skipped(dv):
    """'변경 검사에서 제외'는 열을 숨기지 않으므로 이동에 영향이 없어야 한다.

    (변경이 실린 열 5가 아니라 열 6을 제외한다 — 5를 제외하면 전 행이 '변경 없음'이 돼
    필터가 행을 다 숨기고, 그러면 열 이동이 아니라 행 가시성을 재는 테스트가 된다.)
    """
    dv._on_columns_exclude_set([KEY_COL + 3], True)
    QApplication.instance().processEvents()
    r = _body_row(dv)          # 제외 후 다시 계산 — 필터가 행 구성을 바꿀 수 있다
    assert not dv.panel_a.table.isColumnHidden(KEY_COL + 3)
    assert _move(dv, r, KEY_COL + 4, Qt.Key_Left) == (r, KEY_COL + 3)


# ── 틀 고정이 없으면 Qt 기본 그대로 ───────────────────────────────────────────

def test_without_freeze_behaves_as_before(dv):
    """freeze 를 끄면 숨은 열이 없어 개입 조건(_frozen_count=0)이 성립하지 않는다."""
    r = _body_row(dv)
    dv._freeze["a"].clear()
    QApplication.instance().processEvents()
    tbl = dv.panel_a.table
    from excelmerge.widgets import AX_COL
    assert not tbl.isColumnHidden(KEY_COL)
    assert tbl._frozen_count(AX_COL) == 0, "freeze 해제 후엔 개입 조건이 꺼져야 한다"
    assert _move(dv, r, KEY_COL + 1, Qt.Key_Left) == (r, KEY_COL)
