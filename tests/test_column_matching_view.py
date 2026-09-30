# -*- coding: utf-8 -*-
"""열 매칭 2단계 — 화면에 실제로 켠다.

1단계(tests/test_column_matching.py)는 엔진만이었다. 여기서는 워커가 col_meta 를 만들어
화면까지 흘려보내고, 헤더가 **side 별 자기 열 문자**를 보여 주는 것을 고정한다. 세로
헤더가 row_meta 로 side 별 원본 행 번호를 보여 주는 것과 같은 규칙이다.

★ 저장은 아직 막혀 있다. 저장 경로는 여전히 화면 열 번호를 파일 열로 쓰기 때문에,
   맞춘 상태로 저장하면 A 의 값이 B 의 엉뚱한 열에 들어간다 — 이 기능이 막으려던 사고를
   우리가 내는 꼴이다. 3단계에서 저장이 col_meta 를 쓰게 되면 그 차단은 사라진다.
"""
import pytest
from PyQt5.QtCore import Qt

from excelmerge.theme import HEADER_ONESIDE_BG

# A: ID | NAME | VALUE        B: ID | GRADE | NAME | VALUE   (GRADE 가 중간에 끼었다)
A_DATA = [["ID", "NAME", "VALUE"], ["k1", "칼", "10"], ["k2", "방패", "20"]]
B_DATA = [["ID", "GRADE", "NAME", "VALUE"],
          ["k1", "A", "칼", "10"], ["k2", "S", "방패", "20"]]
SAME_A = [["ID", "NAME"], ["k1", "칼"], ["k2", "방패"]]
SAME_B = [["ID", "NAME"], ["k1", "칼"], ["k2", "몽둥이"]]


@pytest.fixture
def make_view(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    monkeypatch.setenv("APPDATA", str(tmp_path))
    wins = []

    def _make(a=A_DATA, b=B_DATA):
        win = MainWindow()
        win.show()
        wins.append(win)
        view = win.tabs.currentWidget()
        view.panel_a.set_path(str(tmp_path / "a.xlsx"))
        view.panel_b.set_path(str(tmp_path / "b.xlsx"))
        view._on_loaded(a, b)
        w = getattr(view, "_diff_worker", None)
        if w is not None:
            w.wait(8000)
        for _ in range(20):
            qapp.processEvents()
        return view

    yield _make
    for win in wins:
        win.close()
        win.deleteLater()
    qapp.processEvents()


def _letters(view, side, n):
    m = (view.panel_a if side == "a" else view.panel_b).table.model()
    return [m.headerData(c, Qt.Horizontal, Qt.DisplayRole) for c in range(n)]


# ── 워커 → 화면 배선 ─────────────────────────────────────────────────────────

def test_col_meta_reaches_the_view(make_view):
    v = make_view()
    assert v._diff_col_meta == [(0, 0), (None, 1), (1, 2), (2, 3)]


def test_the_matrix_width_follows_col_meta(make_view):
    v = make_view()
    assert len(v._diff_matrix[0]) == 4


def test_matching_is_off_when_headers_are_identical(make_view):
    """맞추나 마나 같은 경우에는 화면에 아무 표시도 뜨면 안 된다."""
    v = make_view(SAME_A, SAME_B)
    assert not v._columns_are_shifted()
    assert "이름으로 맞춤" not in v.status.currentMessage()


def test_a_new_comparison_clears_the_old_col_meta(make_view, qapp):
    v = make_view()
    assert v._diff_col_meta is not None
    v._run_compare()                       # 경로가 가짜라 로드가 실패한다 — 그래도 비워야 한다
    for _ in range(10):
        qapp.processEvents()
    assert v._diff_col_meta is None, "이전 비교의 열 매핑이 남았다"


# ── 헤더: side 별 자기 열 문자 ───────────────────────────────────────────────

def test_each_panel_shows_its_own_column_letters(make_view):
    v = make_view()
    assert _letters(v, "a", 4) == ["A", "-", "B", "C"], "A 에 없는 열은 '-' 여야 한다"
    assert _letters(v, "b", 4) == ["A", "B", "C", "D"]


def test_letters_are_plain_when_not_matched(make_view):
    v = make_view(SAME_A, SAME_B)
    assert _letters(v, "a", 2) == ["A", "B"]
    assert _letters(v, "b", 2) == ["A", "B"]


def test_own_col_maps_to_the_file_column(make_view):
    v = make_view()
    ma, mb = v.panel_a.table.model(), v.panel_b.table.model()
    assert [ma.own_col(c) for c in range(4)] == [0, None, 1, 2]
    assert [mb.own_col(c) for c in range(4)] == [0, 1, 2, 3]


# ── 한쪽에만 있는 열 표시 ────────────────────────────────────────────────────

def test_one_sided_column_is_marked_on_both_panels(make_view):
    v = make_view()
    for panel in (v.panel_a, v.panel_b):
        m = panel.table.model()
        assert [c for c in range(4) if m.one_sided_col(c)] == [1]


def test_one_sided_header_is_tinted(make_view):
    m = make_view().panel_a.table.model()
    assert m.headerData(1, Qt.Horizontal, Qt.BackgroundRole) == HEADER_ONESIDE_BG
    assert m.headerData(2, Qt.Horizontal, Qt.BackgroundRole) != HEADER_ONESIDE_BG


def test_one_sided_header_explains_itself(make_view):
    m = make_view().panel_a.table.model()
    tip = m.headerData(1, Qt.Horizontal, Qt.ToolTipRole)
    assert tip and "한쪽 파일에만" in tip
    assert m.headerData(2, Qt.Horizontal, Qt.ToolTipRole) is None


def test_the_key_column_keeps_its_own_colour(make_view):
    """키 열이 우선 — 키 표시를 한쪽 전용 색이 덮으면 안 된다."""
    from excelmerge.theme import HEADER_KEY_BG
    m = make_view().panel_a.table.model()
    assert m.headerData(0, Qt.Horizontal, Qt.BackgroundRole) == HEADER_KEY_BG


# ── 사용자에게 알리기 ────────────────────────────────────────────────────────

def test_the_status_says_columns_were_matched(make_view):
    msg = make_view().status.currentMessage()
    assert "열을 이름으로 맞춤" in msg, msg
    assert "B 전용 1열" in msg, msg


def test_a_only_columns_are_counted_too(make_view):
    a = [["ID", "MEMO", "VALUE"], ["k1", "메모", "10"]]
    b = [["ID", "VALUE"], ["k1", "10"]]
    msg = make_view(a, b).status.currentMessage()
    assert "A 전용 1열" in msg, msg


# ── 저장 차단 (3단계에서 사라진다) ──────────────────────────────────────────

def test_saving_is_blocked_while_columns_are_shifted(make_view, monkeypatch):
    """저장 경로가 아직 화면 열 번호를 파일 열로 쓴다 — 그대로 쓰면 엉뚱한 열에 들어간다."""
    from excelmerge import diff_view as dv_mod
    from excelmerge.constants import DIR_A2B
    v = make_view()
    warned = []
    monkeypatch.setattr(dv_mod.QMessageBox, "warning",
                        staticmethod(lambda *a, **k: warned.append(a[2] if len(a) > 2 else "")))
    started = []
    monkeypatch.setattr(dv_mod, "StagedMergeWorker",
                        lambda *a, **k: started.append(1))
    v._staged[(1, 2)] = DIR_A2B
    v._save_staged("b")
    assert started == [], "열이 어긋난 채로 저장이 시작됐다"
    assert warned and "열 구성" in warned[0], warned


def test_saving_is_not_blocked_when_columns_line_up(make_view, monkeypatch):
    """맞출 필요가 없던 비교까지 막으면 안 된다."""
    from excelmerge import diff_view as dv_mod
    v = make_view(SAME_A, SAME_B)
    blocked = []
    monkeypatch.setattr(dv_mod.QMessageBox, "warning",
                        staticmethod(lambda *a, **k: blocked.append(1)))
    v._save_staged("b")          # staged 가 없어 조용히 반환 — 차단 경고는 없어야 한다
    assert blocked == []
