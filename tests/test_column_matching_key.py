# -*- coding: utf-8 -*-
"""열 매칭 4단계 — 키 열도 side 별로 읽는다.

1~3단계로 비교·표시·저장은 col_meta 를 쓰게 됐다. 그런데 **키 열**을 아직 한 값으로
쓰는 자리가 남아 있었다. 열을 이름으로 맞추면 두 파일의 키 열 번호가 다를 수 있으므로,
그 자리들은 전부 거짓말을 한다.

여기서 고정하는 것 셋:
  1. '비교에서 빠진 행' 집계를 side 별 키 열로 센다 — 안 그러면 멀쩡한 비교에
     '수천 행 제외됨' 이 뜨거나, 진짜 드롭을 놓친다.
  2. 상태바의 키 위치는 두 파일이 다르면 둘 다 보여 준다.
  3. 한쪽에만 있는 열은 **키로 고를 수 없다** — 고르면 비교가 통째로 위치 기준으로
     되돌아가는데 사용자는 이유를 알 수 없다.
"""
import pytest

from excelmerge.diff_engine import (count_dropped_key_rows, compute_diff,
                                    match_columns)

# 두 파일에서 키(ID)의 자리가 다르다. A 는 0열, B 는 1열.
A_KEY0 = [["ID", "NAME"], ["k1", "칼"], ["k2", "방패"], ["k3", "활"]]
B_KEY1 = [["NAME", "ID"], ["칼", "k1"], ["방패", "k2"], ["활", "k3"]]


# ── 1. 빠진 행 집계 ──────────────────────────────────────────────────────────

def test_dropped_count_uses_each_files_own_key_column():
    """맞추지 않고 세면 B 의 NAME 열을 키로 착각한다 — 여기선 값이 유니크해 0 이지만,
    아래 테스트처럼 중복이 있으면 숫자가 완전히 달라진다."""
    cm = match_columns(A_KEY0, B_KEY1)
    assert cm == [(0, 1), (1, 0)]
    assert count_dropped_key_rows(A_KEY0, B_KEY1, 0, 0, cm) == 0


def test_without_col_meta_the_count_is_wrong_on_shifted_files():
    """맞추지 않으면 B 의 엉뚱한 열을 센다 — 이 차이가 이번 수정의 이유다."""
    a = [["ID", "GRP"], ["k1", "g"], ["k2", "g"], ["k3", "g"]]
    b = [["GRP", "ID"], ["g", "k1"], ["g", "k2"], ["g", "k3"]]
    cm = match_columns(a, b)
    assert cm == [(0, 1), (1, 0)]
    # 맞춰서 세면 양쪽 모두 ID 를 보므로 드롭 0
    assert count_dropped_key_rows(a, b, 0, 0, cm) == 0
    # 안 맞추면 B 의 0열(GRP, 전부 'g')을 키로 봐서 2행이 중복으로 잡힌다
    assert count_dropped_key_rows(a, b, 0, 0, None) == 2


def test_real_duplicates_are_still_counted():
    a = [["ID", "NAME"], ["k1", "x"], ["k1", "y"], ["k2", "z"]]
    b = [["NAME", "ID"], ["x", "k1"], ["y", "k9"], ["z", "k2"]]
    cm = match_columns(a, b)
    assert count_dropped_key_rows(a, b, 0, 0, cm) == 1, "A 의 중복 k1 하나"


def test_row_order_mode_drops_nothing():
    assert count_dropped_key_rows(A_KEY0, B_KEY1, -1, 0, match_columns(A_KEY0, B_KEY1)) == 0


def test_a_key_column_missing_on_one_side_is_not_counted_with_col_meta():
    """키가 한쪽에만 있으면 col_meta 는 쓰이지 않는다(usable_col_meta) — 예전 방식으로 센다."""
    a = [["ONLY_A", "ID"], ["x", "k1"], ["x", "k2"]]
    b = [["ID"], ["k1"], ["k2"]]
    cm = match_columns(a, b)
    assert cm[0] == (0, None)
    # 항등 매핑으로 물러서므로 A 의 0열(ONLY_A, 'x' 중복)에서 1행이 드롭된다
    assert count_dropped_key_rows(a, b, 0, 0, cm) == 1


# ── 화면 ─────────────────────────────────────────────────────────────────────

@pytest.fixture
def make_view(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setenv("APPDATA", str(tmp_path))
    wins = []

    def _make(a, b):
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


# ── 2. 상태바의 키 위치 ─────────────────────────────────────────────────────

def test_the_anchor_shows_both_sides_when_they_differ(make_view):
    v = make_view(A_KEY0, B_KEY1)
    v._on_key_row_changed(0)          # 'load' 가 아닌 재계산 경로 메시지를 띄운다
    label = v._key_anchor_label()
    assert "A:A열" in label and "B:B열" in label, label


def test_the_anchor_stays_simple_when_they_match(make_view):
    same = [["ID", "NAME"], ["k1", "x"]]
    v = make_view(same, [["ID", "NAME"], ["k1", "y"]])
    assert v._key_anchor_label() == "1행 A열", v._key_anchor_label()


def test_the_anchor_says_row_order_without_a_key(make_view):
    v = make_view(A_KEY0, B_KEY1)
    v._key_col = -1
    assert "ROW 순서" in v._key_anchor_label()


# ── 3. 한쪽에만 있는 열은 키가 될 수 없다 ──────────────────────────────────

ONE_SIDED_A = [["ID", "NAME"], ["k1", "칼"]]
ONE_SIDED_B = [["ID", "GRADE", "NAME"], ["k1", "A", "칼"]]


def test_a_one_sided_column_cannot_become_the_key(make_view, monkeypatch):
    from excelmerge import diff_view as dv_mod
    told = []
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: told.append(a[1] if len(a) > 1 else "")))
    v = make_view(ONE_SIDED_A, ONE_SIDED_B)
    assert v._diff_col_meta == [(0, 0), (None, 1), (1, 2)]
    before = v._key_col
    v._on_key_col_changed(1)                 # GRADE — B 에만 있다
    assert v._key_col == before, "한쪽에만 있는 열이 키가 됐다"
    assert told and "키로 쓸 수 없는 열" in told[-1], told


def test_a_common_column_can_still_become_the_key(make_view, qapp):
    v = make_view(ONE_SIDED_A, ONE_SIDED_B)
    v._on_key_col_changed(2)                 # NAME — 양쪽에 있다
    w = getattr(v, "_diff_worker", None)
    if w is not None:
        w.wait(8000)
    for _ in range(20):
        qapp.processEvents()
    assert v._key_col == 2
