# -*- coding: utf-8 -*-
"""열 매칭 5단계 — 파일별 기억을 **헤더 이름**으로 남긴다.

4단계까지로 비교·표시·저장·키가 col_meta 를 쓰게 됐다. 그런데 파일별 기억(키 위치,
검사 제외 열)은 여전히 **화면 열 번호**였다. 열이 하나만 끼어도 그 번호는 다른 열을
가리킨다 — 오늘 만든 '제외 열 기억' 이 내일 엉뚱한 열을 회색칠한다는 뜻이다.

여기서 고정하는 것:
  1. 열이 끼어도 기억이 **같은 열**을 따라간다(이름으로 찾으므로).
  2. 옛 설정(숫자 목록/2원소 키)도 그대로 읽힌다 — 기존 사용자 설정이 깨지지 않는다.
  3. 이름을 못 찾으면 조용히 버린다 — 없는 열을 회색칠하지 않는다.
"""
import pytest

from excelmerge.prefs import (load_last_excluded, load_last_key,
                              load_last_key_name, save_last_excluded,
                              save_last_key)

#  A                              B (GRADE 가 중간에 끼었다)
A1 = [["ID", "NAME", "#주석", "VALUE"], ["k1", "칼", "메모", "10"]]
B1 = [["ID", "GRADE", "NAME", "#주석", "VALUE"], ["k1", "A", "검", "메모2", "10"]]
#  나중에 A 에도 GRADE 가 생긴 모습 — 화면 열 번호가 통째로 밀린다
A2 = [["ID", "GRADE", "NAME", "#주석", "VALUE"], ["k1", "A", "칼", "메모", "10"]]


# ── 저장소 계층 ──────────────────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def _isolated_appdata(monkeypatch, tmp_path):
    monkeypatch.setenv("APPDATA", str(tmp_path))


def test_names_round_trip(tmp_path):
    p = str(tmp_path / "a.xlsx")
    save_last_excluded(p, ["#주석", "en"])
    assert load_last_excluded(p) == ["#주석", "en"]


def test_the_old_number_format_still_loads(tmp_path):
    """기존 사용자 설정이 깨지면 안 된다."""
    p = str(tmp_path / "a.xlsx")
    save_last_excluded(p, [3, 2])                 # 옛 형식
    assert load_last_excluded(p) == [2, 3]


def test_nameless_columns_are_not_remembered(tmp_path):
    """헤더가 빈 열은 다음에 찾을 방법이 없다 — 기억하지 않는다."""
    p = str(tmp_path / "a.xlsx")
    save_last_excluded(p, ["#주석", ""])
    assert load_last_excluded(p) == ["#주석"]


def test_the_key_name_is_stored_alongside_the_number(tmp_path):
    p = str(tmp_path / "a.xlsx")
    save_last_key(p, 0, 2, "NAME")
    assert load_last_key(p) == (0, 2)             # 번호도 그대로 읽힌다(대비책)
    assert load_last_key_name(p) == "NAME"


def test_an_old_two_element_key_has_no_name(tmp_path):
    import excelmerge.prefs as prefs
    p = str(tmp_path / "a.xlsx")
    prefs._write_prefs({"last_keys": {__import__("os").path.abspath(p): [0, 2]}})
    assert load_last_key(p) == (0, 2)
    assert load_last_key_name(p) is None


# ── 화면 연결 ────────────────────────────────────────────────────────────────

@pytest.fixture
def make_view(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    wins = []

    def _make(a, b, path_a="a.xlsx", path_b="b.xlsx"):
        win = MainWindow()
        win.show()
        wins.append(win)
        view = win.tabs.currentWidget()
        view.panel_a.set_path(str(tmp_path / path_a))
        view.panel_b.set_path(str(tmp_path / path_b))
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


def _names(view):
    return view._display_header_names(view._diff_col_meta)


# ── 1. 열이 끼어도 따라간다 ─────────────────────────────────────────────────

def test_display_names_follow_the_matched_columns(make_view):
    v = make_view(A1, B1)
    assert v._diff_col_meta == [(0, 0), (None, 1), (1, 2), (2, 3), (3, 4)]
    assert _names(v) == ["ID", "GRADE", "NAME", "#주석", "VALUE"]


def test_the_exclusion_follows_a_column_insertion(make_view, qapp):
    """A1↔B1 에서 '#주석'은 화면 3열. A 에도 GRADE 가 생기면 여전히 3열이지만,
    파일 안에서는 자리가 달라진다 — 이름으로 기억해야 같은 열을 다시 잡는다."""
    v = make_view(A1, B1)
    i = _names(v).index("#주석")
    v._on_columns_exclude_set([i], True)
    for _ in range(10):
        qapp.processEvents()
    assert load_last_excluded(v.panel_a.get_path()) == ["#주석"]

    again = make_view(A2, B1)                  # A 에도 GRADE 가 생겼다
    j = _names(again).index("#주석")
    assert again._excluded_cols == {j}, (
        f"이름으로 다시 잡지 못했다: {again._excluded_cols} / 기대 {j}")


def test_the_key_follows_a_column_insertion(make_view, qapp):
    """키를 NAME 으로 잡아 두면, 앞에 열이 끼어도 NAME 을 계속 키로 쓴다."""
    v = make_view(A1, B1)
    name_col = _names(v).index("NAME")
    v._on_key_col_changed(name_col)
    w = getattr(v, "_diff_worker", None)
    if w is not None:
        w.wait(8000)
    for _ in range(20):
        qapp.processEvents()
    assert load_last_key_name(v.panel_a.get_path()) == "NAME"

    # A 앞쪽에 열이 하나 더 생겨 화면 열 번호가 밀린 상태로 다시 연다.
    #   화면: EXTRA · ID · GRADE · NAME · #주석 · VALUE   → NAME 은 이제 3열
    #   기억한 번호(2)를 그대로 쓰면 GRADE(한쪽 전용)를 키로 잡는다. 그러면
    #   usable_col_meta 가 열 매칭을 통째로 꺼 버린다 — 이름으로 찾아야 한다.
    a3 = [["EXTRA", "ID", "NAME", "#주석", "VALUE"], ["x", "k1", "칼", "메모", "10"]]
    again = make_view(a3, B1)
    assert again._diff_col_meta is not None, (
        "키를 놓쳐 한쪽 전용 열을 잡는 바람에 열 매칭이 꺼졌다")
    assert again._key_col == 3, f"키가 밀렸다: key={again._key_col}"
    assert _names(again)[again._key_col] == "NAME", _names(again)


def test_a_lost_name_is_dropped_quietly(make_view, qapp):
    """기억한 이름이 이 파일에 없으면 버린다 — 없는 열을 회색칠하지 않는다."""
    v = make_view(A1, B1)
    save_last_excluded(v.panel_a.get_path(), ["없는열이름"])
    again = make_view(A1, B1)
    assert again._excluded_cols == set(), again._excluded_cols


def test_the_old_number_memory_still_applies(make_view):
    """옛 설정(숫자)으로 저장돼 있어도 계속 동작한다."""
    v = make_view(A1, B1)
    save_last_excluded(v.panel_a.get_path(), [3])          # 옛 형식 — 화면 3열
    again = make_view(A1, B1)
    assert again._excluded_cols == {3}, again._excluded_cols


# ── 2. 화면과 계산이 같은 col_meta 를 쓴다 ──────────────────────────────────

def test_the_view_and_the_worker_share_one_col_meta(make_view):
    """예전엔 워커가 따로 계산했다 — 판정이 키에 달려 있어 어긋날 여지가 있었다.

    이제 DiffView 가 한 번 계산해 워커에 넘긴다. 돌아온 값이 넘긴 값과 같아야 한다.
    """
    v = make_view(A1, B1)
    assert v._diff_col_meta == v._effective_col_meta()
    assert v._raw_col_match == [(0, 0), (None, 1), (1, 2), (2, 3), (3, 4)]
