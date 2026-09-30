# -*- coding: utf-8 -*-
"""검사 제외 열을 파일별로 기억한다 — 열 때마다 같은 열을 다시 고르지 않도록.

기억은 **헤더 이름**으로 남는다(열이 하나 끼면 번호는 다른 열을 가리키므로).
옛 숫자 형식도 계속 읽힌다 — 아래 저장소 계층 테스트가 그 호환을 함께 본다.

어떤 열을 안 볼지는 그 테이블의 성질이지 그때의 기분이 아니다. 주석 열(#Desc 등)이나
현지화 열은 늘 같은 것을 뺀다. 그런데 제외는 비교할 때마다 초기화돼서, 파일을 열 때마다
헤더를 우클릭해 같은 열을 다시 골라야 했다. 키 위치를 파일별로 기억하게 만든 것과
같은 이유다(tests/test_key_memory.py).

여기서 고정하는 것은 '기억한다' 보다 **언제 기억하지 않는가 / 무엇을 걸러내는가** 다.
"""
import pytest

from excelmerge.prefs import load_last_excluded, save_last_excluded

HEADER = ["ID", "NAME", "#주석", "en", "VALUE"]
A_DATA = [HEADER] + [[f"k{i}", f"a{i}", f"메모{i}", f"en{i}", str(i)] for i in range(1, 5)]
B_DATA = [HEADER] + [[f"k{i}", f"B{i}", f"다른메모{i}", f"EN{i}", str(i)] for i in range(1, 5)]


# ── 저장소 ───────────────────────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def _isolated_appdata(monkeypatch, tmp_path):
    """실제 사용자 설정을 건드리지 않는다."""
    monkeypatch.setenv("APPDATA", str(tmp_path))


def test_nothing_remembered_is_none(tmp_path):
    assert load_last_excluded(str(tmp_path / "never_seen.xlsx")) is None


def test_roundtrip(tmp_path):
    p = str(tmp_path / "a.xlsx")
    save_last_excluded(p, [3, 2])
    assert load_last_excluded(p) == [2, 3], "정렬해서 돌려줘야 한다"


def test_empty_is_a_choice_not_absence(tmp_path):
    """제외를 전부 해제한 것도 사용자의 선택이다 — 다음에 열 때 되살아나면 안 된다."""
    p = str(tmp_path / "a.xlsx")
    save_last_excluded(p, [2])
    save_last_excluded(p, [])
    assert load_last_excluded(p) == [], "빈 목록이 '기억 없음'으로 퇴화했다"
    assert load_last_excluded(p) is not None


def test_each_file_keeps_its_own(tmp_path):
    a, b = str(tmp_path / "a.xlsx"), str(tmp_path / "b.xlsx")
    save_last_excluded(a, [1])
    save_last_excluded(b, [4])
    assert load_last_excluded(a) == [1] and load_last_excluded(b) == [4]


def test_same_file_different_spelling_is_the_same_file(tmp_path):
    """상대 경로로 들어와도 같은 파일이어야 한다(키 기억과 동일 규칙)."""
    import os
    p = tmp_path / "a.xlsx"
    p.write_text("x", encoding="utf-8")
    save_last_excluded(str(p), [2])
    cwd = os.getcwd()
    os.chdir(str(tmp_path))
    try:
        assert load_last_excluded("a.xlsx") == [2]
    finally:
        os.chdir(cwd)


def test_garbage_in_the_file_does_not_crash(tmp_path, monkeypatch):
    import excelmerge.prefs as prefs
    monkeypatch.setattr(prefs, "_read_prefs",
                        lambda: {"last_excluded": {"x": "이건 목록이 아니다"}})
    assert prefs.load_last_excluded("x") is None


def test_saving_junk_is_ignored(tmp_path):
    p = str(tmp_path / "a.xlsx")
    save_last_excluded(p, ["둘", None])          # 예외 없이 무시
    assert load_last_excluded(p) is None


def test_negative_columns_are_dropped(tmp_path):
    p = str(tmp_path / "a.xlsx")
    save_last_excluded(p, [-1, 2])
    assert load_last_excluded(p) == [2]


# ── 화면 연결 ────────────────────────────────────────────────────────────────

@pytest.fixture
def make_view(qapp, monkeypatch, tmp_path):
    """파일 경로를 붙인 DiffView 를 만들어 준다(여러 번 호출 가능 — '다시 열기' 재현용)."""
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    wins = []

    def _make(path_a="a.xlsx", path_b="b.xlsx", a=A_DATA, b=B_DATA):
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


def _exclude(view, cols, qapp):
    view._on_columns_exclude_set(list(cols), True)
    for _ in range(10):
        qapp.processEvents()


def test_exclusion_survives_reopening_the_file(make_view, qapp):
    view = make_view()
    assert view._excluded_cols == set(), "전제: 처음엔 제외가 없다"
    _exclude(view, [2, 3], qapp)
    assert view._excluded_cols == {2, 3}

    again = make_view()                       # 같은 경로로 다시 열기
    assert again._excluded_cols == {2, 3}, "다시 열었더니 제외가 사라졌다"


def test_another_file_is_not_affected(make_view, qapp):
    _exclude(make_view(), [2], qapp)
    other = make_view(path_a="other_a.xlsx", path_b="other_b.xlsx")
    assert other._excluded_cols == set(), "다른 파일에 남의 제외가 묻었다"


def test_clearing_is_remembered_too(make_view, qapp):
    view = make_view()
    _exclude(view, [2], qapp)
    view._on_columns_exclude_set([2], False)  # 해제
    for _ in range(10):
        qapp.processEvents()
    again = make_view()
    assert again._excluded_cols == set(), "해제했는데 제외가 되살아났다"


def test_columns_outside_the_sheet_are_dropped(make_view, qapp):
    """넓은 시트에서 제외한 열이 좁은 시트에 그대로 오면 없는 열을 회색칠한다."""
    _exclude(make_view(), [4], qapp)
    narrow_header = ["ID", "NAME"]
    narrow = [narrow_header] + [[f"k{i}", f"x{i}"] for i in range(1, 4)]
    narrow_b = [narrow_header] + [[f"k{i}", f"y{i}"] for i in range(1, 4)]
    view = make_view(a=narrow, b=narrow_b)
    assert view._excluded_cols == set(), f"시트 폭 밖의 열이 남았다: {view._excluded_cols}"


def test_the_key_column_is_never_excluded(make_view, qapp, monkeypatch):
    """키 열은 제외할 수 없다 — 기억에 남아 있어도 걸러야 한다."""
    view = make_view()
    save_last_excluded(view.panel_a.get_path(), [0, 2])   # 0 = 키 열
    again = make_view()
    assert 0 not in again._excluded_cols, "키 열이 제외됐다"
    assert 2 in again._excluded_cols, "나머지는 살아야 한다"


def test_a_key_change_does_not_erase_the_memory(make_view, qapp):
    """키를 바꾸면 화면의 제외는 초기화된다(기존 동작) — 그렇다고 기억까지 지우면
    안 된다. 키를 한 번 건드린 것만으로 애써 잡아 둔 제외가 날아간다."""
    view = make_view()
    _exclude(view, [2], qapp)
    view._on_key_col_changed(1)
    w = getattr(view, "_diff_worker", None)
    if w is not None:
        w.wait(8000)
    for _ in range(20):
        qapp.processEvents()
    assert load_last_excluded(view.panel_a.get_path()) == ["#주석"], (
        "키 변경의 자동 초기화가 기억까지 덮어썼다")


def test_restored_exclusion_is_shown_in_the_status(make_view, qapp):
    """회색 열은 변경 셀 수를 줄인다 — 왜 줄었는지 말해 주지 않으면 조용히 결과가 바뀐다."""
    _exclude(make_view(), [2, 3], qapp)
    view = make_view()
    msg = view.status.currentMessage()
    assert "검사 제외 열" in msg and "C" in msg and "D" in msg, msg


def test_restored_exclusion_actually_affects_the_diff(make_view, qapp):
    """복원이 표시만 바꾸는 게 아니라 비교 결과에도 반영돼야 한다."""
    before = make_view()._count_changed()
    _exclude(make_view(), [2], qapp)          # '#주석' 열 — 전 행이 다르다
    after = make_view()._count_changed()
    assert after < before, f"제외가 복원됐는데 변경 수가 그대로다: {before} → {after}"
