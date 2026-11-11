# -*- coding: utf-8 -*-
"""검사 제외 열을 파일별로 기억한다 — 열 때마다 같은 열을 다시 고르지 않도록.

기억은 **헤더 이름 + 그 이름의 몇 번째** 로 남는다(열이 하나 끼면 번호는 다른 열을
가리키므로). 이름 없는 열은 자리로 기억한다. 옛 형식(이름만 · 숫자 목록)도 계속 읽힌다 —
아래 저장소 계층 테스트가 그 호환을 함께 본다.

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


# ── 같은 이름 열 · 이름 없는 열 (실제 xlsx · F5) ────────────────────────────
#
# 기억이 이름 하나뿐이었다. 그래서 '#' 처럼 같은 이름이 여럿이면 **마지막** '#' 를
# 가리켰다 — 첫 '#' 를 빼 두고 F5 를 누르면 다른 '#' 가 회색이 되고, 그 열의 실제
# 변경이 숨었다. 이름 없는 열은 아예 기억하지 않아 F5 한 번에 제외가 풀렸다.
# 이제 (이름, 그 이름의 몇 번째) 로 기억한다 — match_columns 가 같은 이름 열을 짝지을
# 때 쓰는 바로 그 식별자다. 이름 없는 열은 자리로 기억한다.

import os
import time

import openpyxl

from excelmerge import prefs as prefs_mod
from excelmerge.diff_engine import column_tags

DUP = ["ID", "V", "#", "W", "#"]
#  2열('#' 첫째)은 두 행 다 다르고, 4열('#' 둘째)은 한 행만 다르다 — 어느 '#' 가
#  빠졌는지가 변경 셀 수로도 드러나게.
DUP_A = [DUP, ["1", "v", "메모", "w", "x1"], ["2", "v", "메모", "w", "x2"]]
DUP_B = [DUP, ["1", "v", "다른 메모", "w", "x1"], ["2", "v", "또 다른 메모", "w", "X2"]]


def _xlsx(path, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    wb.save(str(path))
    return str(path)


def _wait(pred, timeout=20.0):
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance()
    end = time.time() + timeout
    while time.time() < end:
        app.processEvents()
        if pred():
            return True
        time.sleep(0.01)
    return bool(pred())


def _compared(view):
    return bool(view._diff_matrix) and view.status.currentMessage().startswith("비교 완료")


@pytest.fixture
def win(qapp, monkeypatch):
    from excelmerge import diff_view as dv
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv.QMessageBox, "information", staticmethod(lambda *a, **k: None))
    w = MainWindow()
    w.show()
    yield w
    w.close()
    w.deleteLater()
    qapp.processEvents()


@pytest.fixture
def dup_files(tmp_path):
    return (_xlsx(tmp_path / "dup_a.xlsx", {"S": DUP_A}),
            _xlsx(tmp_path / "dup_b.xlsx", {"S": DUP_B}))


def _open(win, a, b):
    view = win.open_file_compare(a, b)
    assert _wait(lambda: _compared(view)), view.status.currentMessage()
    return view


def _press_f5(view):
    """F5 가 실제로 다시 읽어 다시 비교했는지까지 본다 — 안 그러면 제외가 '남아 있는' 것과
    '되살아난' 것을 구분할 수 없다(다시 비교는 제외를 비우고 시작한다)."""
    gen = view._compare_gen
    view._on_refresh_shortcut()
    assert view._compare_gen == gen + 1, "F5 가 다시 비교하지 않았다"
    assert _wait(lambda: _compared(view)), view.status.currentMessage()


def _raw_put_excluded(path, value):
    """옛 버전이 쓴 그대로의 설정을 심는다 — 새 저장 함수의 정규화를 거치지 않게."""
    prefs_mod._update_prefs(
        lambda c: prefs_mod._lru_put(c, "last_excluded", os.path.abspath(path), value))


def test_column_tags_count_each_name_separately():
    """match_columns 와 제외 기억이 같은 식별자를 쓴다 — 그 규칙 자체를 고정한다."""
    assert column_tags(DUP) == [("ID", 0), ("V", 0), ("#", 0), ("W", 0), ("#", 1)]


@pytest.mark.parametrize("cols", [[2], [4], [2, 4]], ids=["첫#", "둘째#", "둘다"])
def test_duplicate_name_exclusion_survives_f5(win, dup_files, cols):
    view = _open(win, *dup_files)
    view._on_columns_exclude_set(cols, True)
    assert view._excluded_cols == set(cols), "전제: 고른 열이 제외됐다"
    shown = view._count_changed()

    _press_f5(view)
    assert view._excluded_cols == set(cols), (
        f"F5 뒤 다른 '#' 가 제외됐다: {sorted(view._excluded_cols)} / 기대 {cols}")
    assert view._count_changed() == shown, "F5 뒤 보이는 변경 셀 수가 달라졌다"


def test_duplicate_names_are_stored_with_their_occurrence(win, dup_files):
    """첫째는 예전처럼 이름만(옛 버전도 읽는다), 둘째부터는 [이름, 몇 번째]."""
    view = _open(win, *dup_files)
    view._on_columns_exclude_set([2, 4], True)
    for p in dup_files:
        assert load_last_excluded(p) == ["#", ("#", 1)], load_last_excluded(p)


def test_duplicate_name_exclusion_survives_reopening(win, dup_files):
    view = _open(win, *dup_files)
    view._on_columns_exclude_set([2], True)
    win.tabs.removeTab(win.tabs.indexOf(view))
    view.deleteLater()

    again = _open(win, *dup_files)
    assert again is not view
    assert again._excluded_cols == {2}, sorted(again._excluded_cols)


def test_duplicate_name_exclusion_survives_switching_sheets_back(win, tmp_path):
    a = _xlsx(tmp_path / "a.xlsx", {"S1": DUP_A, "S2": DUP_A})
    b = _xlsx(tmp_path / "b.xlsx", {"S1": DUP_B, "S2": DUP_B})
    view = _open(win, a, b)
    view._on_columns_exclude_set([2], True)
    for sheet in ("S2", "S1"):
        view.sheet_tabs.setCurrentIndex(view._sheet_names.index(sheet))
        assert _wait(lambda: view._current_sheet == sheet and _compared(view)), sheet
        assert view._excluded_cols == {2}, (sheet, sorted(view._excluded_cols))


def test_unnamed_column_exclusion_survives_f5(win, tmp_path):
    """헤더가 빈 열 — 찾을 이름이 없으니 자리로 기억한다."""
    hdr = ["ID", "V", None, "W"]
    a = _xlsx(tmp_path / "a.xlsx", {"S": [hdr, ["1", "v", "메모", "w"], ["2", "v", "x", "w"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"S": [hdr, ["1", "v", "다른", "w"], ["2", "v", "x", "W"]]})
    view = _open(win, a, b)
    view._on_columns_exclude_set([2], True)
    assert load_last_excluded(a) == [("", 2)], load_last_excluded(a)

    _press_f5(view)
    assert view._excluded_cols == {2}, f"F5 뒤 이름 없는 열의 제외가 풀렸다: {view._excluded_cols}"


def test_position_memory_never_greys_a_named_column(make_view, qapp):
    """자리로 기억한 것은 '그 자리의 이름 없는 열' 이다. 그 자리에 이름이 생겼으면 다른 열이다 —
    사용자가 고른 적 없는 열을 회색칠해 변경을 숨기느니 제외를 버린다."""
    hdr = ["ID", "V", "", "W"]
    a = [hdr, ["1", "v", "m", "w"]]
    b = [hdr, ["1", "v", "M", "w"]]
    _exclude(make_view(a=a, b=b), [2], qapp)
    named = ["ID", "V", "NOTE", "W"]
    again = make_view(a=[named, a[1]], b=[named, b[1]])
    assert again._excluded_cols == set(), sorted(again._excluded_cols)


def test_the_old_name_only_format_means_the_first_of_that_name(win, dup_files):
    """옛 버전은 이름만 남겼다(같은 이름이면 마지막 열에 붙었다). 이제 그 이름의 **첫째** 로 읽는다."""
    for p in dup_files:
        _raw_put_excluded(p, ["#", "W"])
    view = _open(win, *dup_files)
    assert view._excluded_cols == {2, 3}, sorted(view._excluded_cols)


def test_new_format_round_trip(tmp_path):
    p = str(tmp_path / "a.xlsx")
    save_last_excluded(p, [("#", 1), ("", 3), ("#", 0), "W", ("W", 0)])
    assert load_last_excluded(p) == [("", 3), "#", ("#", 1), "W"]


@pytest.mark.parametrize("junk", [[("#", -1)], [("#", "1")], [("#",)], [3, "#"], [True]])
def test_bad_entries_are_not_saved(tmp_path, junk):
    p = str(tmp_path / "a.xlsx")
    save_last_excluded(p, junk)
    assert load_last_excluded(p) is None, load_last_excluded(p)


@pytest.mark.parametrize("raw", [[["#"]], [["#", -1]], [["#", "x"]], [3, "#"], [None]])
def test_bad_stored_entries_read_as_nothing(tmp_path, raw):
    p = str(tmp_path / "a.xlsx")
    _raw_put_excluded(p, raw)
    assert load_last_excluded(p) is None, load_last_excluded(p)


def test_save_reload_keeps_the_screen_not_the_memory(win, dup_files):
    """저장 뒤 다시 비교는 제외를 기억으로 다시 잡지만, 마지막엔 **저장 직전 화면**을 잇는다
    (_post_save_carry). 기억이 화면과 다를 수 있어서다 — 여기선 다른 창이 같은 파일의 기억을
    바꿔 놓은 경우. 기억 형식을 바꾸면서 이 우선순위가 깨지지 않았는지 본다."""
    from excelmerge.constants import DIR_A2B
    view = _open(win, *dup_files)
    view._on_columns_exclude_set([2], True)
    for p in dup_files:
        save_last_excluded(p, [("#", 1)])          # 다른 창이 둘째 '#' 로 바꿔 놓았다
    gen = view._compare_gen
    view._staged[(2, 4)] = DIR_A2B
    view._set_save_btn_state()
    view._save_staged("b")
    assert _wait(lambda: not view._saving_now() and view._post_save_carry is None
                 and view._compare_gen > gen and _compared(view)), view.status.currentMessage()
    ws = openpyxl.load_workbook(dup_files[1])["S"]
    assert ws["E3"].value == "x2", "전제: 저장이 실제로 B 에 썼다"
    assert view._excluded_cols == {2}, f"저장 뒤 제외가 기억 쪽으로 바뀌었다: {view._excluded_cols}"
