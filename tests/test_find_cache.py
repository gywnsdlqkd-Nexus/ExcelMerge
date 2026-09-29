# -*- coding: utf-8 -*-
"""찾기(F3) 결과 캐시 — 빨라지되, 조건이 바뀌면 반드시 다시 훑어야 한다.

F3 는 누를 때마다 표 전체를 다시 훑었다(실측: 449,288셀 · '변경점만 보기' OFF 에서
한 번에 113~299 ms). 같은 조건이면 결과가 같으므로 캐시한다 — 여기서 고정할 것은
속도가 아니라 **무효화**다. 검색어·옵션·필터·제외 열·병합 준비·표 자체가 바뀌면
캐시가 남아 있으면 안 된다(없는 글자로 이동하거나, 새로 생긴 일치를 놓친다).
"""
import pytest

N_COLS = 4
HEADER = ["ID", "NAME", "NOTE", "ETC"]
# 본문 5행 — 짝수 행만 B 가 달라 '변경점만 보기'에 남는다.
A_DATA = [HEADER] + [[f"k{i}", f"apple{i}", f"메모{i}", "x"] for i in range(1, 6)]
B_DATA = [HEADER] + [[f"k{i}", (f"apple{i}" if i % 2 else f"BANANA{i}"), f"메모{i}", "x"]
                     for i in range(1, 6)]


@pytest.fixture
def dv(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    monkeypatch.setenv("APPDATA", str(tmp_path))
    win = MainWindow()
    win.show()
    try:
        view = win.tabs.currentWidget()
        view.panel_a.set_path("a.xlsx")
        view.panel_b.set_path("b.xlsx")
        view._on_loaded(A_DATA, B_DATA)
        w = getattr(view, "_diff_worker", None)
        if w is not None:
            w.wait(8000)
        for _ in range(20):
            qapp.processEvents()
        view.find_edit.setEnabled(True)
        yield view
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def _count_scans(dv, monkeypatch):
    """_iter_find_matches 호출 횟수를 세는 카운터를 달고 돌려준다."""
    calls = []
    real = dv._iter_find_matches
    monkeypatch.setattr(dv, "_iter_find_matches",
                        lambda m: (calls.append(1), real(m))[1])
    return calls


def _find(dv, term):
    dv.find_edit.setText(term)
    return list(dv._find_cells(term))


# ── 캐시가 실제로 걸린다 ─────────────────────────────────────────────────────

def test_same_query_scans_only_once(dv, monkeypatch):
    calls = _count_scans(dv, monkeypatch)
    first = _find(dv, "apple")
    again = [_find(dv, "apple") for _ in range(4)]
    assert len(calls) == 1, f"같은 조건인데 {len(calls)}번 훑었다"
    assert all(a == first for a in again)
    assert first, "전제: 일치가 있어야 의미 있는 테스트"


def test_changing_the_term_rescans(dv, monkeypatch):
    calls = _count_scans(dv, monkeypatch)
    _find(dv, "apple")
    _find(dv, "메모")
    assert len(calls) == 2


# ── 조건이 바뀌면 결과도 바뀐다(무효화) ─────────────────────────────────────

def test_case_option_changes_results(dv):
    dv.find_case_btn.setChecked(True)          # 대소문자 무시 ON
    ignore = _find(dv, "banana")
    dv.find_case_btn.setChecked(False)         # 정확히 일치만
    exact = _find(dv, "banana")
    assert ignore and not exact, f"대소문자 옵션이 무시됐다: {ignore} / {exact}"


def test_whole_word_option_changes_results(dv):
    dv.find_word_btn.setChecked(False)
    part = _find(dv, "appl")
    dv.find_word_btn.setChecked(True)
    whole = _find(dv, "appl")
    assert part and not whole, f"전체 단어 옵션이 무시됐다: {part} / {whole}"


def test_filter_toggle_changes_results(dv, qapp):
    """'변경점만 보기'로 숨은 행은 찾기에서도 빠진다 — 토글하면 목록이 달라져야 한다."""
    dv.diff_only_btn.setChecked(False)
    for _ in range(10):
        qapp.processEvents()
    all_rows = _find(dv, "메모")
    dv.diff_only_btn.setChecked(True)
    for _ in range(10):
        qapp.processEvents()
    visible_only = _find(dv, "메모")
    assert len(visible_only) < len(all_rows), \
        f"필터를 켰는데 결과가 그대로다: {len(all_rows)} → {len(visible_only)}"
    tbl = dv.panel_a.table
    assert all(not tbl.isRowHidden(r) for r, _c in visible_only)


def test_excluding_a_column_drops_its_matches(dv, qapp):
    before = _find(dv, "메모")
    assert before and all(c == 2 for _r, c in before)
    dv._on_columns_exclude_set([2], True)
    for _ in range(10):
        qapp.processEvents()
    assert _find(dv, "메모") == [], "제외한 열의 일치가 남아 있다"


def test_staging_changes_what_is_searchable(dv, qapp):
    """병합 준비하면 화면 글자가 한쪽 값으로 바뀐다 — 찾기도 그 글자를 따라가야 한다."""
    from excelmerge.constants import DIR_B2A
    hits = _find(dv, "BANANA")
    assert hits, "전제: B 쪽에만 있는 글자"
    r, c = hits[0]
    dv._staged[(r, c)] = DIR_B2A          # B→A 준비 — 양쪽 모두 B 값을 보여준다
    dv._notify_cells([(r, c)])
    for _ in range(10):
        qapp.processEvents()
    assert (r, c) in _find(dv, "BANANA")
    # 반대로 A 쪽 값으로 준비하면 B 쪽 글자는 화면에서 사라진다
    from excelmerge.constants import DIR_A2B
    dv._staged[(r, c)] = DIR_A2B
    dv._notify_cells([(r, c)])
    for _ in range(10):
        qapp.processEvents()
    assert (r, c) not in _find(dv, "BANANA"), "화면에 없는 글자가 검색됐다"


def test_new_comparison_invalidates_the_cache(dv, qapp):
    _find(dv, "apple")
    # 새 비교에도 '변경'이 있어야 기본 필터(변경점만 보기)에서 행이 보인다.
    other_a = [HEADER] + [["z1", "cherry1", "노트", "x"]]
    other_b = [HEADER] + [["z1", "cherry1", "노트", "y"]]
    dv._on_loaded(other_a, other_b)
    w = getattr(dv, "_diff_worker", None)
    if w is not None:
        w.wait(8000)
    for _ in range(20):
        qapp.processEvents()
    assert _find(dv, "apple") == [], "이전 비교의 결과가 남아 있다"
    assert _find(dv, "cherry"), "새 비교의 내용을 못 찾는다"


def test_cached_list_is_not_corrupted_by_navigation(dv):
    """F3 를 눌러 이동해도 캐시된 목록이 줄거나 뒤섞이면 안 된다."""
    first = _find(dv, "apple")
    for _ in range(3):
        dv._goto_find(+1)
    assert _find(dv, "apple") == first


def test_preview_mode_is_not_cached(dv, qapp):
    """파일 1개(미리보기)는 캐시하지 않는다 — 조건을 열쇠로 표현할 수 없어서."""
    data = [HEADER] + [["p1", "apple9", "메모", "x"]]
    dv._diff_matrix = []
    dv._preview_data = {"a": data, "b": []}
    dv.panel_a.table.model().set_preview_data(data)
    dv._find_cache = None
    for _ in range(10):
        qapp.processEvents()
    dv.find_edit.setText("apple9")
    assert dv._find_cells("apple9"), "미리보기에서 찾지 못했다"
    assert dv._find_cache is None, "미리보기인데 캐시를 남겼다"


# ── 안전망: 열쇠로 표현되지 않는 변화 대비 ───────────────────────────────────
# 캐시 열쇠(검색어·옵션·필터·제외 열·병합 준비·표 identity)가 대부분을 덮지만,
# 행 가시성을 다시 계산하거나 표를 다시 그리는 경로에서는 명시적으로도 버린다.
# 이 안전망이 사라지면(누군가 호출을 지우면) 아래 두 테스트가 알려 준다.

def test_refiltering_drops_the_cache(dv):
    _find(dv, "apple")
    assert dv._find_cache is not None, "전제: 캐시가 채워져 있어야 한다"
    dv._apply_diff_filter()
    assert dv._find_cache is None, "행 가시성을 다시 계산했는데 캐시가 남았다"


def test_repopulating_drops_the_cache(dv):
    _find(dv, "apple")
    assert dv._find_cache is not None
    dv._refresh_tables()
    assert dv._find_cache is None, "표를 다시 그렸는데 캐시가 남았다"
