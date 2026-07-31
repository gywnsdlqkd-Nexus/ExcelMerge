# -*- coding: utf-8 -*-
"""대용량 비교 먹통 방지 가드 — char-diff 길이 상한 + A↔B 스크롤 미러 재진입 가드."""
from excelmerge.diff_model import _char_diff_ranges, _CHAR_DIFF_MAX, DiffTableModel


def test_char_diff_skips_when_too_long():
    """긴 문자열은 char-diff(O(n·m) SequenceMatcher)를 건너뛰고 []를 반환한다(먹통 방지)."""
    n = _CHAR_DIFF_MAX  # own+other 가 상한을 넘도록
    own = "a" * n
    other = "b" * n
    assert len(own) + len(other) > _CHAR_DIFF_MAX
    assert _char_diff_ranges(own, other) == []          # 즉시 반환(느린 diff 안 함)


def test_char_diff_works_when_short():
    """짧은 값은 정상적으로 다른 구간을 계산한다."""
    assert _char_diff_ranges("abc", "aXc") == [(1, 2)]   # 가운데 문자만 다름


def test_diff_char_ranges_long_cell_returns_empty(qapp):
    """모델 경로에서도 긴 modified 셀은 char-diff 없이 렌더된다(빈 강조)."""
    long_a = "x" * 5000
    long_b = "y" * 5000
    m = DiffTableModel("a")
    m.set_diff_data([[("modified", long_a, long_b)]],
                    [(0, 0)], {}, set(), set())
    assert m.diff_char_ranges(0, 0) == []


def test_mirror_scroll_reentrancy_guard(qapp):
    """A↔B 스크롤 미러가 재진입 가드로 무한 ping-pong 하지 않는다."""
    from excelmerge.main_window import MainWindow
    w = MainWindow()
    try:
        dv = w.tabs.currentWidget()
        ta, tb = dv.panel_a.table, dv.panel_b.table
        # dst.setValue가 되쏘는 valueChanged를 흉내 — 미러가 재진입해도 가드로 1회만 반영.
        calls = []
        orig = dv._mirror_scroll

        def spy(src, dst, orient, value):
            calls.append((orient, value))
            orig(src, dst, orient, value)
        dv._mirror_scroll = spy
        # 세로 스크롤바 값 변경 → _mirror_scroll 호출. 가드 덕에 폭주하지 않음(무한루프면 여기서 멈춤).
        ta.verticalScrollBar().setValue(0)
        ta.verticalScrollBar().valueChanged.emit(3)
        assert dv._syncing_scroll is False   # 항상 정상 복원(try/finally)
    finally:
        w.close(); w.deleteLater(); qapp.processEvents()


def test_value_caches_memoize_and_invalidate(qapp):
    """col/row_has_values 메모이즈 — 반복 호출은 재스캔하지 않고,
    저장 확정(notify_cells)으로 diff_matrix 가 제자리 변형되면 무효화돼야 한다."""
    from excelmerge.constants import STATUS_SAME, STATUS_MODIFIED
    m = DiffTableModel("a")
    # 2열: 0열엔 값 있음, 1열은 A/B 모두 빈 값
    dm = [[(STATUS_SAME, "x", "x"), (STATUS_SAME, "", "")],
          [(STATUS_SAME, "y", "y"), (STATUS_SAME, "", "")]]
    m.set_diff_data(dm, [(0, 0), (1, 1)], {}, set(), set())

    assert m.col_has_values(0) is True
    assert m.col_has_values(1) is False
    # 캐시 적재 확인 + 재호출이 같은 결과
    assert m._col_values_cache == {0: True, 1: False}
    assert m.col_has_values(1) is False

    # 저장 확정 경로 시뮬레이션: 빈 열에 값이 복사되어 들어옴(제자리 변형, 리셋 없음)
    dm[0][1] = (STATUS_MODIFIED, "new", "new")
    m.notify_cells({(0, 1)})
    assert m._col_values_cache == {}, "notify_cells 후 캐시가 무효화되지 않음"
    assert m.col_has_values(1) is True, "제자리 변형 후 스테일 결과 반환"

    # row 쪽도 동일하게 메모이즈/무효화
    assert m.row_has_values(0) is True
    assert m._row_values_cache == {0: True}
    m.notify_cells({(0, 0)})
    assert m._row_values_cache == {}
