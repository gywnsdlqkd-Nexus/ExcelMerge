# -*- coding: utf-8 -*-
"""모델/뷰 전환 골든 패리티 테스트.

구 QTableWidget populate() 루프의 텍스트/배경색 로직을 오라클로 이식해,
DiffTableModel의 DisplayRole/BackgroundRole이 모든 상태 조합에서 동일한지
검증한다. 추가로 필터 델타, 선택 클램프, 자동 편집 오발동 방지를 확인한다.

실행: QT_QPA_PLATFORM=offscreen python tests/test_model_view.py
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication

from excelmerge.diff_model import DiffTableModel, EXTRA_ROWS, EXTRA_COLS
from excelmerge.theme import DIFF_COLORS


def oracle_cell(diff_matrix, which, merged_set, staged, excluded_cols, r, c):
    """구 populate() 핫루프(텍스트/색 결정)를 그대로 이식한 오라클."""
    status, a_val, b_val = diff_matrix[r][c]
    direction = staged.get((r, c))
    if direction == "a_to_b":
        text = a_val
    elif direction == "b_to_a":
        text = b_val
    else:
        text = a_val if which == "a" else b_val
    if c in excluded_cols:
        color = DIFF_COLORS["same"]
    elif (r, c) in merged_set:
        color = DIFF_COLORS["merged"]
    elif direction is not None:
        color = DIFF_COLORS["staged"]
    else:
        color = DIFF_COLORS[status]
    return text, color


def dump(model):
    out = []
    for r in range(model.rowCount()):
        row = []
        for c in range(model.columnCount()):
            idx = model.index(r, c)
            text = model.data(idx, Qt.DisplayRole) or ""
            bg = model.data(idx, Qt.BackgroundRole)
            row.append((text, bg.getRgb()[:3] if bg is not None else None))
        out.append(row)
    return out


def make_state():
    # 상태 조합 총망라: same/added/modified + staged 양방향 + merged + excluded
    dm = [
        [("same", "H1", "H1"), ("same", "H2", "H2"), ("modified", "H3", "H3x")],
        [("same", "a", "a"), ("modified", "b1", "b2"), ("added", "", "new")],
        [("modified", "c1", "c2"), ("modified", "d1", "d2"), ("same", "e", "e")],
        [("added", "", "f"), ("modified", "g1", ""), ("modified", "h1", "h2")],
    ]
    meta = [(0, 0), (1, 1), (2, None), (None, 3)]
    staged = {(1, 1): "a_to_b", (2, 0): "b_to_a"}
    merged = {(2, 1)}
    excluded = {2}
    return dm, meta, staged, merged, excluded


def test_golden_roles():
    dm, meta, staged, merged, excluded = make_state()
    for side in ("a", "b"):
        m = DiffTableModel(side)
        m.set_diff_data(dm, meta, staged, merged, excluded)
        assert m.rowCount() == len(dm) + EXTRA_ROWS
        assert m.columnCount() == len(dm[0]) + EXTRA_COLS
        for r in range(len(dm)):
            for c in range(len(dm[0])):
                want_text, want_color = oracle_cell(dm, side, merged, staged, excluded, r, c)
                idx = m.index(r, c)
                got_text = m.data(idx, Qt.DisplayRole) or ""
                got_bg = m.data(idx, Qt.BackgroundRole)
                assert got_text == want_text, f"{side} ({r},{c}) text {got_text!r} != {want_text!r}"
                assert got_bg is not None and got_bg.getRgb() == want_color.getRgb(), \
                    f"{side} ({r},{c}) bg mismatch"
                assert m.data(idx, Qt.TextAlignmentRole) == int(Qt.AlignVCenter | Qt.AlignLeft)
        # 여분 행/열: 텍스트 없음 + 기본 배경(None)
        assert m.data(m.index(len(dm), 0), Qt.DisplayRole) is None
        assert m.data(m.index(len(dm), 0), Qt.BackgroundRole) is None
        assert m.data(m.index(0, len(dm[0])), Qt.BackgroundRole) is None
    print("PASS test_golden_roles")


def test_headers():
    dm, meta, staged, merged, excluded = make_state()
    m = DiffTableModel("a")
    m.set_diff_data(dm, meta, staged, merged, excluded)
    m.set_key_col(0)
    assert m.headerData(0, Qt.Horizontal, Qt.DisplayRole) == "🔑 A"
    assert m.headerData(0, Qt.Horizontal, Qt.BackgroundRole).getRgb()[:3] == (255, 213, 0)
    assert m.headerData(2, Qt.Horizontal, Qt.DisplayRole) == "⊘ C"
    assert m.headerData(2, Qt.Horizontal, Qt.ForegroundRole).getRgb()[:3] == (140, 140, 140)
    assert m.headerData(1, Qt.Horizontal, Qt.DisplayRole) == "B"
    assert m.headerData(1, Qt.Horizontal, Qt.BackgroundRole).getRgb()[:3] == (232, 234, 240)
    # 세로: meta 원본 행번호, None이면 "-" (A측: meta[3][0] is None)
    assert m.headerData(0, Qt.Vertical, Qt.DisplayRole) == "1"
    assert m.headerData(3, Qt.Vertical, Qt.DisplayRole) == "-"
    b = DiffTableModel("b")
    b.set_diff_data(dm, meta, staged, merged, excluded)
    assert b.headerData(2, Qt.Vertical, Qt.DisplayRole) == "-"   # meta[2][1] is None
    assert b.headerData(3, Qt.Vertical, Qt.DisplayRole) == "4"
    # 여분 행 라벨은 연속 번호
    assert m.headerData(len(dm), Qt.Vertical, Qt.DisplayRole) == str(len(dm) + 1)
    # 미리보기 모드: 색·아이콘 없는 순수 라벨
    m.set_preview_data([["x", "y"], ["1", "2"]])
    assert m.headerData(0, Qt.Horizontal, Qt.DisplayRole) == "A"
    assert m.headerData(0, Qt.Horizontal, Qt.BackgroundRole) is None
    print("PASS test_headers")


def test_notify_equals_fresh_populate():
    """상태 변형 + notify 후의 모델 스냅샷 == 같은 상태로 새로 populate한 스냅샷."""
    dm, meta, staged, merged, excluded = make_state()
    live = DiffTableModel("b")
    live.set_diff_data(dm, meta, staged, merged, excluded)
    # 변형: 셀 스테이징 추가, 기존 staged 하나 merged로 승격, 편집으로 status 변경
    staged[(3, 2)] = "a_to_b"
    merged.add((1, 1)); del staged[(1, 1)]
    dm[2][0] = ("same", "c2", "c2")
    live.notify_cells({(3, 2), (1, 1), (2, 0)})
    fresh = DiffTableModel("b")
    fresh.set_diff_data(dm, meta, staged, merged, excluded)
    assert dump(live) == dump(fresh), "notify 후 상태가 fresh populate와 다름"
    print("PASS test_notify_equals_fresh_populate")


def test_cell_kind():
    dm, meta, staged, merged, excluded = make_state()
    m = DiffTableModel("a")
    m.set_diff_data(dm, meta, staged, merged, excluded)
    assert m.cell_kind(1, 1) == "staged"
    assert m.cell_kind(2, 1) == "merged"
    assert m.cell_kind(1, 2) == "same"        # excluded 열 → same 취급
    assert m.cell_kind(2, 0) == "staged"
    assert m.cell_kind(1, 0) == "same"
    assert m.cell_kind(3, 0) == "changed"     # added
    assert m.cell_kind(len(dm) + 1, 0) == "same"   # 여분 행
    print("PASS test_cell_kind")


def test_view_behaviors():
    from excelmerge.main_window import MainWindow
    app = QApplication.instance() or QApplication([])
    win = MainWindow()
    dm, meta, staged, merged, excluded = make_state()
    win._diff_matrix = dm
    win._diff_row_meta = meta
    win._staged = staged
    win._merged_cells = merged
    win._excluded_cols = set(excluded)
    win._refresh_tables()
    tbl = win.panel_a.table

    # 1. 필터 델타: 변경행만 보기 토글 시 isRowHidden 벡터가 오라클과 일치
    win._diff_only = True
    win._apply_diff_filter()
    excl = win._excluded_cols
    for r, row in enumerate(dm):
        want_hidden = (r != 0) and not any(
            st != "same" for c, (st, *_) in enumerate(row) if c not in excl)
        assert tbl.isRowHidden(r) == want_hidden, f"row {r} hidden mismatch"
    assert not tbl._user_row_heights, "_user_row_heights 오염됨"
    win._diff_only = False
    win._apply_diff_filter()
    assert not any(tbl.isRowHidden(r) for r in range(len(dm))), "필터 해제 실패"

    # 2. Ctrl+Shift+End 클램프: 데이터 영역 밖 선택 없음
    tbl._select_range(0, 0, tbl.rowCount() - 1, tbl.columnCount() - 1)
    sel = tbl.get_selected_cells()
    assert sel, "선택 없음"
    assert max(r for r, _ in sel) == len(dm) - 1
    assert max(c for _, c in sel) == len(dm[0]) - 1

    # 3. 셀 값 표시란은 읽기전용 (직접 편집 기능 제거됨)
    assert win.panel_a.cell_edit.isReadOnly()
    assert not hasattr(win.panel_a, "cell_value_edited")

    # 4. staged 오버라이드 텍스트가 표시/TSV에 반영
    assert tbl.model().display_text(1, 1) == "b1"   # a_to_b → a_val
    assert win.panel_b.table.model().display_text(1, 1) == "b1"

    # 5. 부분 갱신이 선택/스크롤을 보존 (기본 특성 확인)
    tbl._set_current_cell(2, 2)
    app.processEvents()
    before = tbl.get_selected_cells()
    win.panel_a.table.model().notify_cells({(2, 2)})
    app.processEvents()
    assert tbl.get_selected_cells() == before, "notify가 선택을 파괴함"

    win.close()
    print("PASS test_view_behaviors")


def test_header_multiselect_extension():
    """헤더 선택 후 Shift+방향키(한 칸)·Ctrl+Shift+방향키(끝까지) 확장 회귀 테스트.
    수정 전에는 setCurrentIndex의 SelectCurrent가 범위 선택을 1셀로 붕괴시켰다."""
    from PyQt5.QtTest import QTest
    from excelmerge.main_window import MainWindow
    app = QApplication.instance() or QApplication([])
    win = MainWindow()

    # 값은 0~4열/0~7행에만 존재. 5~6열과 8행은 엑셀 유령 셀처럼 빈 값으로 채워
    # '값 기준 마지막' 경계 검증에 쓴다.
    def row_vals(vals):
        return [("same", v, v) for v in vals]
    dm = [row_vals([f"h{c}" for c in range(5)] + ["", ""])]
    for r in range(1, 8):
        dm.append(row_vals([f"a{r}{c}" for c in range(5)] + ["", ""]))
    dm.append(row_vals([""] * 7))
    meta = [(r, r) for r in range(len(dm))]
    win._diff_matrix, win._diff_row_meta = dm, meta
    win._refresh_tables()
    tbl = win.panel_a.table
    win.show()
    app.processEvents()
    row_n = tbl.rowCount()
    col_n = tbl.columnCount()

    # 열 헤더: B열 선택 → Shift+→ 두 번 → [1,2,3]
    tbl._select_col(1)
    tbl._on_h_section_pressed(1)
    QTest.keyClick(tbl, Qt.Key_Right, Qt.ShiftModifier)
    assert tbl._full_columns_selected() == [1, 2], tbl._full_columns_selected()
    QTest.keyClick(tbl, Qt.Key_Right, Qt.ShiftModifier)
    assert tbl._full_columns_selected() == [1, 2, 3], tbl._full_columns_selected()
    assert len(tbl.selectionModel().selectedIndexes()) == 3 * row_n
    # Shift+← 로 앵커 방향 축소
    QTest.keyClick(tbl, Qt.Key_Left, Qt.ShiftModifier)
    assert tbl._full_columns_selected() == [1, 2], tbl._full_columns_selected()
    # Ctrl+Shift+→ : '값이 있는' 마지막 열(4)까지 — 매트릭스 폭(7)이나
    # 그리드 폭이 아닌 실제 값 기준 (유령 셀 열 5~6 제외)
    assert tbl.model().data_cols == 7
    assert tbl.model().col_has_values(4) and not tbl.model().col_has_values(5)
    QTest.keyClick(tbl, Qt.Key_Right, Qt.ControlModifier | Qt.ShiftModifier)
    assert tbl._full_columns_selected() == [1, 2, 3, 4], tbl._full_columns_selected()
    # 경계에서 한 번 더 → 엑셀처럼 현재 위치 기준 재판정해 그리드 끝까지
    QTest.keyClick(tbl, Qt.Key_Right, Qt.ControlModifier | Qt.ShiftModifier)
    assert tbl._full_columns_selected() == list(range(1, col_n)), tbl._full_columns_selected()
    # 그리드 끝에서 한 번 더 → 변화 없음
    QTest.keyClick(tbl, Qt.Key_Right, Qt.ControlModifier | Qt.ShiftModifier)
    assert tbl._full_columns_selected() == list(range(1, col_n)), tbl._full_columns_selected()

    # 행 헤더: 3행 선택 → Shift+↓ → [2,3], Ctrl+Shift+↓ 값 있는 마지막 행(7)까지
    assert tbl.model().data_rows == 9
    assert tbl.model().row_has_values(7) and not tbl.model().row_has_values(8)
    tbl._select_row(2)
    tbl._on_v_section_pressed(2)
    QTest.keyClick(tbl, Qt.Key_Down, Qt.ShiftModifier)
    assert tbl._full_rows_selected() == [2, 3], tbl._full_rows_selected()
    QTest.keyClick(tbl, Qt.Key_Down, Qt.ControlModifier | Qt.ShiftModifier)
    assert tbl._full_rows_selected() == list(range(2, 8)), tbl._full_rows_selected()
    # 경계에서 한 번 더 → 그리드 끝까지
    QTest.keyClick(tbl, Qt.Key_Down, Qt.ControlModifier | Qt.ShiftModifier)
    assert tbl._full_rows_selected() == list(range(2, row_n)), tbl._full_rows_selected()

    win.close()
    print("PASS test_header_multiselect_extension")


def test_ctrl_jump_single_selection():
    """Ctrl+방향키 점프 회귀 테스트 — 엑셀처럼 단일 선택으로 이동해야 한다.
    수정 전에는 Ctrl 수정자 때문에 setCurrentIndex가 Toggle로 동작해
    원래 셀과 대상 셀이 함께 선택됐다."""
    from PyQt5.QtTest import QTest
    from excelmerge.main_window import MainWindow
    app = QApplication.instance() or QApplication([])
    win = MainWindow()
    # 0~2열 값 있음, 3열 빈 값, 4열 값 있음
    dm = [[("same", f"h{c}", f"h{c}") for c in range(5)]]
    for r in range(1, 6):
        dm.append([("same", "" if c == 3 else f"v{r}{c}",
                    "" if c == 3 else f"v{r}{c}") for c in range(5)])
    meta = [(r, r) for r in range(len(dm))]
    win._diff_matrix, win._diff_row_meta = dm, meta
    win._refresh_tables()
    tbl = win.panel_a.table
    win.show()
    app.processEvents()

    # 값 경계 점프: (2,0) → 연속 데이터 끝 (2,2) — 단일 선택
    tbl._move_current_cell(2, 0)
    QTest.keyClick(tbl, Qt.Key_Right, Qt.ControlModifier)
    assert tbl.get_selected_cells() == {(2, 2)}, tbl.get_selected_cells()
    # 빈 열 건너 다음 값 셀 (2,4) — 단일 선택
    QTest.keyClick(tbl, Qt.Key_Right, Qt.ControlModifier)
    assert tbl.get_selected_cells() == {(2, 4)}, tbl.get_selected_cells()
    # 값이 더 없으면 마지막 셀로 — 단일 선택
    QTest.keyClick(tbl, Qt.Key_Right, Qt.ControlModifier)
    assert tbl.get_selected_cells() == {(2, tbl.columnCount() - 1)}, tbl.get_selected_cells()

    # Ctrl+Shift+방향키: 범위 선택 유지 (붕괴 없음)
    tbl._move_current_cell(2, 0)
    QTest.keyClick(tbl, Qt.Key_Right, Qt.ControlModifier | Qt.ShiftModifier)
    assert tbl.get_selected_cells() == {(2, 0), (2, 1), (2, 2)}, tbl.get_selected_cells()

    # Ctrl+Shift+End: 데이터 영역으로 클램프된 범위, 붕괴 없음
    tbl._move_current_cell(1, 1)
    QTest.keyClick(tbl, Qt.Key_End, Qt.ControlModifier | Qt.ShiftModifier)
    sel = tbl.get_selected_cells()
    assert max(r for r, _ in sel) == len(dm) - 1, sel
    assert max(c for _, c in sel) == 4, sel
    assert len(sel) == (len(dm) - 1) * 4, len(sel)

    # Ctrl+Home: 단일 이동
    QTest.keyClick(tbl, Qt.Key_Home, Qt.ControlModifier)
    assert tbl.get_selected_cells() == {(0, 0)}, tbl.get_selected_cells()

    win.close()
    print("PASS test_ctrl_jump_single_selection")


def test_selection_mirror_compact_ranges():
    """선택 미러가 range를 직사각형으로 압축하는지 회귀 테스트.
    수정 전에는 열 전체 미러가 행마다 range를 만들어(수천 개) 이후 모든
    선택 질의·페인팅이 느려졌다 (헤더 클릭 딜레이의 주범)."""
    from excelmerge.main_window import MainWindow
    app = QApplication.instance() or QApplication([])
    win = MainWindow()
    dm = [[("same", f"v{r}{c}", f"v{r}{c}") for c in range(6)] for r in range(120)]
    meta = [(r, r) for r in range(len(dm))]
    win._diff_matrix, win._diff_row_meta = dm, meta
    win._refresh_tables()
    src, dst = win.panel_a.table, win.panel_b.table
    win.show()
    app.processEvents()

    # 열 전체 선택 → 동기화된 반대 패널도 range 1개 + 셀 집합 동일
    src._select_col(2)
    app.processEvents()
    assert len(dst.selectionModel().selection()) == 1, \
        len(dst.selectionModel().selection())
    assert dst.get_selected_cells() == src.get_selected_cells()

    # mirror_selection(셀 집합) 경로도 직사각형 병합: 2×3 블록 → range 1개
    block = {(r, c) for r in range(10, 12) for c in range(1, 4)}
    dst.mirror_selection(block)
    assert dst.get_selected_cells() == block
    assert len(dst.selectionModel().selection()) == 1

    # 비직사각형(L자)도 셀 집합은 정확히 보존
    lshape = {(0, 0), (1, 0), (1, 1)}
    dst.mirror_selection(lshape)
    assert dst.get_selected_cells() == lshape

    # 단일 셀 판정 헬퍼
    src._move_current_cell(5, 5)
    assert src._single_selected_cell() == (5, 5)
    src._select_col(1)
    assert src._single_selected_cell() is None

    win.close()
    print("PASS test_selection_mirror_compact_ranges")


def test_staging_color_with_empty_initial_state():
    """비교 직후(스테이징 0개) 상태에서 병합 준비 시 색이 바뀌는지 회귀 테스트.
    populate의 `staged or {}` falsy 체크가 빈 dict일 때 새 객체를 만들어
    모델이 MainWindow 상태와 분리되던 버그."""
    from excelmerge.main_window import MainWindow
    from excelmerge.theme import DIFF_COLORS
    app = QApplication.instance() or QApplication([])
    win = MainWindow()
    dm = [[("same", "h", "h")], [("modified", "x", "y")]]
    win._diff_matrix = dm
    win._diff_row_meta = [(0, 0), (1, 1)]
    win._staged = {}            # 실제 흐름: 비교 직후엔 빈 dict
    win._merged_cells = set()
    win._refresh_tables()
    model = win.panel_a.table.model()
    assert model._staged is win._staged, "모델이 MainWindow staged와 분리됨"
    assert model._merged is win._merged_cells, "모델이 merged_cells와 분리됨"
    win._staged[(1, 0)] = "a_to_b"
    win._notify_cells({(1, 0)})
    bg = model.data(model.index(1, 0), Qt.BackgroundRole)
    assert bg.getRgb() == DIFF_COLORS["staged"].getRgb(), bg.getRgb()
    win.close()
    print("PASS test_staging_color_with_empty_initial_state")


def test_ctrl_jump_skips_hidden_rows():
    """'변경 행만 보기' 상태에서 Ctrl+↓ 점프가 숨겨진 행에 착지하지 않는지."""
    from PyQt5.QtTest import QTest
    from excelmerge.main_window import MainWindow
    app = QApplication.instance() or QApplication([])
    win = MainWindow()
    # 행 1, 4만 변경 — 필터 켜면 2, 3, 5는 숨겨짐 (행 0은 헤더로 항상 표시)
    dm = [[("same", "h", "h")]]
    for r in range(1, 6):
        changed = r in (1, 4)
        dm.append([("modified" if changed else "same", f"a{r}", f"b{r}" if changed else f"a{r}")])
    win._diff_matrix = dm
    win._diff_row_meta = [(r, r) for r in range(len(dm))]
    win._refresh_tables()
    win._diff_only = True
    win._apply_diff_filter()
    tbl = win.panel_a.table
    win.show()
    app.processEvents()
    assert tbl.isRowHidden(2) and tbl.isRowHidden(3), "필터 전제 불성립"

    tbl._move_current_cell(1, 0)
    QTest.keyClick(tbl, Qt.Key_Down, Qt.ControlModifier)
    r, c = tbl._current_cell()
    assert not tbl.isRowHidden(r), f"숨겨진 행 {r}에 착지"
    assert r == 4, f"기대 4행, 실제 {r}행"   # 1 다음 보이는 값 행
    win.close()
    print("PASS test_ctrl_jump_skips_hidden_rows")


def test_find_in_preview_mode():
    """파일 하나만 로드된 미리보기 상태에서도 Ctrl+F 검색이 동작하는지."""
    from excelmerge.main_window import MainWindow
    app = QApplication.instance() or QApplication([])
    win = MainWindow()
    win.panel_a.preview([["id", "name"], ["1", "apple"], ["2", "banana"]])
    matches = list(win._iter_find_matches(win._make_find_matcher("banana")))
    assert matches == [(2, 1)], matches
    # diff 없이도 _goto_find가 이동시키는지
    win.find_edit.setText("apple")
    win._set_find_enabled(True)
    win._goto_find(+1)
    assert win.panel_a.table._current_cell() == (1, 1)
    win.close()
    print("PASS test_find_in_preview_mode")


def test_cell_status_one_sided_is_added():
    """한쪽 파일에만 값이 있으면 방향과 무관하게 'added'로 분류."""
    from excelmerge.diff_engine import _cell_status
    assert _cell_status("", "new") == "added"      # B 전용
    assert _cell_status("old", "") == "added"      # A 전용 (기존엔 modified)
    assert _cell_status("a", "b") == "modified"
    assert _cell_status("x", "x") == "same"
    assert _cell_status("", "") == "same"
    print("PASS test_cell_status_one_sided_is_added")


def test_filter_keeps_merged_rows_visible():
    """저장(병합 확정) 후에도 '변경 행만 보기'가 병합됨 행을 숨기지 않는지.
    저장 시 상태가 same이 되면서 행이 즉시 사라져 병합 결과를 볼 수 없던 문제."""
    from excelmerge.main_window import MainWindow
    app = QApplication.instance() or QApplication([])
    win = MainWindow()
    dm = [
        [("same", "h", "h")],
        [("same", "m", "m")],    # 병합 확정된 행 (merged 셀 보유)
        [("same", "u", "u")],    # 그냥 동일한 행
        [("modified", "x", "y")],
    ]
    win._diff_matrix = dm
    win._diff_row_meta = [(r, r) for r in range(len(dm))]
    win._merged_cells = {(1, 0)}
    win._refresh_tables()
    win._diff_only = True
    win._apply_diff_filter()
    tbl = win.panel_a.table
    assert not tbl.isRowHidden(0), "헤더 행"
    assert not tbl.isRowHidden(1), "병합됨 행이 숨겨짐"
    assert tbl.isRowHidden(2), "동일 행은 숨겨져야 함"
    assert not tbl.isRowHidden(3), "변경 행"
    win.close()
    print("PASS test_filter_keeps_merged_rows_visible")


if __name__ == "__main__":
    test_golden_roles()
    test_headers()
    test_notify_equals_fresh_populate()
    test_cell_kind()
    test_view_behaviors()
    test_header_multiselect_extension()
    test_ctrl_jump_single_selection()
    test_selection_mirror_compact_ranges()
    test_staging_color_with_empty_initial_state()
    test_ctrl_jump_skips_hidden_rows()
    test_find_in_preview_mode()
    test_cell_status_one_sided_is_added()
    test_filter_keeps_merged_rows_visible()
    print("ALL MODEL/VIEW TESTS PASS")
