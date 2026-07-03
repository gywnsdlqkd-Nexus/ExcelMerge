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

    # 3. 모델 리셋 중 cell_edit 자동 적용 오발동 없음
    fired = []
    win.panel_a.cell_value_edited.connect(lambda r, c, v: fired.append((r, c, v)))
    tbl._set_current_cell(1, 1)
    app.processEvents()
    win.panel_a.cell_edit.setText("draft-value")   # Enter 없이 입력만
    win._refresh_tables()                          # 모델 리셋 → 선택 해제
    app.processEvents()
    assert not fired, f"리셋 중 자동 적용 발화: {fired}"

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


if __name__ == "__main__":
    test_golden_roles()
    test_headers()
    test_notify_equals_fresh_populate()
    test_cell_kind()
    test_view_behaviors()
    print("ALL MODEL/VIEW TESTS PASS")
