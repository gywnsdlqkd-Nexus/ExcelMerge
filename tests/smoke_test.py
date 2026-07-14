"""리팩토링 단계별 스모크 테스트.

QT_QPA_PLATFORM=offscreen 환경에서 로더 → diff → 테이블 표시 파이프라인이
동작하는지 검증한다. 모놀리스/패키지 어느 단계에서도 실행 가능하도록
excelmerge 패키지를 우선 시도하고 실패 시 excel_diff_merge에서 가져온다.
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

HERE = os.path.dirname(os.path.abspath(__file__))          # tests/
ROOT = os.path.dirname(HERE)                               # 프로젝트 루트
sys.path.insert(0, ROOT)                                   # excelmerge/·진입점 import용
FIXTURE_DIR = os.path.join(HERE, "fixtures")

try:
    from excelmerge.loaders import load_values_any
    from excelmerge.diff_engine import compute_diff
except ImportError:
    from excel_diff_merge import load_values_any, compute_diff

try:
    from excelmerge.widgets import EXTRA_ROWS
except ImportError:
    from excel_diff_merge import EXTRA_ROWS

try:
    from excelmerge.theme import DIFF_COLORS
except ImportError:
    from excel_diff_merge import DIFF_COLORS

import excel_diff_merge  # 진입점이 항상 import 가능해야 한다

try:
    from excelmerge.main_window import MainWindow
except ImportError:
    from excel_diff_merge import MainWindow


def make_fixtures():
    import openpyxl
    os.makedirs(FIXTURE_DIR, exist_ok=True)
    path_a = os.path.join(FIXTURE_DIR, "smoke_a.xlsx")
    path_b = os.path.join(FIXTURE_DIR, "smoke_b.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name", "Score"])
    ws.append([1, "Alice", 10])
    ws.append([2, "Bob", 20])
    ws.append([3, "Carol", 30])
    wb.save(path_a)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name", "Score"])
    ws.append([1, "Alice", 10])
    ws.append([2, "Bobby", 20])
    ws.append([4, "Dave", 40])
    wb.save(path_b)
    return path_a, path_b


def display_text(table, r, c):
    """QTableWidget(item) / QTableView(model) 양쪽에서 표시 텍스트를 읽는다."""
    model = table.model()
    if hasattr(model, "display_text"):
        return model.display_text(r, c)
    item = table.item(r, c)
    return item.text() if item is not None else ""


def main():
    path_a, path_b = make_fixtures()

    # 1. 로더 — 값(계산값)만 로드
    a_vals = load_values_any(path_a)
    b_vals = load_values_any(path_b)
    assert len(a_vals) == 4 and len(a_vals[0]) == 3, f"A shape {len(a_vals)}x{len(a_vals[0])}"
    assert all(isinstance(v, str) for row in a_vals for v in row), "값이 str이 아님"
    assert a_vals[1] == ["1", "Alice", "10"], f"A row1: {a_vals[1]}"

    # 2. diff
    dm, meta = compute_diff(a_vals, b_vals, key_col=0)
    assert len(dm) == 5, f"diff 행수 {len(dm)} != 5"          # 헤더 + 키 1,2,3,4
    assert dm[2][1] == ("modified", "Bob", "Bobby"), f"dm[2][1]: {dm[2][1]}"
    assert dm[4][1][0] == "added", f"B 전용 행 status: {dm[4][1][0]}"
    assert meta[0] == (0, 0) and meta[4][0] is None, f"row_meta: {meta}"

    # 3. UI 파이프라인
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])
    win = MainWindow()
    # MainWindow는 탭 셸 — 실제 비교 UI는 활성 탭의 DiffView에 있다.
    view = win.tabs.currentWidget()
    view._diff_matrix = dm
    view._diff_row_meta = meta
    view.panel_a.table.populate(dm, "a", set(), {}, meta, set())
    view.panel_b.table.populate(dm, "b", set(), {}, meta, set())
    assert view.panel_a.table.rowCount() == len(dm) + EXTRA_ROWS, \
        f"rowCount {view.panel_a.table.rowCount()}"
    assert display_text(view.panel_a.table, 2, 1) == "Bob"
    assert display_text(view.panel_b.table, 2, 1) == "Bobby"
    assert display_text(view.panel_a.table, 4, 1) == ""       # A에 없는 행
    assert display_text(view.panel_b.table, 4, 1) == "Dave"

    # staged 오버라이드 표시
    view.panel_a.table.populate(dm, "a", set(), {(2, 1): "b_to_a"}, meta, set())
    assert display_text(view.panel_a.table, 2, 1) == "Bobby", "staged 텍스트 오버라이드 실패"

    # 4. 색상 불변
    assert DIFF_COLORS["staged"].getRgb()[:3] == (255, 185, 80)
    assert DIFF_COLORS["merged"].getRgb()[:3] == (173, 216, 230)
    assert DIFF_COLORS["added"].getRgb()[:3] == (198, 239, 206)
    assert DIFF_COLORS["modified"].getRgb()[:3] == (255, 235, 156)

    win.close()
    print("SMOKE TEST PASS")


if __name__ == "__main__":
    main()
