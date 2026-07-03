"""리팩토링 단계별 스모크 테스트.

QT_QPA_PLATFORM=offscreen 환경에서 로더 → diff → 테이블 표시 파이프라인이
동작하는지 검증한다. 모놀리스/패키지 어느 단계에서도 실행 가능하도록
excelmerge 패키지를 우선 시도하고 실패 시 excel_diff_merge에서 가져온다.
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
FIXTURE_DIR = os.path.join(HERE, "tests", "fixtures")

try:
    from excelmerge.loaders import load_sheet_with_formulas_any
    from excelmerge.diff_engine import compute_diff
except ImportError:
    from excel_diff_merge import load_sheet_with_formulas_any, compute_diff

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
    ws.append([3, "Carol", "=C2"])  # 캐시값 없는 수식 → 지연 평가 경로
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

    # 1. 로더
    a_vals, a_fmls = load_sheet_with_formulas_any(path_a)
    b_vals, b_fmls = load_sheet_with_formulas_any(path_b)
    assert len(a_vals) == 4 and len(a_vals[0]) == 3, f"A shape {len(a_vals)}x{len(a_vals[0])}"
    assert len(a_fmls) == len(a_vals), "values/formulas 행수 불일치"
    assert all(isinstance(v, str) for row in a_vals for v in row), "값이 str이 아님"
    assert a_vals[1] == ["1", "Alice", "10"], f"A row1: {a_vals[1]}"
    assert a_fmls[3][2] == "=C2", f"수식 텍스트 보존 실패: {a_fmls[3][2]}"

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
    win._diff_matrix = dm
    win._diff_row_meta = meta
    win.panel_a.table.populate(dm, "a", set(), {}, meta, set())
    win.panel_b.table.populate(dm, "b", set(), {}, meta, set())
    assert win.panel_a.table.rowCount() == len(dm) + EXTRA_ROWS, \
        f"rowCount {win.panel_a.table.rowCount()}"
    assert display_text(win.panel_a.table, 2, 1) == "Bob"
    assert display_text(win.panel_b.table, 2, 1) == "Bobby"
    assert display_text(win.panel_a.table, 4, 1) == ""       # A에 없는 행
    assert display_text(win.panel_b.table, 4, 1) == "Dave"

    # staged 오버라이드 표시
    win.panel_a.table.populate(dm, "a", set(), {(2, 1): "b_to_a"}, meta, set())
    assert display_text(win.panel_a.table, 2, 1) == "Bobby", "staged 텍스트 오버라이드 실패"

    # 4. 색상 불변
    assert DIFF_COLORS["staged"].getRgb()[:3] == (255, 185, 80)
    assert DIFF_COLORS["merged"].getRgb()[:3] == (173, 216, 230)
    assert DIFF_COLORS["added"].getRgb()[:3] == (198, 239, 206)
    assert DIFF_COLORS["modified"].getRgb()[:3] == (255, 235, 156)

    win.close()
    print("SMOKE TEST PASS")


if __name__ == "__main__":
    main()
