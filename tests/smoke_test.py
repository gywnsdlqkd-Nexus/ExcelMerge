"""로더 → diff → 테이블 표시 E2E 스모크 테스트.

QT_QPA_PLATFORM=offscreen 환경에서 파이프라인 전체(로더 → compute_diff →
DiffTableModel 표시 → staged 오버라이드 → 색상)를 한 번에 검증한다.
유일한 E2E 경로이므로 pytest 가 수집하도록 `test_` 이름을 유지할 것.
(단독 실행도 가능: `python tests/smoke_test.py`)
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

HERE = os.path.dirname(os.path.abspath(__file__))          # tests/
ROOT = os.path.dirname(HERE)                               # 프로젝트 루트
sys.path.insert(0, ROOT)                                   # excelmerge/·진입점 import용
FIXTURE_DIR = os.path.join(HERE, "fixtures")

from excelmerge.loaders import load_values_any
from excelmerge.diff_engine import compute_diff
from excelmerge.diff_model import EXTRA_ROWS
from excelmerge.theme import DIFF_COLORS
from excelmerge.main_window import MainWindow


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


def test_smoke_pipeline():
    # 진입점 모듈이 항상 import 가능해야 한다(빌드/배포가 이 모듈을 실행한다).
    import excel_diff_merge
    assert hasattr(excel_diff_merge, "main"), "진입점 excel_diff_merge.main 없음"

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
    assert app is not None          # 위젯 생성 동안 QApplication 참조 유지
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
    test_smoke_pipeline()
