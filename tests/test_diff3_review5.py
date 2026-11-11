# -*- coding: utf-8 -*-
"""3파일 병합 — 5차 검증에서 재현된 경우들. 합성 데이터, offscreen."""
import time

import openpyxl
import pytest
from PyQt5.QtWidgets import QApplication, QMessageBox

from excelmerge.diff3_engine import S3_CONFLICT, compute_diff3
from excelmerge.loaders import clear_values_cache, load_values_any


def _xlsx(path, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    wb.save(str(path))
    return str(path)


def _vals(path, sheet="Data"):
    clear_values_cache()
    return load_values_any(path, None, sheet)


def _wait(pred, timeout=20.0):
    app = QApplication.instance()
    end = time.time() + timeout
    while time.time() < end:
        app.processEvents()
        if pred():
            return True
        time.sleep(0.01)
    return bool(pred())


def _conf(m):
    return sum(1 for row in m for c in row if c[0] == S3_CONFLICT)


# ── 엔진 ────────────────────────────────────────────────────────────────────

def test_bulk_edited_pool_with_an_inserted_row_does_not_shift():
    """묶음 전체를 일괄 수정하고 맨 위에 한 행을 끼우면 짝짓기 상한에 걸려 한 칸씩 밀렸다."""
    H = ["Pool", "Item", "Count", "Rate"]
    base = [H] + [["7", f"item_{i:03d}", str(i), str(100 + i * 13)] for i in range(80)]
    left = [H, ["7", "item_new", "1", "1"]] + [["7", r[1], r[2], str(int(r[3]) + 5)]
                                              for r in base[1:]]
    right = [r[:] for r in base]
    right[40][2] = "999"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert (41, 40, 40) in meta
    assert _conf(m) == 0


def test_bulk_edits_in_many_groups_stay_fast():
    H = ["G"] + [f"C{i}" for i in range(39)]
    base = [H]
    for g in range(40):
        base += [[str(g)] + [f"v{g}_{r}_{c}" for c in range(39)] for r in range(50)]
    left = [r[:] for r in base]
    right = [r[:] for r in base]
    for r in left[1:]:
        r[3] = "L" + r[3]
    for r in right[1:]:
        r[5] = "R" + r[5]
    t0 = time.perf_counter()
    m, _meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 3.0
    assert _conf(m) == 0


def test_same_children_under_different_new_parents_stay_apart():
    H = ["ID", "Name", "Item", "Cnt"]
    base = [H, ["1001", "stage1", "", ""], ["", "", "gold", "100"]]
    left = base + [["1003", "stage3", "", ""], ["", "", "gold", "100"]]
    right = base + [["1004", "stage4", "", ""], ["", "", "gold", "100"]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    new = [mm for mm in meta if mm[1] is None]
    assert len(new) == 4, f"다른 부모의 자식이 하나로 합쳐졌다: {new}"


def test_repeated_duplicate_key_separators_stay_in_their_sections():
    H = ["ID", "V"]
    base = [H, ["#", "-----"], ["10", "a"], ["#", "-----"], ["20", "c"], ["#", "-----"], ["30", "e"]]
    left = [H, ["10", "a"], ["#", "-----"], ["20", "c"], ["#", "-----"], ["30", "e"]]
    right = [r[:] for r in base]
    right[5] = ["#", "[END]"]
    m, _meta = compute_diff3(left, base, right, 0, 0)
    assert _conf(m) == 0


def test_two_weak_pairs_do_not_beat_one_strong_pair():
    H = ["ID", "Item", "Cnt", "G"]
    base = [H, ["1", "gold", "100", "A"], ["1", "gem", "5", "A"]]
    left = [H, ["1", "potion", "1", "A"], ["1", "gold", "150", "A"]]
    right = [H, ["1", "gold", "200", "A"], ["1", "gem", "5", "A"]]
    m, meta = compute_diff3(left, base, right, 0, 0)
    row = m[meta.index((2, 1, 1))]
    assert row[2][0] == S3_CONFLICT, "150 대 200 충돌이 숨었다"


# ── 3파일 화면 ──────────────────────────────────────────────────────────────

@pytest.fixture
def win(qapp, monkeypatch):
    from excelmerge import diff3_view as dv
    from excelmerge.main_window import MainWindow
    shown = []
    for name in ("information", "warning", "critical"):
        monkeypatch.setattr(dv.QMessageBox, name,
                            staticmethod(lambda *a, _n=name, **k: shown.append((_n, a[1]))))

    def question(*a, **k):
        shown.append(("question", a[1], a[2]))
        return QMessageBox.No
    monkeypatch.setattr(dv.QMessageBox, "question", staticmethod(question))
    w = MainWindow()
    w.show()
    w.shown = shown
    yield w
    w.close()
    w.deleteLater()
    qapp.processEvents()


def _compared(view):
    return bool(view._matrix) and not view._computing and \
        view.status.currentMessage().startswith("비교 완료")


def test_p4v_question_mentions_misplaced_new_rows_and_lost_order(win, tmp_path):
    H = ["ID", "Name", "Cnt"]
    base = [H, [100, "gold", 10], [200, "gem", 5], [300, "potion", 1], [400, "key", 2]]
    left = [H, [400, "key", 2], [100, "gold", 10], [150, "silver", 7], [200, "gem", 5],
            [300, "potion", 1]]                   # 400 을 맨 위로 옮기고 150 을 끼움
    right = [H, [100, "gold", 10], [200, "gem", 5], [300, "potion", 1], [400, "key", 3]]
    l = _xlsx(tmp_path / "l.xlsx", {"Data": left})
    b = _xlsx(tmp_path / "b.xlsx", {"Data": base})
    r = _xlsx(tmp_path / "r.xlsx", {"Data": right})
    view = win.open_merge3(l, b, r, str(tmp_path / "res.xlsx"))
    assert _wait(lambda: _compared(view))
    view._save()
    asked = [q for q in win.shown if q[0] == "question"]
    assert asked, "결과에 빠지는 것을 묻지 않았다"
    assert "시트 끝에 붙습니다" in asked[0][2] and "행 순서" in asked[0][2], asked[0][2]


def test_three_way_refresh_during_save_is_deferred(win, tmp_path):
    H = ["ID", "V"]
    l = _xlsx(tmp_path / "l.xlsx", {"Data": [H, [1, "A"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"Data": [H, [1, "a"]]})
    r = _xlsx(tmp_path / "r.xlsx", {"Data": [H, [1, "a"]]})
    view = win.open_merge3(l, b, r)
    assert _wait(lambda: _compared(view))
    view._auto_merge()
    view._save()
    assert view._save_in_flight
    t0 = time.perf_counter()
    view._run_refresh()                            # 기다리지 않고 미룬다
    assert time.perf_counter() - t0 < 0.5
    assert view._session_after_save
    assert _wait(lambda: not view._save_in_flight and _compared(view))
    assert _vals(b)[1] == ["1", "A"]


# ── 2파일 ───────────────────────────────────────────────────────────────────

@pytest.fixture
def win2(qapp, monkeypatch):
    from excelmerge import diff_view as dv
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv.QMessageBox, "information", staticmethod(lambda *a, **k: None))
    w = MainWindow()
    w.show()
    yield w
    w.close()
    w.deleteLater()
    qapp.processEvents()


def test_two_way_cannot_start_a_second_save_while_saving(win2, tmp_path):
    from excelmerge.constants import DIR_A2B
    a = _xlsx(tmp_path / "a.xlsx", {"S": [["ID", "V"], [1, "a1"], [3, "a3"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"S": [["ID", "V"], [1, "b1"], [2, "b2"], [3, "b3"]]})
    view = win2.open_file_compare(a, b)
    assert _wait(lambda: bool(view._diff_matrix) and view.status.currentMessage().startswith("비교 완료"))
    r2 = next(r for r, row in enumerate(view._diff_matrix) if row[0][2] == "2")
    view._staged[(r2, 0)] = DIR_A2B
    view._staged[(r2, 1)] = DIR_A2B
    view._set_save_btn_state()
    view._save_staged("b")
    first = view._staged_merge_worker
    r3 = next(r for r, row in enumerate(view._diff_matrix) if row[0][1] == "3")
    view._staged[(r3, 1)] = DIR_A2B                # 저장 중에 준비
    view._set_save_btn_state()
    assert not view.panel_b.save_btn.isEnabled(), "저장 중에 저장 버튼이 다시 켜졌다"
    view._save_staged("b")
    assert view._staged_merge_worker is first, "두 번째 저장이 겹쳐 시작됐다"
    assert _wait(lambda: not view._saving_now() and view._post_save_carry is None
                 and view.status.currentMessage().startswith("비교 완료"))
    assert _vals(b, "S") == [["ID", "V"], ["1", "b1"], ["3", "b3"]]


def test_two_way_new_file_during_save_does_not_freeze(win2, tmp_path):
    from excelmerge.constants import DIR_A2B
    a = _xlsx(tmp_path / "a.xlsx", {"S": [["ID", "V"], [1, "a1"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"S": [["ID", "V"], [1, "b1"]]})
    a2 = _xlsx(tmp_path / "a2.xlsx", {"S": [["ID", "V"], [1, "z1"]]})
    view = win2.open_file_compare(a, b)
    assert _wait(lambda: bool(view._diff_matrix) and view.status.currentMessage().startswith("비교 완료"))
    view._staged[(1, 1)] = DIR_A2B
    view._set_save_btn_state()
    view._save_staged("b")
    t0 = time.perf_counter()
    view.panel_a.set_path(a2)
    view._on_file_loaded("a", a2)                  # 기다리지 않고 미룬다
    assert time.perf_counter() - t0 < 0.5
    assert _wait(lambda: not view._saving_now() and bool(view._diff_matrix)
                 and view._diff_matrix[1][1][1] == "z1"
                 and view.status.currentMessage().startswith("비교 완료"))
    assert view._diff_matrix[1][1][2] == "a1", "저장 뒤의 B 를 읽어야 한다"
