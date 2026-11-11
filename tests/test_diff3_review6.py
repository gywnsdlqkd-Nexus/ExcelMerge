# -*- coding: utf-8 -*-
"""3파일 병합 — 6차 검증에서 재현된 경우들. 합성 데이터, offscreen."""
import random
import time

import openpyxl
import pytest
from PyQt5.QtWidgets import QApplication, QMessageBox

from excelmerge.diff3_engine import S3_CONFLICT, compute_diff3


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

@pytest.mark.parametrize("extra", [4, 10])
def test_constant_columns_do_not_dilute_the_pairing(extra):
    """빈 열·늘 같은 값 열이 많으면 약한 짝 둘이 강한 짝 하나를 이겨 진짜 충돌이 숨었다."""
    H = ["ID", "Item", "Cnt", "G"] + [f"E{i}" for i in range(extra)]
    pad = [""] * extra
    base = [H, ["1", "gold", "100", "A"] + pad, ["1", "gem", "5", "A"] + pad]
    left = [H, ["1", "potion", "1", "A"] + pad, ["1", "gold", "150", "A"] + pad]
    right = [H, ["1", "gold", "200", "A"] + pad, ["1", "gem", "5", "A"] + pad]
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert (2, 1, 1) in meta and m[meta.index((2, 1, 1))][2][0] == S3_CONFLICT


def test_big_pool_with_an_inserted_row_and_bulk_edit_does_not_shift():
    H = ["Pool", "Item", "Count", "Rate"] + [f"X{i}" for i in range(36)]
    tail = [f"x{c}" for c in range(36)]
    base = [H] + [["7", f"item_{i:04d}", str(i % 60), str(100 + i * 13)] + tail
                  for i in range(360)]
    left = [H, ["7", "item_new", "1", "1"] + tail] + \
        [[r[0], r[1], r[2], str(int(r[3]) + 5)] + r[4:] for r in base[1:]]
    right = [r[:] for r in base]
    right[180][2] = "9999"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert (181, 180, 180) in meta
    assert _conf(m) == 0


def test_rows_of_mostly_default_values_do_not_cross_pair():
    """'0' 이 대부분인 묶음을 양쪽이 일괄 수정하면 우연히 같아진 행끼리 짝지어 거짓 충돌이 수천 칸 났다."""
    H = ["G"] + [f"C{i}" for i in range(19)]
    rng = random.Random(1)
    base = [H]
    for g in range(4):
        base += [[str(g)] + [("0" if rng.random() < 0.85 else str(rng.randint(1, 9)))
                             for _ in range(19)] for _ in range(500)]
    left = [r[:] for r in base]
    right = [r[:] for r in base]
    for r in left[1:]:
        r[2] = str(int(r[2]) + 1)
    for r in right[1:]:
        r[4] = str(int(r[4]) + 2)
    t0 = time.perf_counter()
    m, _meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 5.0
    assert _conf(m) == 0


def test_moving_a_unique_row_does_not_split_the_following_group():
    H = ["ID", "Item", "Cnt"]
    base = [H, ["A", "a", ""], ["1", "gold", "10"], ["1", "gem", "5"], ["B", "b", ""],
            ["2", "x", "1"], ["2", "y", "2"]]
    left = [H, ["1", "gold", "10"], ["1", "gem", "5"], ["A", "a", ""], ["B", "b", ""],
            ["2", "x", "1"], ["2", "y", "2"]]
    right = [r[:] for r in base]
    right[3] = ["1", "gem", "30"]
    m, _meta = compute_diff3(left, base, right, 0, 0)
    changed = sum(1 for row in m for c in row if c[0] != "same")
    assert changed == 1 and _conf(m) == 0


# ── 화면 ────────────────────────────────────────────────────────────────────

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


def test_reordering_duplicate_key_groups_is_reported(win, tmp_path):
    H = ["GroupID", "Item", "Cnt"]
    base = [H, [100, "gold", 10], [200, "x", 1], [100, "gem", 5], [200, "y", 2], [300, "p", 1]]
    left = [H, [100, "gold", 10], [100, "gem", 5], [200, "x", 1], [200, "y", 2], [300, "p", 1]]
    right = [H, [100, "gold", 10], [200, "x", 1], [100, "gem", 5], [200, "y", 2], [300, "p", 7]]
    l = _xlsx(tmp_path / "l.xlsx", {"Data": left})
    b = _xlsx(tmp_path / "b.xlsx", {"Data": base})
    r = _xlsx(tmp_path / "r.xlsx", {"Data": right})
    view = win.open_merge3(l, b, r, str(tmp_path / "res.xlsx"))
    assert _wait(lambda: _compared(view))
    view._stage_plan({k: v for k, v in
                      __import__("excelmerge.diff3_engine", fromlist=["x"]).auto_merge_plan(
                          view._matrix).items()}, undo=False)
    view._save()
    asked = [q for q in win.shown if q[0] == "question"]
    assert asked and "행 순서" in asked[0][2], asked


def test_three_way_does_not_inherit_the_global_key_column(win, tmp_path):
    from excelmerge.prefs import save_key_prefs
    save_key_prefs(0, 1)                           # 2파일에서 B열을 키로 잡아 둔 상태
    H = ["ID", "Name", "V"]
    l = _xlsx(tmp_path / "l.xlsx", {"Data": [H, [1, "a", "x"], [2, "b2", "y"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"Data": [H, [1, "a", "x"], [2, "b", "y"]]})
    r = _xlsx(tmp_path / "r.xlsx", {"Data": [H, [1, "a", "x"], [2, "b", "Y"]]})
    view = win.open_merge3(l, b, r, str(tmp_path / "res.xlsx"))
    assert _wait(lambda: _compared(view))
    assert view._key_col == 0
    assert all(None not in mm for mm in view._row_meta), "ID 가 아닌 열로 행을 맞췄다"


def test_three_way_uses_the_key_remembered_for_the_file(win, tmp_path):
    from excelmerge.prefs import save_last_key
    H = ["Group", "Code", "V"]
    l = _xlsx(tmp_path / "l.xlsx", {"Data": [H, ["g", "c1", "x"], ["g", "c2", "Y"]]})
    b = _xlsx(tmp_path / "b.xlsx", {"Data": [H, ["g", "c1", "x"], ["g", "c2", "y"]]})
    r = _xlsx(tmp_path / "r.xlsx", {"Data": [H, ["g", "c1", "x"], ["g", "c2", "y"]]})
    res = str(tmp_path / "res.xlsx")
    save_last_key(res, 0, 1, "Code")
    view = win.open_merge3(l, b, r, res)
    assert _wait(lambda: _compared(view))
    assert view._key_col == 1
