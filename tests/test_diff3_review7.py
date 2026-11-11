# -*- coding: utf-8 -*-
"""3파일 병합 — 7차 검증에서 재현된 경우들. 합성 데이터, offscreen."""
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

def test_rotating_a_block_in_a_big_group_keeps_rows_paired():
    """큰 묶음에서 위 45행을 빼고 끝에 45행을 붙이면 띠 밖이라 묶음 전체가 한 칸씩 밀렸다."""
    H = ["Pool", "ItemID", "Price", "W"]
    base = [H] + [["1", f"it{i:04d}", str(10 + i), "w"] for i in range(300)]
    left = [H] + base[46:] + [["1", f"new{i:03d}", str(i), "w"] for i in range(45)]
    right = [r[:] for r in base]
    right[150][2] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert _conf(m) == 0
    row = m[meta.index(next(x for x in meta if x[1] == 150))]
    assert row[1][2] == "it0149" and row[2][3] == "7777"


def test_an_unchanged_row_beats_two_weak_pairs():
    """복사해 옆에 넣은 새 행 + 지운 이웃 — 다른 쪽 수정이 새 행에 붙었다."""
    H = ["PoolID", "ItemType", "ItemID", "Count", "Weight", "Desc"]
    base = [H, ["1", "3", "100", "1", "50", ""], ["1", "5", "2200", "10", "50", ""],
            ["1", "2", "9", "100", "50", ""], ["1", "3", "300", "1", "50", ""]]
    left = [H, ["1", "3", "100", "1", "50", ""], ["1", "5", "200", "10", "50", ""],
            ["1", "5", "2200", "10", "50", ""], ["1", "3", "300", "1", "50", ""]]
    right = [r[:] for r in base]
    right[2][3] = "15"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert (3, 2, 2) in meta, meta
    assert m[meta.index((3, 2, 2))][3][0] == "right"


def test_big_groups_with_scattered_edits_stay_fast():
    rng = random.Random(3)
    H = ["Pool"] + [f"C{i}" for i in range(19)]
    base = [H]
    for p in range(50):
        base += [[str(p)] + [str(rng.randint(0, 999)) for _ in range(19)] for _ in range(200)]
    left = [r[:] for r in base]
    right = [r[:] for r in base]
    for i in rng.sample(range(1, len(base)), 500):
        left[i][3] = "L" + left[i][3]
    for i in rng.sample(range(1, len(base)), 500):
        right[i][7] = "R" + right[i][7]
    t0 = time.perf_counter()
    m, _meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 2.0
    assert _conf(m) == 0


def test_a_20k_row_blank_key_sheet_with_few_edits_stays_fast():
    rng = random.Random(3)
    H = ["K", "A", "B", "C"]
    base = [H] + [["", f"a{i}", str(rng.randint(0, 99)), "c"] for i in range(20000)]
    left = [r[:] for r in base]
    right = [r[:] for r in base]
    for i in rng.sample(range(1, 20001), 10):
        left[i][2] = "L"
    for i in rng.sample(range(1, 20001), 10):
        right[i][3] = "R"
    t0 = time.perf_counter()
    m, _meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 3.0
    assert _conf(m) == 0 and sum(1 for row in m for c in row if c[0] != "same") == 20


# ── 화면 ────────────────────────────────────────────────────────────────────

@pytest.fixture
def win(qapp, monkeypatch):
    from excelmerge import diff3_view as dv
    from excelmerge.main_window import MainWindow
    for name in ("information", "warning", "critical"):
        monkeypatch.setattr(dv.QMessageBox, name, staticmethod(lambda *a, **k: None))
    monkeypatch.setattr(dv.QMessageBox, "question", staticmethod(lambda *a, **k: QMessageBox.Yes))
    w = MainWindow()
    w.show()
    yield w
    w.close()
    w.deleteLater()
    qapp.processEvents()


def _compared(view):
    return bool(view._matrix) and not view._computing and \
        view.status.currentMessage().startswith("비교 완료")


def _three(tmp_path, tag, sheets_l, sheets_b, sheets_r):
    return (_xlsx(tmp_path / f"{tag}_l.xlsx", sheets_l), _xlsx(tmp_path / f"{tag}_b.xlsx", sheets_b),
            _xlsx(tmp_path / f"{tag}_r.xlsx", sheets_r))


def test_new_files_in_the_same_tab_do_not_keep_the_old_key(win, tmp_path):
    X = [["Type", "Code", "V"], ["t", "c1", "x"], ["t", "c2", "y"]]
    xl, xb, xr = _three(tmp_path, "x", {"Data": X}, {"Data": X}, {"Data": X})
    view = win.open_merge3(xl, xb, xr)
    assert _wait(lambda: _compared(view))
    view._on_key_col_changed(1)
    assert _wait(lambda: _compared(view) and view._key_col == 1)
    Y = [["ID", "Name", "V"], [1, "a", "x"], [2, "b", "y"]]
    yl, yb, yr = _three(tmp_path, "y", {"Data": Y}, {"Data": Y}, {"Data": Y})
    for role, p in (("l", yl), ("b", yb), ("r", yr)):
        view.panels[role].set_path(p)
    view._run_refresh()
    assert _wait(lambda: _compared(view) and view.paths()[1] == yb)
    assert view._key_col == 0, "앞 표의 키 열이 남았다"


def test_a_key_chosen_on_one_sheet_follows_by_name_to_another(win, tmp_path):
    item = [["Type", "ItemID", "Name"], ["w", "1001", "sword"], ["a", "1002", "plate"]]
    text_b = [["ItemID", "Ko", "En"], ["1001", "검", "sword"], ["1002", "갑옷", "plate"]]
    l, b, r = _three(tmp_path, "s", {"Item": item, "ItemText": text_b},
                     {"Item": item, "ItemText": text_b}, {"Item": item, "ItemText": text_b})
    view = win.open_merge3(l, b, r, str(tmp_path / "res.xlsx"))
    assert _wait(lambda: _compared(view))
    if view._current_sheet != "Item":
        view.sheet_tabs.setCurrentIndex(view._sheet_names.index("Item"))
        assert _wait(lambda: view._current_sheet == "Item" and _compared(view))
    view._on_key_col_changed(1)                    # ItemID
    assert _wait(lambda: _compared(view) and view._key_col == 1)
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("ItemText"))
    assert _wait(lambda: view._current_sheet == "ItemText" and _compared(view))
    assert view._key_col == 0, "다른 시트에서 키가 이름(ItemID)이 아니라 번호로 따라갔다"


def test_positional_columns_still_use_the_remembered_key(win, tmp_path):
    from excelmerge.prefs import save_last_key
    T = [["Memo", "ID", "", "V"], ["m", "1", "z", "x"], ["m", "2", "z", "y"]]
    l, b, r = _three(tmp_path, "p", {"Data": T}, {"Data": T}, {"Data": T})
    res = str(tmp_path / "res.xlsx")
    save_last_key(res, 0, 1, "ID")
    view = win.open_merge3(l, b, r, res)
    assert _wait(lambda: _compared(view))
    assert view._col_match is None
    assert view._key_col == 1


# ── 8차 검증 ─────────────────────────────────────────────────────────────────

def test_a_remembered_key_name_missing_on_another_sheet_falls_back_to_a(win, tmp_path):
    """한 시트에서 고른 키 이름(ItemID)이 다른 시트에 없으면 열 번호(B)로 따라가 같은 ID 두 행이 저장됐다."""
    from excelmerge.loaders import clear_values_cache, load_values_any
    item = [["Type", "ItemID", "Name"], ["w", "1001", "sword"], ["a", "1002", "plate"]]
    gb = [["Grade", "Color", "Desc"], [1, "red", "a"], [2, "blue", "b"], [3, "green", "c"]]
    gl = [["Grade", "Color", "Desc"], [1, "red", "a"], [2, "blue", "b"], [3, "lime", "c"]]
    gr = [["Grade", "Color", "Desc"], [1, "red", "a"], [2, "blue", "b"], [3, "green", "c2"]]
    l, b, r = _three(tmp_path, "g", {"Item": item, "Grade": gl}, {"Item": item, "Grade": gb},
                     {"Item": item, "Grade": gr})
    res = str(tmp_path / "res.xlsx")
    view = win.open_merge3(l, b, r, res)
    assert _wait(lambda: _compared(view))
    if view._current_sheet != "Item":
        view.sheet_tabs.setCurrentIndex(view._sheet_names.index("Item"))
        assert _wait(lambda: view._current_sheet == "Item" and _compared(view))
    view._on_key_col_changed(1)
    assert _wait(lambda: _compared(view) and view._key_col == 1)
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("Grade"))
    assert _wait(lambda: view._current_sheet == "Grade" and _compared(view))
    assert view._key_col == 0, "기억한 이름이 없는 시트에서 열 번호로 따라갔다"
    view._save()
    assert _wait(lambda: not view._saving_now())
    clear_values_cache()
    got = load_values_any(res, None, "Grade")
    assert [row for row in got if row[0] == "3"] == [["3", "lime", "c2"]], got


def test_moving_a_block_while_renumbering_every_row_keeps_pairs():
    """큰 묶음에서 블록을 옮기며 정렬 번호를 모두 다시 매기면 닻이 없어 짝이 밀렸다."""
    H = ["ShopType", "ItemID", "Sort", "Count", "Price"]
    base = [H] + [["1", str(100000 + i), str(i + 1), str(i % 7 + 1), str(100 + i)] for i in range(150)]
    moved = base[46:] + base[1:46]
    left = [H] + [[r[0], r[1], str(k + 1), r[3], r[4]] for k, r in enumerate(moved)]
    right = [r[:] for r in base]
    ri = next(i for i, r in enumerate(right) if r[1] == "100080")
    right[ri][3] = "77"
    m, _meta = compute_diff3(left, base, right, 0, 0)
    assert _conf(m) == 0
    row = next(row for row in m if row[1][2] == "100080")
    assert row[3][3] == "77" and row[3][0] == "right"


def test_hundred_row_groups_stay_fast():
    rng = random.Random(5)
    H = ["Pool"] + [f"C{i}" for i in range(19)]
    base = [H]
    for p in range(100):
        base += [[str(p)] + [str(rng.randint(0, 999)) for _ in range(19)] for _ in range(100)]
    left = [r[:] for r in base]
    right = [r[:] for r in base]
    for i in rng.sample(range(1, len(base)), 500):
        left[i][3] = "L" + left[i][3]
    for i in rng.sample(range(1, len(base)), 500):
        right[i][7] = "R" + right[i][7]
    t0 = time.perf_counter()
    m, _meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 1.5
    assert _conf(m) == 0
