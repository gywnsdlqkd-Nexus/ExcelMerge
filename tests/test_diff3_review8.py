# -*- coding: utf-8 -*-
"""3파일 병합 — 9차 검증에서 재현된 경우들. 합성 데이터, offscreen."""
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


# ── 엔진: 다시 매긴 순번 열을 닻으로 고르지 않는다 ──────────────────────────

@pytest.mark.parametrize("kind", ["ins", "del", "move"])
def test_renumbered_sequence_column_is_not_used_as_anchor(kind):
    H = ["GroupID", "Seq", "ItemID", "Count", "Weight"] + [f"O{i}" for i in range(6)]
    base = [H] + [["1", str(i + 1), str(100000 + i), str(i % 9 + 1), str(i % 5 + 1)] + ["0"] * 6
                  for i in range(150)]
    body = [r[:] for r in base[1:]]
    if kind == "ins":
        body.insert(10, ["1", "x", "999999", "1", "1"] + ["0"] * 6)
    elif kind == "del":
        del body[10]
    else:
        body = body[1:] + body[:1]
    left = [H] + [[r[0], str(k + 1)] + r[2:] for k, r in enumerate(body)]
    right = [r[:] for r in base]
    ri = next(i for i, r in enumerate(right) if r[2] == "100070")
    right[ri][3] = "777"
    m, _meta = compute_diff3(left, base, right, 0, 0)
    assert _conf(m) == 0
    row = next(row for row in m if row[2][2] == "100070")
    assert row[3][0] == "right" and row[3][3] == "777"


@pytest.mark.parametrize("kind", ["move1", "ins_top", "del_top"])
def test_shop_table_with_sort_before_item_id_keeps_pairs(kind):
    H = ["ShopType", "Sort", "ItemID", "Count", "Price", "CostType", "ResetType", "Start", "End",
         "Limit", "A", "B"]
    n = 150
    base = [H] + [["1", str(i + 1), str(500000 + i), str(i % 4 + 1), str(100 + i % 3),
                   "G", "D", "s", "e", "5", "a", "b"] for i in range(n)]
    body = [r[:] for r in base[1:]]
    if kind == "move1":
        body = body[1:] + body[:1]
    elif kind == "ins_top":
        body.insert(0, ["1", "x", "599999", "1", "100", "G", "D", "s", "e", "5", "a", "b"])
    else:
        body = body[1:]
    left = [H] + [[r[0], str(k + 1)] + r[2:] for k, r in enumerate(body)]
    right = [r[:] for r in base]
    right[75][3] = "77"
    m, _meta = compute_diff3(left, base, right, 0, 0)
    assert _conf(m) == 0
    row = next(row for row in m if row[2][2] == base[75][2])
    assert row[3][0] == "right" and row[3][3] == "77"


# ── 화면: 키는 시트별로 기억한다 ────────────────────────────────────────────

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


def _goto(view, name):
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index(name))
    assert _wait(lambda: view._current_sheet == name and _compared(view))


def test_a_key_name_that_is_a_foreign_key_on_another_sheet_is_not_used(win, tmp_path):
    """Item 에서 고른 키 이름(ItemID)이 ShopItem 에 외래 키로 있으면 그 열이 키가 돼 같은 ShopID 두 행이 저장됐다."""
    item = [["#", "ItemID", "Name"], ["", "1001", "sword"], ["", "1002", "plate"]]
    sb = [["ShopID", "ItemID", "Price"], ["S1", "1001", "100"], ["S2", "1002", "200"]]
    sl = [["ShopID", "ItemID", "Price"], ["S1", "1001", "100"], ["S2", "1004", "200"]]
    sr = [["ShopID", "ItemID", "Price"], ["S1", "1001", "100"], ["S2", "1002", "250"]]
    l = _xlsx(tmp_path / "l.xlsx", {"Item": item, "ShopItem": sl})
    b = _xlsx(tmp_path / "b.xlsx", {"Item": item, "ShopItem": sb})
    r = _xlsx(tmp_path / "r.xlsx", {"Item": item, "ShopItem": sr})
    res = str(tmp_path / "res.xlsx")
    view = win.open_merge3(l, b, r, res)
    assert _wait(lambda: _compared(view))
    if view._current_sheet != "Item":
        _goto(view, "Item")
    view._on_key_col_changed(1)
    assert _wait(lambda: _compared(view) and view._key_col == 1)
    _goto(view, "ShopItem")
    assert view._key_col == 0, "다른 시트의 키 이름(외래 키 열)이 키가 됐다"
    view._save()
    assert _wait(lambda: not view._saving_now())
    clear_values_cache()
    got = load_values_any(res, None, "ShopItem")
    assert [row for row in got if row[0] == "S2"] == [["S2", "1004", "250"]], got


def test_a_two_way_file_memory_is_not_used_on_multi_sheet_workbooks(win, tmp_path):
    from excelmerge.prefs import save_last_key
    item = [["ItemID", "Name"], ["1001", "sword"]]
    shop = [["ShopID", "ItemID", "Price"], ["S1", "1001", "100"], ["S2", "1002", "200"]]
    l = _xlsx(tmp_path / "l.xlsx", {"Item": item, "ShopItem": shop})
    b = _xlsx(tmp_path / "b.xlsx", {"Item": item, "ShopItem": shop})
    r = _xlsx(tmp_path / "r.xlsx", {"Item": item, "ShopItem": shop})
    res = str(tmp_path / "res.xlsx")
    save_last_key(res, 0, 0, "ItemID")            # 2파일이 남긴 파일 단위 기억
    view = win.open_merge3(l, b, r, res)
    assert _wait(lambda: _compared(view))
    if view._current_sheet != "ShopItem":
        _goto(view, "ShopItem")
    assert view._key_col == 0


def test_a_key_chosen_in_three_way_is_remembered_for_that_sheet(win, tmp_path):
    T = [["Type", "Code", "V"], ["t", "c1", "x"], ["t", "c2", "y"]]
    l = _xlsx(tmp_path / "l.xlsx", {"A": T, "B": T})
    b = _xlsx(tmp_path / "b.xlsx", {"A": T, "B": T})
    r = _xlsx(tmp_path / "r.xlsx", {"A": T, "B": T})
    res = str(tmp_path / "res.xlsx")
    view = win.open_merge3(l, b, r, res)
    assert _wait(lambda: _compared(view))
    _goto(view, "A") if view._current_sheet != "A" else None
    view._on_key_col_changed(1)
    assert _wait(lambda: _compared(view) and view._key_col == 1)
    _goto(view, "B")
    assert view._key_col == 0, "다른 시트에 앞 시트의 키가 따라갔다"
    _goto(view, "A")
    assert view._key_col == 1, "그 시트에서 고른 키를 기억하지 못했다"
