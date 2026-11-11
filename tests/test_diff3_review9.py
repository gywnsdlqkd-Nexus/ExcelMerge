# -*- coding: utf-8 -*-
"""3파일 병합 — 10차 검증에서 재현된 경우들. 합성 데이터, offscreen."""
import random
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


# ── 엔진: 닻은 이웃보다 비슷한 짝만 ─────────────────────────────────────────

@pytest.mark.parametrize("seed", range(4))
@pytest.mark.parametrize("kind", ["ins", "del"])
def test_reward_table_with_renumbered_seq_keeps_pairs(seed, kind):
    """거의 같은 값 열(RewardType=Item, Count=1) 덕에 한 칸 밀린 짝이 닻으로 고정됐다."""
    rng = random.Random(seed)
    H = ["GroupID", "Seq", "RewardType", "RewardID", "Count", "Weight"]
    rows = []
    for i in range(150):
        if rng.random() < 0.3:
            t = rng.choice(["Gold", "Gem"])
            rid, cnt = ("1" if t == "Gold" else "2"), str(rng.randint(100, 999))
        else:
            t, rid, cnt = "Item", str(rng.randint(10000, 99999)), "1"
        rows.append(["7", str(i + 1), t, rid, cnt, str(rng.choice([10, 20, 50, 100]))])
    base = [H] + rows
    body = [r[:] for r in rows]
    if kind == "ins":
        body.insert(9, ["7", "x", "Item", "55555", "1", "10"])
    else:
        del body[9]
    left = [H] + [[r[0], str(k + 1)] + r[2:] for k, r in enumerate(body)]
    right = [r[:] for r in base]
    ti = next(i for i in range(120, 150) if right[i][2] == "Item")
    right[ti][5] = "777"
    m, _meta = compute_diff3(left, base, right, 0, 0)
    row = next(row for row in m if row[3][2] == right[ti][3])
    assert row[5][0] == "right" and row[5][3] == "777" and row[3][1] == right[ti][3]


def test_balance_pass_on_a_blank_key_sheet_keeps_pairs_and_stays_fast():
    """스탯 열 일괄 수정 + 맨 위 70행 추가 — 진짜 ID 짝까지 떨어져 위치로 짝지어 밀렸다."""
    H = ["#", "ID", "Name", "Grade"] + [f"S{i}" for i in range(10)] + [f"U{i}" for i in range(20)]
    rng = random.Random(1)
    base = [H] + [["", str(1000000 + i), f"mon{i}", str(rng.randint(1, 5))]
                  + [str(rng.randint(10, 999)) for _ in range(10)] + ["0"] * 20 for i in range(10000)]
    left = [H] + [["", f"9{i:06d}", f"new{i}", "1"] + ["1"] * 10 + ["0"] * 20 for i in range(70)] + \
        [r[:4] + [str(int(int(v) * 1.1)) for v in r[4:14]] + r[14:] for r in base[1:]]
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[1] == "1004999")
    right[ti][2] = "Renamed"
    t0 = time.perf_counter()
    m, _meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 5.0
    assert _conf(m) == 0
    row = next(row for row in m if row[1][2] == "1004999")
    assert row[2][0] == "right" and row[2][3] == "Renamed"


# ── 화면: 기억이 없는 시트는 1행 머리글 ─────────────────────────────────────

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


INFO = [["설명: 상점 기본값", "", ""], ["Key", "Value", "Note"], ["a", "1", ""], ["b", "2", ""]]
SHOP_B = [["ShopID", "ItemID", "Cnt", "Lim", "Flag", "Price"], ["S1", "1001", 0, 0, 0, 100],
          ["S2", "1002", 5, 1, 0, 200], ["S3", "1003", 3, 2, 1, 300], ["S4", "1004", 7, 3, 0, 400]]
SHOP_L = [["ShopID", "ItemID", "Bonus", "Cnt", "Lim", "Flag", "Price"], ["S1", "1001", 0, 0, 0, 0, 100],
          ["S2", "1002", 9, 5, 1, 0, 200], ["S3", "1003", 8, 3, 2, 1, 300], ["S4", "1004", 6, 7, 3, 0, 400]]
SHOP_R = [r[:] for r in SHOP_B]
SHOP_R[3] = ["S3", "1003", 3, 2, 1, 350]


def test_a_sheet_without_memory_does_not_inherit_the_previous_key_row(win, tmp_path):
    from excelmerge.prefs import save_sheet_key
    l = _xlsx(tmp_path / "l.xlsx", {"Info": INFO, "Shop": SHOP_L})
    b = _xlsx(tmp_path / "b.xlsx", {"Info": INFO, "Shop": SHOP_B})
    r = _xlsx(tmp_path / "r.xlsx", {"Info": INFO, "Shop": SHOP_R})
    res = str(tmp_path / "res.xlsx")
    save_sheet_key(res, "Info", 1, 0, "Key")
    view = win.open_merge3(l, b, r, res)
    assert _wait(lambda: _compared(view))
    if view._current_sheet != "Info":
        view.sheet_tabs.setCurrentIndex(view._sheet_names.index("Info"))
        assert _wait(lambda: view._current_sheet == "Info" and _compared(view))
    assert view._key_row == 1
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("Shop"))
    assert _wait(lambda: view._current_sheet == "Shop" and _compared(view))
    assert view._key_row == 0, "앞 시트의 키 행이 넘어왔다"
    view._save()
    assert _wait(lambda: not view._saving_now())
    clear_values_cache()
    got = load_values_any(res, None, "Shop")
    assert got[0] == ["ShopID", "ItemID", "Cnt", "Lim", "Flag", "Price"]
    assert got[3] == ["S3", "1003", "3", "2", "1", "350"]


def test_three_way_does_not_inherit_the_global_key_row(win, tmp_path):
    from excelmerge.prefs import save_key_prefs
    save_key_prefs(1, 0)                             # 2파일에서 다른 표의 키 행을 2행으로 둔 상태
    l = _xlsx(tmp_path / "l.xlsx", {"Shop": SHOP_L})
    b = _xlsx(tmp_path / "b.xlsx", {"Shop": SHOP_B})
    r = _xlsx(tmp_path / "r.xlsx", {"Shop": SHOP_R})
    view = win.open_merge3(l, b, r, str(tmp_path / "res.xlsx"))
    assert _wait(lambda: _compared(view))
    assert view._key_row == 0
