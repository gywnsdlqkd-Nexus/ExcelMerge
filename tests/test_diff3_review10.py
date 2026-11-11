# -*- coding: utf-8 -*-
"""3파일 병합 — 11차 검증에서 재현된 경우들. 합성 데이터, offscreen."""
import random
import time

import openpyxl
import pytest
from PyQt5.QtWidgets import QApplication, QMessageBox

from excelmerge.diff3_engine import S3_CONFLICT, auto_merge_plan, compute_diff3


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


def _within(t_spent, limit, left_base_right, key_col=0):
    """속도 한도 — 절대 한도 또는 같은 데이터의 '바뀐 것 없는 비교' 시간의 4배 가운데 큰 쪽. 다른 일이 돌거나 느린
    러너에서 코드와 무관하게 거짓 실패하지 않게(검토 실측: 릴리스 게이트가 5초 한도에 걸림)."""
    _l, b, _r = left_base_right
    t0 = time.perf_counter()
    compute_diff3(b, b, b, key_col, 0)
    return t_spent < max(limit, 4 * (time.perf_counter() - t0))


def _conf(m):
    return sum(1 for row in m for c in row if c[0] == S3_CONFLICT)


# ── 엔진: 다시 매긴 순번 열로 k칸 밀린 짝 ───────────────────────────────────

def _reward_rows(rng, n):
    rows = []
    for i in range(n):
        if rng.random() < 0.3:
            t = rng.choice(["Gold", "Gem"])
            rid, cnt = ("1" if t == "Gold" else "2"), str(rng.randint(100, 999))
        else:
            t, rid, cnt = "Item", str(rng.randint(10000, 99999)), "1"
        rows.append(["7", str(i + 1), t, rid, cnt, str(rng.choice([10, 20, 50, 100]))])
    return rows


@pytest.mark.parametrize("seed", range(12))
@pytest.mark.parametrize("kind,k", [("ins", 2), ("del", 2), ("ins", 3), ("ins", 5), ("del", 5)])
def test_renumbered_seq_after_several_rows_keeps_pairs(seed, kind, k):
    """2행 이상 끼우거나 지우고 Seq 를 다시 매기면, k칸 밀린 짝이 우연히 이웃 검사를 넘어 닻이 됐다."""
    rng = random.Random(seed)
    H = ["GroupID", "Seq", "RewardType", "RewardID", "Count", "Weight"]
    rows = _reward_rows(rng, 150)
    base = [H] + rows
    body = [r[:] for r in rows]
    if kind == "ins":
        for d in range(k):
            body.insert(10, ["7", "x", "Item", str(50000 + d), "1", "10"])
    else:
        del body[10:10 + k]
    left = [H] + [[r[0], str(n + 1)] + r[2:] for n, r in enumerate(body)]
    right = [r[:] for r in base]
    ti = next(i for i in range(120, 150) if right[i][2] == "Item")
    right[ti][5] = "777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    # 가운데 쪽 RewardID 가 같은 행끼리만 짝이어야 한다 — 짝지은 행의 RewardID 가 다르면 밀린 짝
    wrong = [row[3] for row, mt in zip(m, meta)
             if mt[0] is not None and mt[1] is not None and row[3][1] != row[3][2]]
    assert not wrong, wrong[:5]
    row = next(row for row in m if row[3][2] == right[ti][3])
    assert row[5][0] == "right" and row[5][3] == "777" and row[3][1] == right[ti][3]


def test_renumbered_seq_on_a_blank_key_sheet_keeps_pairs():
    """빈 '#' 시트(시트 전체가 한 묶음)에서 300행째에 2행을 끼우고 No 를 다시 매긴다."""
    rng = random.Random(3)
    H = ["#", "No", "GroupID", "RewardType", "RewardID", "Count", "Weight"]
    rows = [["", str(i + 1), str(i // 40)] + r[2:] for i, r in enumerate(_reward_rows(rng, 2000))]
    base = [H] + rows
    body = [r[:] for r in rows]
    body.insert(300, ["", "x", "7", "Item", "50001", "1", "10"])
    body.insert(300, ["", "x", "7", "Item", "50002", "1", "10"])
    left = [H] + [[r[0], str(n + 1)] + r[2:] for n, r in enumerate(body)]
    right = [r[:] for r in base]
    ti = next(i for i in range(1500, 2000) if right[i][3] == "Item")
    right[ti][6] = "777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    wrong = [row[4] for row, mt in zip(m, meta)
             if mt[0] is not None and mt[1] is not None and row[4][1] != row[4][2]]
    assert not wrong, wrong[:5]
    row = next(row for row in m if row[4][2] == right[ti][4] and row[6][3] == "777")
    assert row[6][0] == "right"


# ── 엔진: ID 말고 행마다 다른 열이 없는 표의 일괄 수정 ────────────────────────

def _stat_table(n):
    rng = random.Random(1)
    H = ["#", "ID", "HP", "ATK", "DEF"] + [f"U{i}" for i in range(10)]
    return [H] + [["", str(1000000 + i), str(rng.randint(100, 9999)), str(rng.randint(100, 9999)),
                   str(rng.randint(100, 9999))] + ["0"] * 10 for i in range(n)]


def _balanced(rows):
    return [r[:2] + [str(int(int(v) * 1.2)) for v in r[2:5]] + r[5:] for r in rows]


def test_stat_only_table_bulk_edit_with_new_rows_keeps_pairs():
    """스탯 열을 모두 일괄 수정하면 진짜 ID 짝도 이웃과 동점이라 닻이 없어 위치로 짝지어졌다."""
    base = _stat_table(5000)
    left = [base[0]] + [["", f"9{i:06d}", "1", "1", "1"] + ["0"] * 10 for i in range(70)] + \
        _balanced(base[1:])
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[1] == "1002500")
    right[ti][5] = "5"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 3.0
    k = next(n for n, row in enumerate(m) if row[1][2] == "1002500")
    assert meta[k] == (ti + 70, ti, ti)
    assert m[k][1][0] == "same" and m[k][5][0] == "right" and m[k][5][3] == "5"
    assert _conf(m) == 0


def test_stat_only_table_bulk_edit_without_row_changes_stays_fast():
    """행 수 변화 없이 스탯만 일괄 수정 — 결과는 맞았지만 띠 정렬로 1만 행에 5초 가까이 걸렸다."""
    base = _stat_table(10000)
    left = [base[0]] + _balanced(base[1:])
    right = [r[:] for r in base]
    right[5001][5] = "5"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 2.0
    assert all(mt == (i, i, i) for i, mt in enumerate(meta))
    assert m[5001][5][0] == "right" and _conf(m) == 0


def test_big_blank_key_sheet_with_one_bulk_edited_column_stays_fast():
    """빈 키 넓은 시트(2만 행·80열, 거의 고유한 열 8개)에서 한 열을 모든 행에서 고쳤다."""
    rng = random.Random(7)
    ncol = 80
    H = ["#"] + [f"C{i}" for i in range(1, ncol)]
    base = [H]
    for i in range(20000):
        row = [""] + [f"u{c}_{i}" if c <= 8 else str(rng.randint(0, 9)) for c in range(1, ncol)]
        base.append(row)
    left = [r[:] for r in base]
    for r in left[1:]:
        r[20] = r[20] + "x"
    right = [r[:] for r in base]
    right[12345][30] = "R"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert _within(time.perf_counter() - t0, 6.0, (left, base, right))
    assert all(mt == (i, i, i) for i, mt in enumerate(meta))
    assert _conf(m) == 0


# ── 화면: 위치로 맞춘 시트를 되살릴 때 키 기준 열 ──────────────────────────

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


ITEM_B = [["아이템 번호", "", "이름", ""], ["ItemID", "Grade", "Name", "Price"],
          ["1001", "1", "Sword", "100"], ["1002", "1", "Shield", "200"], ["1003", "2", "Bow", "300"]]
ITEM_L = [r[:] for r in ITEM_B]
ITEM_L[3] = ["1002", "1", "Shield", "250"]
SHOP = [["Seq", "Tab", "ShopID", "Cost"], ["1", "A", "S1", "10"], ["2", "A", "S2", "20"]]
SHOP_R = [r[:] for r in SHOP]
SHOP_R[2] = ["2", "A", "S2", "25"]


def test_restoring_a_positional_sheet_resets_the_key_base(win, tmp_path):
    """Shop(C열 키)을 들렀다 위치로 맞춘 Item 으로 돌아와 키 행만 바꾸면 키 열이 C열로 옮겨졌다."""
    from excelmerge.prefs import load_sheet_key, save_sheet_key
    l = _xlsx(tmp_path / "l.xlsx", {"Item": ITEM_L, "Shop": SHOP})
    b = _xlsx(tmp_path / "b.xlsx", {"Item": ITEM_B, "Shop": SHOP})
    r = _xlsx(tmp_path / "r.xlsx", {"Item": ITEM_B, "Shop": SHOP_R})
    res = str(tmp_path / "res.xlsx")
    save_sheet_key(res, "Shop", 0, 2, "ShopID")
    view = win.open_merge3(l, b, r, res)
    assert _wait(lambda: _compared(view))
    if view._current_sheet != "Item":
        view.sheet_tabs.setCurrentIndex(view._sheet_names.index("Item"))
        assert _wait(lambda: view._current_sheet == "Item" and _compared(view))
    assert view._key_col == 0 and view._key_base == 0
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("Shop"))
    assert _wait(lambda: view._current_sheet == "Shop" and _compared(view))
    assert view._key_base == 2
    view.sheet_tabs.setCurrentIndex(view._sheet_names.index("Item"))
    assert _wait(lambda: view._current_sheet == "Item" and _compared(view))
    assert view._key_col == 0 and view._key_base == 0, "앞 시트의 키 기준 열이 남았다"
    view._on_key_row_changed(1)
    assert _wait(lambda: view._key_row == 1 and _compared(view))
    assert view._key_col == 0
    mem = load_sheet_key(res, "Item")
    assert mem is not None and mem[1] == 0 and mem[2] == "ItemID", mem


# ── 12차: 순번 열 다시 매김의 남은 범위 ─────────────────────────────────────

def _reward5(rng, n):
    """[GroupID, Seq, RewardType, RewardID, Count] — Item 은 Count 1, Gold/Gem 30%."""
    rows = []
    for i in range(n):
        if rng.random() < 0.3:
            t = rng.choice(["Gold", "Gem"])
            rid, cnt = ("1" if t == "Gold" else "2"), str(rng.randint(100, 999))
        else:
            t, rid, cnt = "Item", str(rng.randint(10000, 99999)), "1"
        rows.append(["7", str(i + 1), t, rid, cnt])
    return rows


def _wrong_pairs(m, meta, rid_col):
    return [row[rid_col] for row, mt in zip(m, meta)
            if mt[0] is not None and mt[1] is not None and row[rid_col][1] != row[rid_col][2]]


def _renumber(body, seq_col=1, start=1, step=1):
    return [r[:seq_col] + [str(start + step * n)] + r[seq_col + 1:] for n, r in enumerate(body)]


@pytest.mark.parametrize("seed", range(8))
@pytest.mark.parametrize("edit", ["top10", "del10", "move10", "ins9@20"])
def test_renumbered_seq_beyond_the_window_keeps_pairs(seed, edit):
    """10행 이상 끼우거나 지우거나 옮기고 Seq 를 다시 매기면, 창 밖으로 밀린 짝이 동점 판에서 닻이 됐다."""
    rng = random.Random(100 + seed)
    H = ["GroupID", "Seq", "RewardType", "RewardID", "Count"]
    rows = _reward5(rng, 200 if edit == "ins9@20" else 150)
    body = [r[:] for r in rows]
    new = [["7", "x", "Item", str(60000 + d), "1"] for d in range(10)]
    if edit == "top10":
        body = new + body
    elif edit == "del10":
        del body[5:15]
    elif edit == "move10":
        blk = body[40:50]
        del body[40:50]
        body[100:100] = blk
    else:
        body[20:20] = new[:9]
    base = [H] + rows
    left = [H] + _renumber(body)
    right = [r[:] for r in base]
    items = [i for i in range(60, len(right)) if right[i][2] == "Item"][:4]
    for i in items:
        right[i][4] = "777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    for i in items:
        row = next(row for row in m if row[3][2] == right[i][3])
        assert row[4][0] == "right" and row[4][3] == "777"


@pytest.mark.parametrize("seed", range(8))
@pytest.mark.parametrize("shape,edit", [("gap", "ins3"), ("gap", "del2"), ("sort10", "ins2")])
def test_seq_with_a_hole_renumbered_keeps_pairs(seed, shape, edit):
    """가운데 Seq 에 빈 번호(또는 Sort 에 끼운 값)가 있으면 순번 열로 안 보여 ±1 동점 판이 밀린 짝을 받았다."""
    rng = random.Random(200 + seed)
    H = ["GroupID", "Seq", "RewardType", "RewardID", "Count", "Weight"]
    rows = [r + [str(rng.choice([10, 20, 50, 100]))] for r in _reward5(rng, 150)]
    if shape == "gap":
        for n, r in enumerate(rows):
            r[1] = str(n + 1 if n < 60 else n + 2)          # 61 이 빠져 있다
    else:
        for n, r in enumerate(rows):
            r[1] = str(10 * (n + 1) if n < 50 else 10 * n + 5)  # 500 뒤에 505, 515 …
    body = [r[:] for r in rows]
    if edit == "ins3":
        body[10:10] = [["7", "x", "Item", str(60000 + d), "1", "10"] for d in range(3)]
    elif edit == "ins2":
        body[10:10] = [["7", "x", "Item", str(60000 + d), "1", "10"] for d in range(2)]
    else:
        del body[10:12]
    base = [H] + rows
    left = [H] + (_renumber(body) if shape == "gap" else _renumber(body, start=10, step=10))
    right = [r[:] for r in base]
    ti = next(i for i in range(120, len(right)) if right[i][2] == "Item")
    right[ti][5] = "777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    row = next(row for row in m if row[3][2] == right[ti][3])
    assert row[5][0] == "right" and row[5][3] == "777"


def _blank_reward(n, seed):
    rng = random.Random(seed)
    H = ["#", "No", "GroupID", "RewardType", "RewardID", "Count", "Weight"]
    rows = [["", str(i + 1), str(i // 40)] + r[2:] + [str(rng.choice([10, 20, 50, 100]))]
            for i, r in enumerate(_reward5(rng, n))]
    return H, rows


@pytest.mark.parametrize("n,k,at", [(8000, 120, 500), (10000, 70, 500), (20000, 2, 1000)])
def test_big_blank_key_sheet_with_renumbered_no_keeps_pairs_and_stays_fast(n, k, at):
    """빈 '#' 큰 시트에서 행을 끼우고 전역 No 를 다시 매기면, 띠 정렬로 수 초 걸리거나 띠를 넘어 위치로
    짝지어져 yours 수정이 다른 보상에 저장됐다."""
    H, rows = _blank_reward(n, 9)
    body = [r[:] for r in rows]
    body[at:at] = [["", "x", "999", "Item", str(500000 + d), "1", "10"] for d in range(k)]
    base = [H] + rows
    left = [H] + _renumber(body)
    right = [r[:] for r in base]
    ti = next(i for i in range(int(n * 0.75), n) if right[i][3] == "Item")
    right[ti][6] = "777"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 3.0
    assert not _wrong_pairs(m, meta, 4)
    row = next(row for row in m if row[4][2] == right[ti][4] and row[6][3] == "777")
    assert row[6][0] == "right" and _conf(m) == 0


def test_wide_blank_key_sheet_with_several_sequence_columns_stays_fast():
    """[#, No, ID(연속), Name, …] 2만 행·120열에서 한 열 일괄 수정 — 순번 열 창 검사로 느려졌다."""
    rng = random.Random(11)
    ncol = 120
    H = ["#", "No", "ID", "Name"] + [f"C{i}" for i in range(4, ncol)]
    base = [H] + [["", str(i + 1), str(3000000 + i), f"name{i}"]
                  + [str(rng.randint(0, 9)) for _ in range(4, ncol)] for i in range(20000)]
    left = [r[:] for r in base]
    for r in left[1:]:
        r[50] = r[50] + "x"
    right = [r[:] for r in base]
    right[15000][60] = "R"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert _within(time.perf_counter() - t0, 5.0, (left, base, right))
    assert all(mt == (i, i, i) for i, mt in enumerate(meta))
    assert _conf(m) == 0


# ── 13차: 순서와 상관없는 확실한 짝, 일괄 수정 열, 연번 ID ────────────────────

def _stat_wide(n, nstat, seed):
    rng = random.Random(seed)
    H = ["#", "ID"] + [f"S{i}" for i in range(1, nstat + 1)]
    return [H] + [["", str(100001 + i)] + [str(rng.randint(100000, 999999)) for _ in range(nstat)]
                  for i in range(n)]


@pytest.mark.parametrize("n,nstat,k", [(300, 8, 0), (500, 9, 3), (12000, 10, 50)])
def test_many_near_unique_stat_columns_bulk_edited_keep_id_pairs(n, nstat, k):
    """거의 고유한 스탯 열이 8개 넘으면 연번 ID 가 후보 상한에서 잘려 시트 전체가 지움+새 행·위치 짝이 됐다."""
    base = _stat_wide(n, nstat, 5)
    body = [r[:2] + [str(int(int(v) * 1.1)) for v in r[2:10]] + r[10:] for r in base[1:]]
    at = n // 8
    body[at:at] = [["", str(900001 + d)] + ["1"] * nstat for d in range(k)]
    left = [base[0]] + body
    right = [r[:] for r in base]
    ti = int(n * 0.8)
    right[ti][-1] = "777"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 3.0
    assert not _wrong_pairs(m, meta, 1)
    assert sum(1 for mt in meta if None not in mt) == n + 1
    row = next(row for row in m if row[1][2] == right[ti][1])
    if nstat > 8:
        assert row[-1][0] == "right" and row[-1][3] == "777" and _conf(m) == 0
    else:                                   # 스탯 8개 모두 theirs 가 고쳤다 — 같은 칸이라 진짜 충돌 하나
        assert row[-1][0] == "conflict" and row[-1][3] == "777" and _conf(m) == 1


REWARD_H = ["GroupID", "RewardType", "RewardID", "Count", "Weight"]


def test_sorting_a_duplicate_key_group_does_not_pair_different_rewards():
    """묶음을 RewardID 순으로 정렬만 하면 다른 보상끼리 짝지어, yours 수정이 다른 보상에 저장됐다."""
    grp = [["5003", "Item", rid, "1", w] for rid, w in
           [("60576", "50"), ("63421", "20"), ("99388", "10"), ("89405", "20"), ("52780", "10")]]
    other = [["5002", "Item", "11111", "1", "10"], ["5002", "Gold", "1", "500", "10"]]
    base = [REWARD_H] + other + grp
    left = [REWARD_H] + other + sorted(grp, key=lambda r: r[2])
    right = [r[:] for r in base]
    right[-1][3] = "5"                                       # 52780 Count 1 -> 5
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    row = next(row for row in m if row[2][2] == "52780")
    assert row[2][1] == "52780" and row[3][0] == "right" and row[3][3] == "5"
    assert _conf(m) == 0


@pytest.mark.parametrize("seed", range(10))
def test_swapping_rows_in_a_group_does_not_attach_an_edit_to_another_reward(seed):
    rng = random.Random(300 + seed)
    base = [REWARD_H]
    for g in range(10):
        for _ in range(rng.randint(5, 20)):
            if rng.random() < 0.3:
                base.append([str(g), "Gold", "1", str(rng.randint(100, 999)), "10"])
            else:
                base.append([str(g), "Item", str(rng.randint(10000, 99999)), "1", "10"])
    left = [r[:] for r in base]
    items = [i for i in range(1, len(base)) if base[i][1] == "Item"]
    ti = rng.choice(items)
    mates = [i for i in range(1, len(base)) if base[i][0] == base[ti][0] and i != ti]
    oi = rng.choice(mates)
    left[ti], left[oi] = left[oi], left[ti]                  # 같은 묶음의 두 행을 맞바꿈
    right = [r[:] for r in base]
    right[ti][3] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    bad = [row for row in m if row[3][3] == "5" and row[3][0] == "right" and row[2][1] != row[2][2]]
    assert not bad, bad


@pytest.mark.parametrize("shape", ["weight_x2_renumber", "season_no_seq_ins", "season_no_seq_del"])
def test_big_blank_key_sheet_with_a_bulk_edited_column_and_row_changes_keeps_pairs(shape):
    """한 열을 모든 행에서 고치고 행을 끼우거나 지우면, 같은 행이 없어 띠를 넘는 크기에서 위치로 짝지어졌다."""
    H, rows = _blank_reward(20000, 13)
    if shape != "weight_x2_renumber":
        H = H[:1] + ["Season"] + H[2:]
        rows = [r[:1] + ["3"] + r[2:] for r in rows]
    body = [r[:] for r in rows]
    if shape == "weight_x2_renumber":
        body = [r[:6] + [str(int(r[6]) * 2)] for r in body]
        body[6666:6666] = [["", "x", "999", "Item", str(500000 + d), "1", "20"] for d in range(15)]
        body = _renumber(body)
    elif shape == "season_no_seq_ins":
        body = [r[:1] + ["4"] + r[2:] for r in body]
        body[5000:5000] = [["", "4", "999", "Item", str(500000 + d), "1", "10"] for d in range(10)]
    else:
        body = [r[:1] + ["4"] + r[2:] for r in body]
        del body[5000:5010]
    base = [H] + rows
    left = [H] + body
    right = [r[:] for r in base]
    ti = next(i for i in range(16000, 20000) if right[i][3] == "Item")
    right[ti][5] = "5"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 3.0
    assert not _wrong_pairs(m, meta, 4)
    row = next(row for row in m if row[4][2] == right[ti][4] and row[5][3] == "5")
    assert row[5][0] == "right"


@pytest.mark.parametrize("bulk", ["count_x10", "enddate"])
def test_item_pool_renumbered_with_a_bulk_edited_column_keeps_pairs(bulk):
    """Item 전용 풀(Count 1, Weight 고정)에 10행을 끼워 No 를 다시 매기고 한 열을 일괄 수정하면, 동점 판이
    창 밖으로 밀린 짝을 닻으로 받았다."""
    rng = random.Random(21)
    H = ["#", "No", "GroupID", "RewardType", "RewardID", "Count", "Weight", "EndDate"]
    pool = [str(rng.randint(10000, 99999)) for _ in range(1500)]
    rows = [["", str(i + 1), str(i // 200), "Item", rng.choice(pool), "1", "100", "20261231"]
            for i in range(2000)]
    body = [r[:] for r in rows]
    body[300:300] = [["", "x", "1", "Item", str(500000 + d), "1", "100", "20261231"] for d in range(10)]
    if bulk == "count_x10":
        body = [r[:5] + ["10"] + r[6:] for r in body]
    else:
        body = [r[:7] + ["20270331"] for r in body]
    left = [H] + _renumber(body)
    base = [H] + rows
    right = [r[:] for r in base]
    targets = list(range(400, 420))
    for i in targets:
        right[i][6] = "777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 4)
    hits = [row for row in m if row[6][3] == "777"]
    assert len(hits) == len(targets) and all(row[6][0] == "right" for row in hits)
    assert sorted(row[4][2] for row in hits) == sorted(right[i][4] for i in targets)


def _monster(n, seed):
    rng = random.Random(seed)
    H = ["#", "ID", "Grade", "HP", "ATK", "DEF"] + [f"U{i}" for i in range(10)]
    return [H] + [["", str(1000000 + i), str(rng.randint(1, 5))]
                  + [str(rng.randint(100, 9999)) for _ in range(3)] + ["0"] * 10 for i in range(n)]


@pytest.mark.parametrize("n,edit", [(20000, "del50"), (10000, "top70"), (20000, "none")])
def test_balance_patch_with_grade_changes_keeps_id_pairs(n, edit):
    """스탯 x1.2 + 등급 일부 변경 + 행 지움/추가 — 가운데만 연속인 진짜 ID 를 순번으로 봐 모든 행이 밀렸다."""
    rng = random.Random(4)
    base = _monster(n, 8)
    body = []
    for r in base[1:]:
        r2 = r[:3] + [str(int(int(v) * 1.2)) for v in r[3:6]] + r[6:]
        if rng.random() < 0.15:
            r2[2] = str(rng.randint(1, 5))
        body.append(r2)
    if edit == "del50":
        for i in sorted(rng.sample(range(len(body)), 50), reverse=True):
            del body[i]
    elif edit == "top70":
        body = [["", f"9{d:06d}", "1", "1", "1", "1"] + ["0"] * 10 for d in range(70)] + body
    left = [base[0]] + body
    right = [r[:] for r in base]
    alive = {r[1] for r in body}
    ti = next(i for i in range(n // 2, n) if right[i][1] in alive)
    right[ti][6] = "5"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    if edit != "none":
        # 행 수 변화가 없으면(양쪽 모두 연속 ID) 띠 정렬로 수 초 걸린다 — 결과는 맞아 받아들인 한계(RG13-2)
        assert time.perf_counter() - t0 < 3.0
    assert not _wrong_pairs(m, meta, 1)
    row = next(row for row in m if row[1][2] == right[ti][1])
    assert row[6][0] == "right" and row[6][3] == "5"


def test_moving_a_block_and_renumbering_on_a_big_blank_key_sheet_keeps_pairs():
    """큰 빈 키 시트에서 100행 블록을 옮기고 No 를 다시 매긴다 — 순번 열을 '값이 같으면 순번 아님'으로 보면
    다시 틀어지던 경우(검토 실측)."""
    H, rows = _blank_reward(20000, 17)
    body = [r[:] for r in rows]
    blk = body[1000:1100]
    del body[1000:1100]
    body[4800:4800] = blk
    left = [H] + _renumber(body)
    base = [H] + rows
    right = [r[:] for r in base]
    tis = [i for i in range(1001, 1101) if right[i][3] == "Item"][:5]
    for i in tis:
        right[i][6] = "777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 4)
    hits = [row for row in m if row[6][3] == "777"]
    assert len(hits) == len(tis) and all(row[6][0] == "right" for row in hits)


# ── 14차: 작은 가운데에서 열 통계를 쓰지 않는다, 끝에 붙인 연번 ID ─────────────

SEQ_H = ["GroupID", "Seq", "RewardType", "RewardID", "Count", "Weight"]


def _seq_group(rewards, gid="5001"):
    return [[gid, str(n + 1)] + r for n, r in enumerate(rewards)]


GRP5 = [["Gold", "1", "500", "20"], ["Item", "10231", "1", "20"], ["Item", "10077", "1", "20"],
        ["Item", "10400", "1", "20"], ["Item", "10500", "1", "20"]]


@pytest.mark.parametrize("edit", ["ins_before_last", "del_2nd_last", "del_3rd_last"])
def test_renumbered_seq_near_the_end_of_a_group_keeps_pairs(edit):
    """묶음 끝 근처에 끼우거나 지우고 Seq 를 다시 매기면, 작은 가운데에서 Seq 가 ID 처럼 쓰여 위치로 짝지어졌다."""
    other = _seq_group([["Item", "11111", "1", "10"], ["Gold", "1", "300", "10"]], "5000")
    after = _seq_group([["Item", "22222", "1", "10"]], "5002")
    body = [r[:] for r in GRP5]
    if edit == "ins_before_last":
        body.insert(4, ["Item", "20012", "1", "10"])
    elif edit == "del_2nd_last":
        del body[3]
    else:
        del body[2]
    base = [SEQ_H] + other + _seq_group(GRP5) + after
    left = [SEQ_H] + other + _seq_group(body) + after
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[3] == "10500")
    right[ti][4] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    row = next(row for row in m if row[3][2] == "10500")
    assert row[3][1] == "10500" and row[4][0] == "right" and row[4][3] == "5"
    assert _conf(m) == 0


G5003 = [["Item", "60576", "1", "50"], ["Item", "63421", "1", "20"], ["Item", "99388", "1", "10"],
         ["Item", "89405", "1", "20"], ["Item", "52780", "1", "10"]]


def _g(rows, gid="5003"):
    return [[gid] + r for r in rows]


@pytest.mark.parametrize("where", ["end", "below"])
def test_replacing_a_reward_in_a_small_group_is_not_a_move(where):
    """보상 하나를 지우고 같은 설정의 새 보상을 넣으면, 작은 가운데의 열 통계로 두 행을 '이동'으로 짝지었다."""
    other = _g([["Item", "11111", "1", "10"]], "5002")
    after = _g([["Item", "33333", "1", "10"]], "5004")
    body = [r[:] for r in G5003[1:]]
    new = ["Item", "70001", "1", "50"]
    if where == "end":
        body.append(new)
    else:
        body.insert(1, new)
    base = [REWARD_H] + other + _g(G5003) + after
    left = [REWARD_H] + other + _g(body) + after
    right = [r[:] for r in base]
    right[2][3] = "5"                                         # 60576 Count 1 -> 5
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    row = next(row for row, mt in zip(m, meta) if row[2][2] == "60576")
    assert row[2][1] == "" and row[3][0] == "conflict"       # 지운 행을 고침 — 행 단위 충돌
    for rid in ("63421", "99388", "89405", "52780"):
        r2, mt = next((row, mt) for row, mt in zip(m, meta) if row[2][2] == rid)
        assert None not in mt and all(c[0] == "same" for c in r2), rid
    new_row, mt = next((row, mt) for row, mt in zip(m, meta) if row[2][1] == "70001")
    assert mt[1] is None and mt[2] is None


def test_swapping_weights_of_consecutive_reward_ids_is_not_a_row_swap():
    """연속 RewardID 묶음에서 두 행의 Weight 만 맞바꾸면, RewardID 를 순번으로 보고 행을 맞바꾼 것으로 짝지었다."""
    grp = [["Item", "1001", "1", "50"], ["Item", "1002", "1", "30"], ["Item", "1003", "1", "15"],
           ["Item", "1004", "1", "5"]]
    body = [r[:] for r in grp]
    body[0][3], body[2][3] = "15", "50"
    other = _g([["Item", "11111", "1", "10"]], "5002")
    base = [REWARD_H] + other + _g(grp)
    left = [REWARD_H] + other + _g(body)
    right = [r[:] for r in base]
    right[2][3] = "5"                                         # 1001 Count 1 -> 5
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    row = next(row for row in m if row[2][2] == "1001")
    assert row[2][0] == "same" and row[3][0] == "right" and row[4][0] == "left" and row[4][1] == "15"


@pytest.mark.parametrize("yours", ["edit", "delete"])
def test_deleting_a_reward_and_adding_one_of_the_same_grade(yours):
    """지운 보상과 같은 등급의 새 보상을 묶음 끝에 넣으면 둘을 '옮기고 RewardID 를 바꾼 같은 행'으로 짝지었다."""
    grp = [["Item", "60576", "1", "50"], ["Item", "63421", "1", "20"], ["Item", "99388", "1", "10"],
           ["Item", "89405", "1", "20"], ["Gold", "1", "300", "100"], ["Item", "52780", "1", "50"]]
    body = [r for r in grp if r[1] != "99388"] + [["Item", "70001", "1", "10"]]
    base = [REWARD_H] + _g(grp)
    left = [REWARD_H] + _g(body)
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[2] == "99388")
    if yours == "edit":
        right[ti][3] = "5"
    else:
        del right[ti]
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    new_row, mt = next((row, mt) for row, mt in zip(m, meta) if row[2][1] == "70001")
    assert mt[1] is None and mt[2] is None
    if yours == "edit":
        row = next(row for row in m if row[2][2] == "99388")
        assert row[3][0] == "conflict"
    else:
        assert _conf(m) == 0


@pytest.mark.parametrize("n,k", [(5000, 70), (20000, 30), (20000, -30)])
def test_stat_table_bulk_edit_with_rows_appended_or_cut_at_the_end(n, k):
    """연번 ID 표에 다음 ID 로 끝에 붙이거나 끝을 지우면, 동점 판이 순번 열을 건너뛰어 수 초 정렬하거나
    시트 전체가 '지움+새 행'이 됐다."""
    base = _stat_table(n)
    body = _balanced(base[1:])
    if k > 0:
        body += [["", str(1000000 + n + d), "1", "1", "1"] + ["0"] * 10 for d in range(k)]
    else:
        body = body[:k]
    left = [base[0]] + body
    right = [r[:] for r in base]
    ti = n // 2
    right[ti][5] = "5"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert time.perf_counter() - t0 < 3.0
    assert not _wrong_pairs(m, meta, 1)
    assert len(m) == 1 + max(n, n + k)
    row = next(row for row in m if row[1][2] == right[ti][1])
    assert row[5][0] == "right" and row[5][3] == "5" and _conf(m) == 0


# ── 15차: ID 열 하나로 짝, 연번 ID 와 다시 매긴 순번 가르기 ─────────────────────

@pytest.mark.parametrize("dist", ["end", "20_below"])
@pytest.mark.parametrize("yours", ["edit", "delete"])
def test_replacing_a_reward_with_one_of_the_same_weight_far_away_in_a_big_pool(dist, yours):
    """가운데가 20행 이상이면 Weight 가 ID 노릇을 해 지운 보상과 같은 확률의 새 보상을 같은 행으로 짝지었다."""
    rng = random.Random(41)
    weights = rng.sample(range(500, 30000), 30)
    grp = [["Item", str(31000 + i * 7), "1", str(weights[i])] for i in range(30)]
    gone = grp[4]
    new = ["Item", "40001", "1", gone[3]]
    body = [r[:] for r in grp if r is not gone]
    if dist == "end":
        body.append(new)
    else:
        body.insert(24, new)
    base = [REWARD_H] + _g(grp, "7001")
    left = [REWARD_H] + _g(body, "7001")
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[2] == gone[1])
    if yours == "edit":
        right[ti][3] = "2"
    else:
        del right[ti]
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    new_row, mt = next((row, mt) for row, mt in zip(m, meta) if row[2][1] == "40001")
    assert mt[1] is None and mt[2] is None
    if yours == "edit":
        row = next(row for row in m if row[2][2] == gone[1])
        assert row[3][0] == "conflict"
    else:
        assert _conf(m) == 0


@pytest.mark.parametrize("n,edit,grade", [(10000, "del_front10", False), (10000, "del_front10", True),
                                          (20000, "del_front30", False), (20000, "add_front30", False),
                                          (5000, "del_front200", False), (20000, "add_end30", True)])
def test_consecutive_ids_cut_or_added_at_the_front_keep_pairs(n, edit, grade):
    """맨 앞을 지우거나 이전 ID 로 앞에 붙이면 진짜 연번 ID 를 다시 매긴 순번으로 판정해 수 초 정렬하거나
    시트 전체가 '지움+새 행'이 되고, 등급을 바꾼 행은 이웃과 엇갈렸다."""
    rng = random.Random(6)
    base = _monster(n, 12)
    body = []
    for r in base[1:]:
        r2 = r[:3] + [str(int(int(v) * 1.2)) for v in r[3:6]] + r[6:]
        if grade and rng.random() < 0.15:
            r2[2] = str(rng.randint(1, 5))
        body.append(r2)
    k = int(edit.split("front")[-1] if "front" in edit else edit.split("end")[-1])
    if edit.startswith("del_front"):
        body = body[k:]
    elif edit.startswith("add_front"):
        body = [["", str(1000000 - k + d), "1", "1", "1", "1"] + ["0"] * 10 for d in range(k)] + body
    else:
        body += [["", str(1000000 + n + d), "1", "1", "1", "1"] + ["0"] * 10 for d in range(k)]
    left = [base[0]] + body
    right = [r[:] for r in base]
    ti = n // 2
    right[ti][6] = "5"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    if not grade:
        assert time.perf_counter() - t0 < 3.0
    assert not _wrong_pairs(m, meta, 1)
    deleted = sum(1 for mt in meta if mt[0] is None and mt[1] is not None)
    assert deleted == (k if edit.startswith("del") else 0)
    row = next(row for row in m if row[1][2] == right[ti][1])
    assert row[6][0] == "right" and row[6][3] == "5" and _conf(m) == 0


@pytest.mark.parametrize("change", ["swap_adjacent", "take_neighbours_old"])
def test_consecutive_reward_ids_with_weight_changes_on_neighbours(change):
    """연속 RewardID 묶음에서 이웃한 두 행의 확률만 바꿔도 RewardID 를 순번으로 보고 행을 엇갈려 짝지었다."""
    grp = [["Item", "1001", "1", "50"], ["Item", "1002", "1", "30"], ["Item", "1003", "1", "15"],
           ["Item", "1004", "1", "5"]]
    body = [r[:] for r in grp]
    if change == "swap_adjacent":
        body[1][3], body[2][3] = "15", "30"
    else:
        body[1][3], body[2][3] = "15", "7"
    other = _g([["Item", "11111", "1", "10"]], "5002")
    base = [REWARD_H] + other + _g(grp)
    left = [REWARD_H] + other + _g(body)
    right = [r[:] for r in base]
    right[4][3] = "5"                                         # 1003 Count 1 -> 5
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    assert all(None not in mt for mt in meta)
    row = next(row for row in m if row[2][2] == "1003")
    assert row[2][0] == "same" and row[3][0] == "right" and row[4][0] == "left"


def test_copying_a_reward_to_the_next_id_below_keeps_the_original_pair():
    """89405 를 다음 ID 89406 으로 바로 아래 복사하고 원래 행의 Weight 를 낮추면 원래 행이 새 행과 짝지어졌다."""
    body = [r[:] for r in G5003]
    body[3] = ["Item", "89405", "1", "10"]
    body.insert(4, ["Item", "89406", "1", "20"])
    other = _g([["Item", "11111", "1", "10"]], "5002")
    base = [REWARD_H] + other + _g(G5003)
    left = [REWARD_H] + other + _g(body)
    right = [r[:] for r in base]
    right[5][3] = "5"                                         # 89405 Count 1 -> 5
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    row = next(row for row in m if row[2][2] == "89405")
    assert row[2][0] == "same" and row[3][0] == "right" and row[4][0] == "left"
    new_row, mt = next((row, mt) for row, mt in zip(m, meta) if row[2][1] == "89406")
    assert mt[1] is None


@pytest.mark.parametrize("edit", ["sort", "swap_2_4"])
def test_reordering_a_small_group_and_renumbering_seq_keeps_pairs(edit):
    """묶음을 정렬하거나 두 행을 맞바꾸고 Seq 를 다시 매기면, 작은 가운데에서 Seq 로 짝지어 다른 보상에 저장됐다."""
    grp = [["Gold", "1", "500", "100"], ["Item", "60576", "1", "50"], ["Item", "63421", "1", "20"],
           ["Item", "99388", "1", "10"], ["Item", "52780", "1", "10"]]
    body = [r[:] for r in grp]
    if edit == "sort":
        body = sorted(body, key=lambda r: r[1])
    else:
        body[1], body[3] = body[3], body[1]
    base = [SEQ_H] + _seq_group(grp)
    left = [SEQ_H] + _seq_group(body)
    right = [r[:] for r in base]
    tid = "52780" if edit == "sort" else "60576"
    ti = next(i for i, r in enumerate(right) if r[3] == tid)
    right[ti][4] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    row = next(row for row in m if row[3][2] == tid)
    assert row[3][0] == "same" and row[4][0] == "right" and row[4][3] == "5" and _conf(m) == 0


# ── 16차: ID 열은 머리글 이름으로 ────────────────────────────────────────────

W_FIRST_H = ["GroupID", "Weight", "RewardType", "RewardID", "Count"]


@pytest.mark.parametrize("n", [5, 30])
@pytest.mark.parametrize("yours", ["edit", "delete"])
def test_weight_left_of_reward_id_is_not_used_as_an_id(n, yours):
    """Weight 가 RewardID 왼쪽에 있으면 Weight 가 ID 노릇을 해 지운 보상과 같은 확률의 새 보상을 짝지었다."""
    rng = random.Random(51 + n)
    weights = rng.sample(range(500, 30000), n)
    grp = [[str(weights[i]), "Item", str(31000 + i * 7), "1"] for i in range(n)]
    gone = grp[min(4, n - 2)]
    body = [r[:] for r in grp if r is not gone] + [[gone[0], "Item", "40001", "1"]]
    base = [W_FIRST_H] + _g(grp, "7001")
    left = [W_FIRST_H] + _g(body, "7001")
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[3] == gone[2])
    if yours == "edit":
        right[ti][4] = "2"
    else:
        del right[ti]
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    _new, mt = next((row, mt) for row, mt in zip(m, meta) if row[3][1] == "40001")
    assert mt[1] is None and mt[2] is None


@pytest.mark.parametrize("n,at", [(40, 10), (40, 35), (200, 50), (2000, 500), (2000, 1990)])
def test_copying_a_monster_to_the_next_id_at_the_end_keeps_the_original_pair(n, at):
    """다음 ID 로 끝에 복사하고 원래 행의 HP 를 낮추면, HP 가 ID 노릇을 해 원래 행을 복사본과 짝지었다."""
    base = _monster(n, 3)
    body = [r[:] for r in base[1:]]
    copy = body[at][:]
    copy[1] = str(1000000 + n)
    body[at][3] = str(int(body[at][3]) // 2)
    body.append(copy)
    left = [base[0]] + body
    right = [r[:] for r in base]
    right[at + 1][6] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    row = next(row for row in m if row[1][2] == right[at + 1][1])
    assert row[1][0] == "same" and row[6][0] == "right" and row[6][3] == "5" and _conf(m) == 0


@pytest.mark.parametrize("edit", ["append_next", "del_last_take_w", "del_first_take_w"])
def test_consecutive_reward_ids_with_rows_added_or_removed_at_the_ends(edit):
    """연속 RewardID 묶음의 끝에 다음 ID 를 붙이거나 첫·끝 행을 지우면 RewardID 를 순번으로 빼 엇갈렸다."""
    grp = [["Item", "1001", "1", "50"], ["Item", "1002", "1", "30"], ["Item", "1003", "1", "15"],
           ["Item", "1004", "1", "5"]]
    body = [r[:] for r in grp]
    if edit == "append_next":
        body.append(["Item", "1005", "1", "3"])
        body[2][3] = "12"
        tid = "1003"
    elif edit == "del_last_take_w":
        body[2][3] = "5"
        del body[3]
        tid = "1004"
    else:
        body[1][3] = "50"
        del body[0]
        tid = "1001"
    other = _g([["Item", "11111", "1", "10"]], "5002")
    base = [REWARD_H] + other + _g(grp)
    left = [REWARD_H] + other + _g(body)
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[2] == tid)
    right[ti][3] = "7"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    row = next(row for row in m if row[2][2] == tid)
    if edit == "append_next":
        assert row[2][0] == "same" and row[3][0] == "right"
    else:
        assert row[3][0] == "conflict"                    # 지운 행을 고침 — 행 단위 충돌


@pytest.mark.parametrize("yours", ["edit", "delete"])
def test_gold_tiers_in_a_group_do_not_make_weight_an_id(yours):
    """Gold 단계 행이 여럿이라 RewardID 가 80% 고유를 못 넘으면 Weight 가 ID 로 뽑혔다."""
    grp = [["Gold", "1", "500", "101"], ["Gold", "1", "1000", "102"], ["Gold", "1", "3000", "103"],
           ["Gold", "1", "5000", "104"], ["Gem", "2", "50", "105"], ["Item", "60576", "1", "50"],
           ["Item", "63421", "1", "20"], ["Item", "99388", "1", "10"], ["Item", "89405", "1", "30"],
           ["Item", "52780", "1", "40"]]
    body = [r for r in grp if r[1] != "99388"] + [["Item", "70001", "1", "10"]]
    base = [REWARD_H] + _g(grp)
    left = [REWARD_H] + _g(body)
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[2] == "99388")
    if yours == "edit":
        right[ti][3] = "5"
    else:
        del right[ti]
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    _new, mt = next((row, mt) for row, mt in zip(m, meta) if row[2][1] == "70001")
    assert mt[1] is None and mt[2] is None
    if yours == "delete":
        assert _conf(m) == 0


def test_partial_hp_change_does_not_pair_different_monsters():
    """HP 일부만 고쳐도 HP 가 ID 노릇을 해 다른 몬스터끼리 짝지어졌다."""
    rng = random.Random(2)
    base = _monster(1000, 2)
    body = [r[:] for r in base[1:]]
    for r in body:
        if rng.random() < 0.3:
            r[3] = str(int(int(r[3]) * 1.2))
    left = [base[0]] + body
    right = [r[:] for r in base]
    for i in range(1, len(right), 97):
        right[i][6] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    assert _conf(m) == 0 and all(None not in mt for mt in meta)


@pytest.mark.parametrize("tail", [50, 48, 52])
def test_cutting_the_top_and_appending_at_the_end_keeps_id_pairs(tail):
    """맨 위 50행을 지우고 끝에 비슷한 수만큼 붙여 행 수가 (거의) 같으면, 동점 판이 ID 를 버려 엇갈렸다."""
    rng = random.Random(9)
    n = 3000
    base = _monster(n, 21)
    body = []
    for r in base[51:]:
        r2 = r[:3] + [str(int(int(v) * 1.2)) for v in r[3:6]] + r[6:]
        if rng.random() < 0.15:
            r2[2] = str(rng.randint(1, 5))
        body.append(r2)
    body += [["", str(1000000 + n + d), "1", "1", "1", "1"] + ["0"] * 10 for d in range(tail)]
    left = [base[0]] + body
    right = [r[:] for r in base]
    right[1500][6] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    row = next(row for row in m if row[1][2] == right[1500][1])
    assert row[6][0] == "right" and _conf(m) == 0


def test_swapping_items_in_a_group_with_gold_tiers_and_renumbering_seq():
    """Gold·Gem 행이 섞여 RewardID 가 덜 고유한 묶음에서 두 아이템을 맞바꾸고 Seq 를 다시 매기면 엇갈렸다."""
    grp = [["Gold", "1", "500", "100"], ["Gold", "1", "1000", "50"], ["Gem", "2", "50", "50"],
           ["Item", "60576", "1", "50"], ["Item", "63421", "1", "20"], ["Item", "99388", "1", "10"],
           ["Item", "52780", "1", "10"], ["Gold", "1", "3000", "20"]]
    body = [r[:] for r in grp]
    body[3], body[5] = body[5], body[3]
    base = [SEQ_H] + _seq_group(grp)
    left = [SEQ_H] + _seq_group(body)
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[3] == "60576")
    right[ti][4] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    row = next(row for row in m if row[3][2] == "60576")
    assert row[3][0] == "same" and row[4][0] == "right" and _conf(m) == 0


# ── 17차: ID 이름의 묶음 순번, 분류 ID, 이름 없는 ID 열 ────────────────────────

@pytest.mark.parametrize("head", ["SubID", "SlotID", "SortKey"])
@pytest.mark.parametrize("edit", ["ins_top", "ins_mid", "del_2nd", "swap", "sort"])
def test_id_named_slot_column_renumbered_is_not_an_id(head, edit):
    """SlotID·SubID 같은 묶음 안 순번을 다시 매기면, 이름이 ID 라 그 열로 같은 번호끼리(위치로) 짝지었다."""
    H = ["GroupID", head, "RewardType", "RewardID", "Count", "Weight"]
    grp = [["Gold", "1", "500", "20"], ["Item", "60576", "1", "50"], ["Item", "63421", "1", "20"],
           ["Item", "99388", "1", "10"], ["Item", "52780", "1", "10"]]
    body = [r[:] for r in grp]
    new = ["Item", "70001", "1", "30"]
    if edit == "ins_top":
        body.insert(0, new)
    elif edit == "ins_mid":
        body.insert(1, new)
    elif edit == "del_2nd":
        del body[1]
    elif edit == "swap":
        body[1], body[3] = body[3], body[1]
    else:
        body = sorted(body, key=lambda r: r[1])
    base = [H] + _seq_group(grp)
    left = [H] + _seq_group(body)
    right = [r[:] for r in base]
    tid = "63421" if edit != "swap" else "60576"
    ti = next(i for i, r in enumerate(right) if r[3] == tid)
    right[ti][4] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    row = next(row for row in m if row[3][2] == tid)
    assert row[3][0] == "same" and row[4][0] == "right" and _conf(m) == 0


def test_slot_id_in_a_forty_row_group_with_one_insert():
    H = ["GroupID", "SlotID", "RewardType", "RewardID", "Count"]
    rng = random.Random(77)
    grp = [["Item", str(10000 + rng.randint(0, 89999)), "1"] for _ in range(40)]
    body = [r[:] for r in grp]
    body.insert(20, ["Item", "70001", "1"])
    base = [H] + [["9", str(n + 1)] + r for n, r in enumerate(grp)]
    left = [H] + [["9", str(n + 1)] + r for n, r in enumerate(body)]
    right = [r[:] for r in base]
    right[25][4] = "777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    row = next(row for row in m if row[3][2] == right[25][3])
    assert row[4] == ("right", "1", "1", "777")


def test_grade_id_left_of_item_id_is_not_used_as_the_id():
    """GradeID 가 ItemID 왼쪽이면 같은 등급의 다른 아이템끼리 짝지었다."""
    H = ["GroupID", "GradeID", "ItemID", "Count", "Weight"]
    groups = []
    for g in range(6):
        groups += [[str(5000 + g), str(k + 1), str(90000 + g * 10 + k), "1", "10"] for k in range(5)]
    base = [H] + groups
    left_rows = [r[:] for r in groups if r[2] != "90032"]
    at = max(i for i, r in enumerate(left_rows) if r[0] == "5003") + 1
    left_rows.insert(at, ["5003", "3", "70001", "1", "10"])
    left = [H] + left_rows
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[2] == "90032")
    right[ti][3] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    row = next(row for row in m if row[2][2] == "90032")
    assert row[3][0] == "conflict"


@pytest.mark.parametrize("head", ["Index", "Idx", "No", "번호"])
def test_monster_table_with_an_index_header_keeps_id_pairs(head):
    """ID 열 머리글이 Index·No 면 ID 짝이 꺼져, 밸런스 패치 + 행 지움에서 지운 몬스터와 다음 몬스터가 짝지어졌다."""
    rng = random.Random(0)
    base = _monster(1000, 0)
    base[0] = base[0][:1] + [head] + base[0][2:]
    body = []
    for r in base[1:]:
        r2 = r[:3] + [str(int(int(v) * 1.2)) for v in r[3:6]] + r[6:]
        if rng.random() < 0.15:
            r2[2] = str(rng.randint(1, 5))
        body.append(r2)
    gone = set(rng.sample(range(len(body)), 50))
    body = [r for k, r in enumerate(body) if k not in gone]
    left = [base[0]] + body
    right = [r[:] for r in base]
    tk = sorted(gone)[len(gone) // 2]
    right[tk + 1][6] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    row = next(row for row in m if row[1][2] == right[tk + 1][1])
    assert row[6][0] == "conflict"                         # 지운 행을 고침


def test_two_row_header_opened_at_row_one_still_finds_the_id_column():
    """1행이 한글 설명, 2행이 영문 머리글인 시트를 키 행 1행으로 열면 RewardID 를 못 찾아 복사본과 짝지었다."""
    desc = ["그룹", "보상 종류", "보상 번호", "수량", "확률"]
    grp = [["Item", "60576", "1", "50"], ["Item", "63421", "1", "25"], ["Item", "99388", "1", "10"],
           ["Item", "89405", "1", "15"], ["Item", "52780", "1", "5"]]
    body = [r[:] for r in grp]
    body[3] = ["Item", "89405", "1", "8"]
    body.insert(4, ["Item", "89406", "1", "15"])
    other = _g([["Item", "11111", "1", "10"]], "5002")
    base = [desc, REWARD_H] + other + _g(grp)
    left = [desc, REWARD_H] + other + _g(body)
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[2] == "89405")
    right[ti][3] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    row = next(row for row in m if row[2][2] == "89405")
    assert row[2][0] == "same" and row[3][0] == "right" and row[4][0] == "left"


# ── 18차: 대신 쓰는 ID 열은 시트 전체 고유·오름차순, 키 열은 ID 가 아니다 ──────────

NO_ID_HEADS = [["Group", "Weight", "RewardType", "Reward", "Count"],
               ["그룹", "확률", "보상 종류", "보상", "수량"]]


def _w_group(gid):
    return [[gid, w, "Item", rid, "1"] for rid, w in
            [("60576", "50"), ("63421", "25"), ("99388", "10"), ("89405", "15"), ("52780", "5")]]


@pytest.mark.parametrize("H", NO_ID_HEADS)
@pytest.mark.parametrize("edit", ["swap_weight", "replace_same_weight"])
def test_weight_right_after_the_key_is_not_a_fallback_id(H, edit):
    """ID 이름 열이 없는 보상 표에서 키 바로 오른쪽 Weight 를 대신 ID 로 써, 확률을 맞바꾸거나 같은 확률의 새
    보상으로 바꾸면 다른 보상끼리 짝지었다."""
    other = [["5002", "10", "Item", "11111", "1"]]
    grp = _w_group("5003")
    body = [r[:] for r in grp]
    if edit == "swap_weight":
        body[1][1], body[2][1] = "10", "25"
        tid = "63421"
    else:
        body = [r for r in body if r[3] != "99388"] + [["5003", "10", "Item", "70001", "1"]]
        tid = "99388"
    base = [H] + other + grp
    left = [H] + other + body
    right = [r[:] for r in base]
    ti = next(i for i, r in enumerate(right) if r[3] == tid)
    right[ti][4] = "9"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    row = next(row for row in m if row[3][2] == tid)
    if edit == "swap_weight":
        assert row[3][0] == "same" and row[4][0] == "right" and row[1][0] == "left"
    else:
        assert row[4][0] == "conflict"


@pytest.mark.parametrize("seed", range(4))
def test_stage_id_key_with_an_index_column_keeps_id_pairs(seed):
    """키 열 이름이 StageID 면 키 열이 ID 후보로 잡혀 Index 를 대신 쓰지 않았고, 지운 몬스터와 다음 몬스터가 짝지어졌다."""
    rng = random.Random(seed)
    H = ["StageID", "Index", "Grade", "HP", "ATK", "DEF"] + [f"U{i}" for i in range(6)]
    base = [H]
    k = 0
    for st in (100, 200, 300):
        for _ in range(200):
            base.append([str(st), str(1000000 + k), str(rng.randint(1, 5))]
                        + [str(rng.randint(100, 9999)) for _ in range(3)] + ["0"] * 6)
            k += 1
    body = []
    for r in base[1:]:
        r2 = r[:3] + [str(int(int(v) * 1.2)) for v in r[3:6]] + r[6:]
        if rng.random() < 0.15:
            r2[2] = str(rng.randint(1, 5))
        body.append(r2)
    gone = set(rng.sample(range(len(body)), 30))
    left = [H] + [r for i, r in enumerate(body) if i not in gone]
    right = [r[:] for r in base]
    g = sorted(gone)[len(gone) // 2]
    alive = next(i for i in range(g + 1, len(body)) if i not in gone)
    right[g + 1][6] = "5"
    right[alive + 1][7] = "6"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    r_gone = next(row for row in m if row[1][2] == right[g + 1][1])
    assert r_gone[6][0] == "conflict"
    r_alive = next(row for row in m if row[1][2] == right[alive + 1][1])
    assert r_alive[7][0] == "right" and r_alive[1][0] == "same"


@pytest.mark.parametrize("name", ["RewardIdx", "보상 번호", "Reward_No"])
def test_non_standard_id_header_names_are_recognised(name):
    """RewardIdx·보상 번호 같은 이름이면 ID 짝이 꺼져, 다음 ID 로 복사한 새 행과 원래 행을 엇갈려 짝지었다."""
    H = ["GroupID", "RewardType", name, "Count", "Weight"]
    body = [r[:] for r in G5003]
    body[3] = ["Item", "89405", "1", "10"]
    body.insert(4, ["Item", "89406", "1", "20"])
    other = _g([["Item", "11111", "1", "10"]], "5002")
    base = [H] + other + _g(G5003)
    left = [H] + other + _g(body)
    right = [r[:] for r in base]
    right[5][3] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    row = next(row for row in m if row[2][2] == "89405")
    assert row[2][0] == "same" and row[3][0] == "right"


# ── 19차: 시트 전체 행 번호, 10 간격 정렬 번호, 대신 쓰는 ID 의 순서 밖 새 ID ──────────

@pytest.mark.parametrize("head", ["No", "번호", "Index", "Idx", "#"])
@pytest.mark.parametrize("pos", [1, 5])
@pytest.mark.parametrize("edit", ["ins", "del"])
@pytest.mark.parametrize("seed", range(3))
def test_sheet_wide_row_number_renumbered_is_not_an_id(head, pos, edit, seed):
    """시트 전체 1..N 행 번호(No·번호·Index)를 다시 매기면 그 열로 짝지어 아래 묶음이 모두 한 칸씩 엇갈렸다."""
    rng = random.Random(seed)
    cols = ["GroupID", "RewardType", "RewardID", "Count", "Weight"]
    cols.insert(pos, head)
    rows = []
    for g in range(6):
        for _ in range(6):
            if rng.random() < 0.3:
                rows.append([str(5001 + g), "Gold", "1", str(rng.randint(100, 999)), "10"])
            else:
                rows.append([str(5001 + g), "Item", str(rng.randint(10000, 99999)), "1", "10"])
    body = [r[:] for r in rows]
    if edit == "ins":
        body.insert(2, ["5001", "Item", "70001", "1", "10"])
    else:
        del body[2]

    def number(rs):
        return [r[:pos] + [str(n + 1)] + r[pos:] for n, r in enumerate(rs)]
    base = [cols] + number(rows)
    left = [cols] + number(body)
    right = [r[:] for r in base]
    rid_col = cols.index("RewardID")
    cnt_col = cols.index("Count")
    ti = next(i for i in range(19, 25) if right[i][cols.index("RewardType")] == "Item")
    right[ti][cnt_col] = "999"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, rid_col)
    row = next(row for row in m if row[rid_col][2] == right[ti][rid_col])
    assert row[rid_col][0] == "same" and row[cnt_col][0] == "right" and _conf(m) == 0


@pytest.mark.parametrize("head", ["Sort_No", "SortIndex", "정렬 번호"])
@pytest.mark.parametrize("edit", ["ins", "del"])
def test_sort_number_in_steps_of_ten_renumbered_is_not_an_id(head, edit):
    """10 간격 정렬 번호를 다시 매기면 같은 번호끼리(같은 자리끼리) 짝지어 다른 보상에 저장됐다."""
    H = ["GroupID", head, "RewardType", "Reward", "Count", "Weight"]
    rng = random.Random(3)
    groups = []
    for g, n in enumerate([5, 8, 5, 8]):
        groups.append([["Item", str(rng.randint(10000, 99999)), "1", "10"] for _ in range(n)])
    body = [[r[:] for r in grp] for grp in groups]
    if edit == "ins":
        body[1].insert(1, ["Item", "70001", "1", "10"])
    else:
        del body[1][1]

    def sheet(gs):
        out = []
        for g, grp in enumerate(gs):
            out += [[str(5001 + g), str(10 * (n + 1))] + r for n, r in enumerate(grp)]
        return out
    base = [H] + sheet(groups)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    ti = 5 + 4
    right[ti][4] = "999"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    row = next(row for row in m if row[3][2] == right[ti][3])
    assert row[3][0] == "same" and row[4][0] == "right" and _conf(m) == 0


@pytest.mark.parametrize("shape", ["new_id_out_of_order", "section_rows"])
def test_fallback_id_survives_a_new_out_of_order_id_and_section_rows(shape):
    """ID 이름이 아닌 ID 열(Monster)에서 순서 밖 새 ID 한 행이나 빈칸 구간 행이 있으면 대신 쓰는 ID 가 꺼졌다."""
    rng = random.Random(3)
    H = ["#", "Monster", "Grade", "HP", "ATK", "DEF"] + [f"U{i}" for i in range(10)]
    rows = [["", str(1000000 + i), str(rng.randint(1, 5))] + [str(rng.randint(100, 9999)) for _ in range(3)]
            + ["0"] * 10 for i in range(1000)]
    if shape == "section_rows":
        for at in (750, 500, 250, 0):
            rows.insert(at, [f"Stage{at // 250 + 1}"] + [""] * 15)
    base = [H] + rows
    body = []
    for r in rows:
        if r[1] == "":
            body.append(r[:])
            continue
        r2 = r[:3] + [str(int(int(v) * 1.2)) for v in r[3:6]] + r[6:]
        if rng.random() < 0.15:
            r2[2] = str(rng.randint(1, 5))
        body.append(r2)
    ids = [k for k, r in enumerate(body) if r[1] != ""]
    gone = set(rng.sample(ids, 50))
    body = [r for k, r in enumerate(body) if k not in gone]
    if shape == "new_id_out_of_order":
        body.insert(500, ["", "2000000", "1", "1", "1", "1"] + ["0"] * 10)
    left = [H] + body
    right = [r[:] for r in base]
    tk = sorted(gone)[25]
    right[tk + 1][6] = "5"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    row = next(row for row in m if row[1][2] == right[tk + 1][1])
    assert row[6][0] == "conflict"


# ── 20차: 1..N 으로 매긴 진짜 ID 는 행 번호가 아니다 ───────────────────────────

def _monster_from(n, seed, start, step, head):
    rng = random.Random(seed)
    H = ["#", head, "Grade", "HP", "ATK", "DEF"] + [f"U{i}" for i in range(10)]
    return [H] + [["", str(start + step * i), str(rng.randint(1, 5))]
                  + [str(rng.randint(100, 9999)) for _ in range(3)] + ["0"] * 10 for i in range(n)]


@pytest.mark.parametrize("head", ["MonsterID", "Index", "ID"])
@pytest.mark.parametrize("start,step", [(1, 1), (0, 1), (100, 100)])
@pytest.mark.parametrize("edit", ["del50", "none"])
@pytest.mark.parametrize("seed", range(2))
def test_ids_numbered_from_one_are_not_row_numbers(head, start, step, edit, seed):
    """1..N(0..N, 100,200,…) 으로 매긴 진짜 ID 를 시트 전체 행 번호로 봐 ID 짝이 꺼지고, 밸런스 패치에서 이웃
    몬스터와 엇갈렸다."""
    rng = random.Random(100 + seed)
    base = _monster_from(1000, seed, start, step, head)
    body = []
    for r in base[1:]:
        r2 = r[:3] + [str(int(int(v) * 1.2)) for v in r[3:6]] + r[6:]
        if rng.random() < 0.15:
            r2[2] = str(rng.randint(1, 5))
        body.append(r2)
    gone = set(rng.sample(range(len(body)), 50)) if edit == "del50" else set()
    left = [base[0]] + [r for k, r in enumerate(body) if k not in gone]
    right = [r[:] for r in base]
    edited = []
    for k in sorted(gone)[::10]:
        right[k + 1][6] = "5"
        edited.append(k)
    alive = [k for k in range(0, 1000, 97) if k not in gone]
    for k in alive:
        right[k + 1][7] = "7"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    for k in edited:
        row = next(row for row in m if row[1][2] == right[k + 1][1])
        assert row[6][0] == "conflict"
    for k in alive:
        row = next(row for row in m if row[1][2] == right[k + 1][1])
        assert row[7][0] == "right" and row[1][0] == "same"


@pytest.mark.parametrize("edit", ["ins", "del"])
def test_renumbered_row_number_with_blank_section_rows(edit):
    """구간 표시 행의 No 가 빈칸이면 행 번호로 안 보여, 다시 매긴 No 로 짝지어 아래 묶음이 엇갈렸다."""
    rng = random.Random(8)
    H = ["GroupID", "No", "RewardType", "RewardID", "Count", "Weight"]
    groups = []
    for g in range(6):
        grp = []
        for _ in range(6):
            if rng.random() < 0.3:
                grp.append([str(5001 + g), "Gold", "1", str(rng.randint(100, 999)), "10"])
            else:
                grp.append([str(5001 + g), "Item", str(rng.randint(10000, 99999)), "1", "10"])
        groups.append(grp)
    body = [[r[:] for r in grp] for grp in groups]
    if edit == "ins":
        body[0].insert(2, ["5001", "Item", "70001", "1", "10"])
    else:
        del body[0][2]

    def sheet(gs):
        out, n = [], 0
        for g, grp in enumerate(gs):
            out.append(["", "", f"// group {g}", "", "", ""])
            for r in grp:
                n += 1
                out.append(r[:1] + [str(n)] + r[1:])
        return out
    base = [H] + sheet(groups)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    ti = next(i for i in range(len(right)) if right[i][0] == "5004" and right[i][2] == "Item")
    right[ti][4] = "999"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    row = next(row for row in m if row[3][2] == right[ti][3])
    assert row[3][0] == "same" and row[4][0] == "right"


# ── 21차: 순서만 바꾸고 자리대로 다시 매긴 행 번호(목록은 그대로) ─────────────────

@pytest.mark.parametrize("head", ["No", "Index", "Num", "#"])
@pytest.mark.parametrize("edit", ["swap", "sort", "rev", "move"])
@pytest.mark.parametrize("seed", range(4))
def test_reordered_rows_with_a_row_number_renumbered_in_place(head, edit, seed):
    """묶음 안 행을 맞바꾸거나 정렬하고 시트 전체 No 를 자리대로 다시 매기면(목록은 1..N 그대로) No 로 자리끼리
    짝지어 yours 수정이 다른 보상에 저장됐다."""
    rng = random.Random(seed)
    H = ["GroupID", head, "RewardType", "RewardID", "Count", "Weight"]
    groups = []
    for g in range(6):
        groups.append([["Item", str(rng.randint(10000, 99999)), "1", str(rng.randint(1, 99))]
                       for _ in range(6)])
    body = [[r[:] for r in grp] for grp in groups]
    tg = body[2]
    if edit == "swap":
        tg[1], tg[4] = tg[4], tg[1]
    elif edit == "sort":
        tg.sort(key=lambda r: r[1])
    elif edit == "rev":
        tg.reverse()
    else:
        tg.insert(4, tg.pop(1))

    def sheet(gs):
        out, n = [], 0
        for g, grp in enumerate(gs):
            for r in grp:
                n += 1
                out.append([str(5001 + g), str(n)] + r)
        return out
    base = [H] + sheet(groups)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    targets = [groups[2][1][1], groups[2][4][1]]
    for i, r in enumerate(right):
        if r[3] in targets:
            right[i][4] = "777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 3)
    for rid in targets:
        row = next(row for row in m if row[3][2] == rid)
        assert row[3][0] == "same" and row[4][0] == "right" and row[4][3] == "777"


# ── 22차: 다시 매긴 번호가 있는 정렬 — 자리 옮김을 넓게 감지, 빈 키 구간 정렬 ─────────────

@pytest.mark.parametrize("variant", ["seq_too", "sort10", "weight_too", "reweight_sorted"])
@pytest.mark.parametrize("head", ["No", "Index", "Num"])
@pytest.mark.parametrize("seed", range(4))
def test_reordered_rows_with_renumbered_numbers_and_other_changes(variant, head, seed):
    """행을 맞바꾸거나 정렬하면서 다른 자리 번호(Seq·Sort)도 다시 매기거나 Weight 를 함께 고치면, '똑같은 행'이
    없어 다시 매긴 No 를 ID 로 써 자리끼리 짝지었다."""
    rng = random.Random(seed)
    H = ["GroupID", head, "Seq", "RewardType", "RewardID", "Count", "Weight"]
    groups = [[["Item", str(rng.randint(10000, 99999)), "1", str(rng.randint(1, 99))] for _ in range(6)]
              for _ in range(6)]
    body = [[r[:] for r in grp] for grp in groups]
    tg = body[2]
    if variant == "reweight_sorted":
        for r in tg:
            r[3] = str(rng.randint(100, 999))
        tg.sort(key=lambda r: int(r[3]))
    else:
        tg[1], tg[4] = tg[4], tg[1]
        if variant == "weight_too":
            tg[1][3] = str(int(tg[1][3]) + 100)
            tg[4][3] = str(int(tg[4][3]) + 100)

    def sheet(gs):
        out, n = [], 0
        for g, grp in enumerate(gs):
            for s_, r in enumerate(grp):
                n += 1
                seq = str(10 * n) if variant == "sort10" else str(s_ + 1)
                out.append([str(5001 + g), str(n), seq] + r)
        return out
    base = [H] + sheet(groups)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    targets = [groups[2][1][1], groups[2][4][1]]
    for i, r in enumerate(right):
        if r[4] in targets:
            right[i][5] = "777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 4)
    for rid in targets:
        row = next(row for row in m if row[4][2] == rid)
        assert row[4][0] == "same" and row[5][0] == "right" and row[5][3] == "777"


@pytest.mark.parametrize("head", ["No", "Seq", "Index"])
@pytest.mark.parametrize("size", [12, 30])
@pytest.mark.parametrize("seed", range(3))
def test_sorting_blank_key_sections_with_renumbered_numbers_keeps_pairs(head, size, seed):
    """빈 키 구간을 등급순으로 정렬하고 시트 전체 번호를 자리대로 다시 매기면, 옮긴 행이 '지움 + 새 행'으로 갈려
    저장 때 중복되거나 다른 구간으로 갔다."""
    rng = random.Random(seed)
    H = ["#", head, "Name", "Grade", "HP", "ATK"]
    secs = [[[f"Mob{s}_{i}", str(rng.randint(1, 5)), str(rng.randint(100, 999)), str(rng.randint(10, 99))]
             for i in range(size)] for s in range(5)]
    body = [[r[:] for r in sec] for sec in secs]
    for s in (1, 2):
        body[s].sort(key=lambda r: (r[1], r[0]))

    def sheet(ss):
        out, n = [], 0
        for s, sec in enumerate(ss):
            out.append([f"//sec{s}", "", "", "", "", ""])
            for r in sec:
                n += 1
                out.append(["", str(n)] + r)
        return out
    base = [H] + sheet(secs)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    names = [secs[1][3][0], secs[2][5][0], secs[3][1][0]]
    for i, r in enumerate(right):
        if r[2] in names:
            right[i][4] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    assert all(not (mt[0] is None and mt[1] is not None) for mt in meta), "theirs 쪽에 짝 잃은 행"
    assert _conf(m) == 0
    for nm in names:
        row = next(row for row in m if row[2][2] == nm)
        assert row[4][0] == "right" and row[4][3] == "7777"


# ── 23차: 끝에 붙인 1..N 진짜 ID, 다음 ID 로 복사, 고유 키 큰 시트 속도 ─────────────

@pytest.mark.parametrize("head", ["MonsterID", "ID", "Index"])
@pytest.mark.parametrize("start,step", [(1, 1), (0, 1), (100, 100)])
@pytest.mark.parametrize("edit", ["app1", "app3", "cut30"])
@pytest.mark.parametrize("seed", range(2))
def test_ids_numbered_from_one_with_rows_added_or_cut_at_the_end(head, start, step, edit, seed):
    """1..N 진짜 ID 표에서 밸런스 패치와 함께 끝에 다음 ID 를 붙이거나 끝을 자르면, '목록이 달라짐'을 다시 매김
    으로 봐 ID 가 정렬 점수에서 빠지고 이웃 몬스터끼리 엇갈렸다."""
    rng = random.Random(200 + seed)
    n = 1000
    base = _monster_from(n, seed, start, step, head)
    body = []
    for r in base[1:]:
        r2 = r[:3] + [str(int(int(v) * 1.2)) for v in r[3:6]] + r[6:]
        if rng.random() < 0.15:
            r2[2] = str(rng.randint(1, 5))
        body.append(r2)
    if edit.startswith("app"):
        k = int(edit[3:])
        body += [["", str(start + step * (n + d)), "1", "1", "1", "1"] + ["0"] * 10 for d in range(k)]
    else:
        body = body[:-30]
    left = [base[0]] + body
    right = [r[:] for r in base]
    alive = [k for k in range(0, 960, 37)]
    for k in alive:
        right[k + 1][7] = "7"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    for k in alive:
        row = next(row for row in m if row[1][2] == right[k + 1][1])
        assert row[7][0] == "right" and row[1][0] == "same"


@pytest.mark.parametrize("head", ["MonsterID", "ID"])
@pytest.mark.parametrize("start,step", [(1, 1), (0, 1), (100, 100)])
@pytest.mark.parametrize("n,at", [(40, 39), (40, 10), (200, 150)])
def test_copying_to_the_next_id_with_ids_from_one_keeps_the_original_pair(head, start, step, n, at):
    """1..N 진짜 ID 표에서 다음 ID 로 끝에 복사하고 원래 행을 고치면, 원래 행이 복사본과 짝지어졌다."""
    rng = random.Random(at)
    H = ["#", head, "Grade", "HP", "ATK", "DEF", "SkillID"]
    base = [H] + [["", str(start + step * i), str(rng.randint(1, 5))]
                  + [str(rng.randint(100, 9999)) for _ in range(3)] + ["0"] for i in range(n)]
    body = [r[:] for r in base[1:]]
    copy = body[at][:]
    copy[1] = str(start + step * n)
    body[at] = body[at][:3] + [str(int(int(v) * 1.2)) for v in body[at][3:6]] + body[at][6:]
    body.append(copy)
    left = [H] + body
    right = [r[:] for r in base]
    right[at + 1][6] = "5001"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    row = next(row for row in m if row[1][2] == right[at + 1][1])
    assert row[1][0] == "same" and row[6][0] == "right" and row[6][3] == "5001"


def test_unique_key_sheet_with_a_row_number_and_many_id_columns_stays_fast():
    """키가 모두 고유한 큰 시트(5만 행, No + ID 이름 열 10개)에서 내용 짝 준비 계산이 비교를 몇 배 느리게 했다."""
    rng = random.Random(1)
    H = ["Key", "No"] + [f"Ref{i}ID" for i in range(10)] + ["HP"]
    base = [H] + [[f"K{i}", str(i + 1)] + [str(rng.randint(1, 10 ** 6)) for _ in range(10)]
                  + [str(rng.randint(1, 999))] for i in range(50000)]
    left = [r[:] for r in base]
    for r in left[1::3]:
        r[-1] = str(int(r[-1]) + 1)
    right = [r[:] for r in base]
    right[40001][-1] = "5"                     # 왼쪽이 고치지 않은 행(왼쪽은 1, 4, 7 … 행을 고침)
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert _within(time.perf_counter() - t0, 4.0, (left, base, right))
    assert _conf(m) == 0


# ── 24차: 강한 이름의 시트 전체 순번, 참조 ID 맞바꿈 ───────────────────────────

@pytest.mark.parametrize("head", ["SortKey", "RowID", "SlotID", "SeqID", "OrderID"])
@pytest.mark.parametrize("edit", ["ins", "del", "ins3"])
@pytest.mark.parametrize("n", [60, 300])
def test_strong_named_sheet_wide_order_renumbered_is_not_an_id(head, edit, n):
    """SortKey·RowID 처럼 이름이 ID 인 시트 전체 1..N 순번을 끼움·지움 뒤 다시 매기면 번호끼리(자리끼리) 짝지었다."""
    rng = random.Random(n)
    H = ["#", head, "Name", "Grade", "HP", "ATK"]
    rows = [[f"Mob{i}", str(rng.randint(1, 5)), str(rng.randint(100, 999)), str(rng.randint(1000, 4999))]
            for i in range(n)]
    body = [r[:] for r in rows]
    at = n // 3
    if edit == "ins":
        body.insert(at, ["MobNew0", "1", "1", "1"])
    elif edit == "ins3":
        body[at:at] = [[f"MobNew{d}", "1", "1", "1"] for d in range(3)]
    else:
        del body[at]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    targets = [rows[i][0] for i in range(at + 1, min(n, at + 12), 2)]
    for i, r in enumerate(right):
        if r[2] in targets:
            right[i][5] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    for nm in targets:
        row = next(row for row in m if row[2][2] == nm)
        assert row[5][0] == "right" and row[5][3] == "7777"


@pytest.mark.parametrize("ref", ["DropID", "RewardID", "SkillID"])
@pytest.mark.parametrize("n", [30, 200])
@pytest.mark.parametrize("patch", [True, False])
def test_swapping_reference_ids_between_two_monsters_is_not_a_row_move(ref, n, patch):
    """두 몬스터의 DropID 만 맞바꾸면 MonsterID 를 다시 매긴 번호로 봐, DropID 로 두 몬스터를 엇갈려 짝지었다."""
    rng = random.Random(n)
    H = ["#", "MonsterID", "Name", "Grade", "HP", "ATK", ref, "Memo"]
    rows = [["", str(i + 1), f"Mob{i}", str(rng.randint(1, 5)), str(rng.randint(100, 9999)),
             str(rng.randint(100, 9999)), str(500000 + i * 7), ""] for i in range(n)]
    body = [r[:] for r in rows]
    if patch:
        for r in body:
            r[4] = str(int(int(r[4]) * 1.2))
    a, b = n // 5, n - n // 4
    body[a][6], body[b][6] = body[b][6], body[a][6]
    base = [H] + rows
    left = [H] + body
    right = [r[:] for r in base]
    right[b + 1][7] = "tuned"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    row = next(row for row in m if row[1][2] == right[b + 1][1])
    assert row[1][0] == "same" and row[7][0] == "right" and row[6][0] == "left"


# ── 25차: 밸런스 패치와 함께 다시 매긴 강한 이름 순번, 자리 재활용 ───────────────

@pytest.mark.parametrize("head", ["SortKey", "RowID", "OrderID"])
@pytest.mark.parametrize("edit", ["ins", "del", "ins3", "del3"])
@pytest.mark.parametrize("n", [60, 300])
def test_strong_named_order_renumbered_with_a_bulk_patch(head, edit, n):
    """모든 행 HP ×1.2 와 함께 끼움·지움 뒤 순번을 다시 매기면 '똑같은 행'이 없어 순번으로 자리끼리 짝지었다."""
    rng = random.Random(n + len(edit))
    H = ["#", head, "Name", "Grade", "HP", "ATK"]
    rows = [[f"Mob{i}", str(rng.randint(1, 5)), str(rng.randint(100, 999)), str(rng.randint(1000, 4999))]
            for i in range(n)]
    body = [r[:2] + [str(int(int(r[2]) * 1.2))] + r[3:] for r in rows]
    at = n // 3
    if edit == "ins":
        body.insert(at, ["MobNew0", "1", "1", "1"])
    elif edit == "ins3":
        body[at:at] = [[f"MobNew{d}", "1", "1", "1"] for d in range(3)]
    elif edit == "del":
        del body[at]
    else:
        del body[at:at + 3]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    targets = [rows[i][0] for i in range(at + 4, min(n, at + 20), 3)]
    for i, r in enumerate(right):
        if r[2] in targets:
            right[i][5] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    for nm in targets:
        row = next(row for row in m if row[2][2] == nm)
        assert row[5][0] == "right" and row[5][3] == "7777"


@pytest.mark.parametrize("head", ["SeqID", "SortKey"])
@pytest.mark.parametrize("edit", ["ins", "del"])
def test_strong_named_order_renumbered_with_a_column_changed_everywhere(head, edit):
    """모든 행의 Count 를 바꾸고 행 하나를 끼우거나 지운 뒤 순번을 다시 매기면 순번으로 자리끼리 짝지었다."""
    rng = random.Random(4)
    H = ["GroupID", head, "ItemName", "Count", "Weight"]
    names = ["Key", "Dust", "Shard", "Elixir", "Gold", "Gem", "Scroll"]
    groups = []
    for g in range(40):
        groups.append([[rng.choice(names), str(rng.randint(1, 99)), str(rng.randint(10, 99))]
                       for _ in range(rng.randint(3, 6))])
    body = [[[r[0], str(int(r[1]) * 2), r[2]] for r in grp] for grp in groups]
    if edit == "ins":
        body[3].insert(1, ["NewItem", "2", "10"])
    else:
        del body[3][1]

    def sheet(gs):
        out, n = [], 0
        for g, grp in enumerate(gs):
            for r in grp:
                n += 1
                out.append([str(5001 + g), str(n)] + r)
        return out
    base = [H] + sheet(groups)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    for i, r in enumerate(right[1:], start=1):
        if int(r[0]) >= 5011 and i % 2 == 0:
            right[i][4] = "999"
    m, meta = compute_diff3(left, base, right, 0, 0)
    bad = [row for row, mt in zip(m, meta) if row[4][3] == "999" and row[4][0] == "right"
           and mt[0] is not None and mt[1] is not None and row[2][1] != row[2][2]]
    assert not bad, bad[:3]


@pytest.mark.parametrize("start", [1, 0])
def test_pasting_one_monster_over_another_existing_id_keeps_the_original_pair(start):
    """몬스터 A 의 내용을 이미 있는 몬스터 B 자리에 붙여 넣고 A 를 버프하면, A 가 B 번호로 옮겨 간 것으로 봐
    MonsterID 를 버렸다."""
    rng = random.Random(7)
    H = ["#", "MonsterID", "Grade", "HP", "ATK", "DEF", "SkillID"]
    n = 200
    base = [H] + [["", str(start + i), str(rng.randint(1, 5))]
                  + [str(rng.randint(100, 9999)) for _ in range(3)] + ["0"] for i in range(n)]
    body = [r[:] for r in base[1:]]
    a, b = 50, 150
    body[b] = [body[b][0], body[b][1]] + body[a][2:]
    body[a] = body[a][:3] + [str(int(int(v) * 1.2)) for v in body[a][3:6]] + body[a][6:]
    left = [H] + body
    right = [r[:] for r in base]
    right[a + 1][6] = "5001"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    row = next(row for row in m if row[1][2] == right[a + 1][1])
    assert row[1][0] == "same" and row[6][0] == "right" and row[6][3] == "5001"


# ── 26차: 끝 근처 끼움·지움, 묶음 교체, 복사와 이동 가르기, 큰 시트 속도 ─────────────

@pytest.mark.parametrize("head", ["SortKey", "OrderID", "RowID"])
@pytest.mark.parametrize("edit,pos", [("ins", 1), ("ins", 2), ("ins", 3), ("ins3", 4), ("del", 3), ("del3", 4),
                                      ("del", 2)])
@pytest.mark.parametrize("patch", [False, True])
@pytest.mark.parametrize("seed", range(3))
def test_strong_named_order_edited_near_the_end(head, edit, pos, patch, seed):
    """끝 가까이에 끼우거나 지우고 순번을 다시 매기면 옮긴 행이 0~1개라 다시 매김을 놓치고 자리끼리 짝지었다."""
    rng = random.Random(seed)
    n = 60
    H = ["#", head, "Name", "Grade", "HP", "ATK"]
    rows = [[f"Mob{i}", str(rng.randint(1, 5)), str(rng.randint(100, 999)), str(rng.randint(1000, 4999))]
            for i in range(n)]
    body = [r[:2] + ([str(int(int(r[2]) * 1.2))] if patch else [r[2]]) + r[3:] for r in rows]
    at = n - pos
    if edit == "ins":
        body.insert(at, ["New0", "1", "1", "1"])
    elif edit == "ins3":
        body[at:at] = [[f"New{d}", "1", "1", "1"] for d in range(3)]
    elif edit == "del":
        del body[at]
    else:
        del body[at:at + 3]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    gone = {r[0] for r in rows} - {r[0] for r in body}
    targets = [rows[i][0] for i in range(n - 6, n) if rows[i][0] not in gone]
    for i, r in enumerate(right):
        if r[2] in targets:
            right[i][5] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    for nm in targets:
        row = next(row for row in m if row[2][2] == nm)
        assert row[5][0] == "right" and row[5][3] == "7777"


@pytest.mark.parametrize("head", ["SortID", "RowID"])
@pytest.mark.parametrize("edit", ["replace_block", "raise_some"])
def test_strong_named_order_with_a_replaced_block_or_partial_raise(head, edit):
    """가운데 몹 묶음을 새 몹으로 바꾸거나(10% 넘게) 일부 몹 스탯만 올리면서 순번을 다시 매기면, 일괄 수정 열을
    빼느라 옮김 근거가 사라져 순번끼리 짝지었다."""
    rng = random.Random(2)
    n = 100
    H = ["#", head, "Name", "Grade", "HP", "ATK", "DEF"]
    rows = [[f"Mob{i}", str(rng.randint(1, 5))] + [str(rng.randint(100, 9999)) for _ in range(3)]
            for i in range(n)]
    body = [r[:] for r in rows]
    if edit == "replace_block":
        body[40:60] = [[f"NewMob{d}", "3", "500", "500", "500"] for d in range(25)]
    else:
        for r in body:
            if rng.random() < 0.35:
                r[2:5] = [str(int(int(v) * 1.3)) for v in r[2:5]]
        body.insert(30, ["NewMob", "1", "1", "1", "1"])

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    gone = {r[0] for r in rows} - {r[0] for r in body}
    targets = [rows[i][0] for i in range(65, 95, 4) if rows[i][0] not in gone]
    for i, r in enumerate(right):
        if r[2] in targets:
            right[i][3] = "9"                  # theirs 가 올린 스탯과 겹치지 않는 Grade
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    for nm in targets:
        row = next(row for row in m if row[2][2] == nm)
        assert row[3][0] == "right" and row[3][3] == "9"


def test_two_pastes_over_existing_ids_keep_original_pairs():
    """이미 있는 두 ID 자리에 다른 몬스터 내용을 붙여 넣으면 '옮긴 행 2개'로 봐 MonsterID 를 버렸다."""
    rng = random.Random(9)
    H = ["#", "MonsterID", "Grade", "HP", "ATK", "DEF", "SkillID"]
    n = 200
    base = [H] + [["", str(i + 1), str(rng.randint(1, 5))] + [str(rng.randint(100, 9999)) for _ in range(3)]
                  + ["0"] for i in range(n)]
    body = [r[:] for r in base[1:]]
    for a, b in ((40, 150), (60, 170)):
        body[b] = [body[b][0], body[b][1]] + body[a][2:]
        body[a] = body[a][:3] + [str(int(int(v) * 1.2)) for v in body[a][3:6]] + body[a][6:]
    left = [H] + body
    right = [r[:] for r in base]
    right[41][6] = "5001"
    right[61][6] = "5002"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 1)
    for i, val in ((41, "5001"), (61, "5002")):
        row = next(row for row in m if row[1][2] == right[i][1])
        assert row[1][0] == "same" and row[6][0] == "right" and row[6][3] == val


def test_duplicate_key_sheet_with_rows_appended_at_the_end_stays_fast():
    """겹치는 키 큰 시트(5만 행, No+Sort, ID 이름 열 20개)에서 한쪽 끝에 몇 행만 붙여도 다시 매김 검사가 시트
    전체를 다시 훑어 비교가 몇 초 늘었다."""
    rng = random.Random(3)
    H = ["Group", "No", "Sort"] + [f"Ref{i}ID" for i in range(20)] + ["Count"]
    rows = []
    for i in range(50000):
        rows.append([str(i // 10), str(i + 1), str(10 * (i + 1))]
                    + [str(rng.randint(1, 10 ** 7)) for _ in range(20)] + ["1"])
    base = [H] + rows
    left = [H] + [r[:] for r in rows] + [[str(5000), str(50001 + d), str(10 * (50001 + d))]
                                          + [str(9 * 10 ** 7 + d)] * 20 + ["1"] for d in range(3)]
    right = [r[:] for r in base]
    right[123][-1] = "5"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert _within(time.perf_counter() - t0, 6.0, (left, base, right))
    assert _conf(m) == 0


# ── 27차: 복사본을 위에 끼우고 이름만 바꿈, 같은 수만큼 지우고 넣음 ─────────────────

@pytest.mark.parametrize("extra", [["Icon"], ["Icon", "Prefab"], ["Desc"]])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_copy_inserted_above_and_renamed_is_a_move(extra, shape):
    """마지막 몬스터를 복사해 바로 위에 끼우고 이름만 바꾸면, Icon 같은 다른 고유 글자 열이 같아 복사로 봐
    다시 매김을 놓쳤다. ('마지막 1행' 모양은 '끝에 백업 복사 + 원본 개편'과 파일이 같아 백업 쪽으로 정했다 — 34차.)"""
    rng = random.Random(5)
    n = 60
    H = ["#", "SortKey", "Name"] + extra + ["Grade", "HP", "ATK"]
    rows = [[f"Mob{i}"] + [f"{e}_{i:03d}" for e in extra]
            + [str(rng.randint(1, 5)), str(rng.randint(100, 999)), str(rng.randint(1000, 4999))] for i in range(n)]
    body = [r[:] for r in rows]
    ne = len(extra)

    def variant(r):
        v = r[:]
        v[0] = r[0] + "_Hard"
        v[1 + ne] = "5"
        v[2 + ne] = str(int(r[2 + ne]) * 2)
        return v
    if shape == "last1":
        body.insert(n - 1, variant(rows[n - 1]))
        targets = [rows[n - 1][0]]
    elif shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        targets = [r[0] for r in rows[n - 3:]]
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        targets = [rows[30][0]]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    atk = len(H) - 1
    for i, r in enumerate(right):
        if r[2] in targets:
            right[i][atk] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    for nm in targets:
        row = next(row for row in m if row[2][2] == nm)
        assert row[atk][0] == "right" and row[atk][3] == "7777"


@pytest.mark.parametrize("head", ["SortID", "RowID"])
@pytest.mark.parametrize("shape", ["mid_del_insert_later", "front_del_end_add"])
def test_deleting_and_adding_the_same_number_of_rows_with_renumbering(head, shape):
    """몹 20행을 지우고 같은 수를 다른 곳에 넣어 행 수가 같으면, 이름 열까지 일괄 수정으로 빠져 옮김 근거가
    사라지고 순번끼리 짝지었다."""
    rng = random.Random(6)
    n = 150
    H = ["#", head, "Name", "HP", "ATK", "U0"]
    rows = [[f"Mob{i}", str(rng.randint(100, 9999)), str(rng.randint(100, 9999)), "0"] for i in range(n)]
    new = [[f"New{d}", str(rng.randint(100, 9999)), str(rng.randint(100, 9999)), "0"] for d in range(20)]
    if shape == "mid_del_insert_later":
        body = rows[:30] + rows[50:100] + new + rows[100:]
    else:
        body = rows[20:] + new

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    targets = ["Mob55", "Mob61", "Mob90"]
    for i, r in enumerate(right):
        if r[2] in targets:
            right[i][5] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    for nm in targets:
        row = next(row for row in m if row[2][2] == nm)
        assert row[5][0] == "right" and row[5][3] == "7777"


# ── 28차: 끝에 백업 복사 + 원본 개편, 묶음 교체 + 패치 + 다시 매김 ────────────────

@pytest.mark.parametrize("head", ["MonsterID", "SortKey"])
@pytest.mark.parametrize("extra", [["Desc"], ["Icon", "Desc"]])
@pytest.mark.parametrize("n,at", [(30, 10), (60, 20), (200, 67)])
@pytest.mark.parametrize("edit", ["desc_hp", "name_grade"])
def test_backup_copy_at_the_end_and_reworked_original_keep_pairs(head, extra, n, at, edit):
    """옛 버전을 끝에 다음 번호로 그대로 복사해 두고 원래 몬스터의 Desc(또는 Name)와 스탯을 고치면, 원래 몬스터를
    복사본과 짝지어 yours 수정이 복사본에 저장됐다."""
    rng = random.Random(n + at)
    H = ["#", head, "Name"] + extra + ["Grade", "HP", "ATK", "DEF"]
    rows = [[f"Mob{i}"] + [f"{e}_{i:03d}" for e in extra]
            + [str(rng.randint(1, 5))] + [str(rng.randint(100, 9999)) for _ in range(3)] for i in range(n)]
    body = [r[:] for r in rows]
    copy = rows[at][:]
    ne = len(extra)
    if edit == "desc_hp":
        body[at][ne] = body[at][ne] + " (개편)"
        body[at][2 + ne] = "1500"
    else:
        body[at][0] = body[at][0] + "_Old"
        body[at][1 + ne] = "5"
    body.append(copy)

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    right[at + 1][len(H) - 1] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    row = next(row for row, mt in zip(m, meta) if mt[1] == at + 1)
    assert row[1] == ("same", str(at + 1), str(at + 1), str(at + 1))
    assert row[len(H) - 1][0] == "right" and row[len(H) - 1][3] == "7777"


@pytest.mark.parametrize("head", ["MonsterID", "ID"])
@pytest.mark.parametrize("n", [30, 200])
def test_last_monster_copied_to_the_next_id_with_its_desc_edited(head, n):
    """마지막 몬스터를 다음 ID 로 그대로 복사하고 원래 행의 Desc 를 고치면 이동으로 봐, 다른 쪽이 원래 몬스터에
    한 수정이 복사본에 저장됐다."""
    rng = random.Random(n)
    H = ["#", head, "Name", "Icon", "Desc", "HP", "ATK"]
    rows = [[f"Mob{i}", f"Icon_{i:03d}", f"설명_{i:03d}", str(rng.randint(100, 9999)), str(rng.randint(100, 9999))]
            for i in range(n)]
    body = [r[:] for r in rows]
    body.append(rows[-1][:])
    body[n - 1][2] = body[n - 1][2] + " (개편)"

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    yours = [H] + sheet(body)
    theirs = [r[:] for r in base]
    theirs[n][5] = "7777"
    m, meta = compute_diff3(theirs, base, yours, 0, 0)
    row = next(row for row, mt in zip(m, meta) if mt[1] == n)
    assert row[1] == ("same", str(n), str(n), str(n))
    assert row[5][0] == "left" and row[5][1] == "7777" and _conf(m) == 0


@pytest.mark.parametrize("head", ["SortID", "RowID"])
@pytest.mark.parametrize("k", [20, 21, 25])
@pytest.mark.parametrize("where", ["mid", "front"])
def test_replacing_a_block_with_a_bulk_patch_and_renumbering(head, k, where):
    """몹 묶음을 새 몹으로 바꾸면서 모든 행 HP 패치와 순번 다시 매김을 함께 하면 두 투영 모두 근거가 없어 순번끼리
    짝지었다."""
    rng = random.Random(k)
    n = 150
    H = ["#", head, "Name", "HP", "ATK", "U0"]
    rows = [[f"Mob{i}", str(rng.randint(100, 9999)), str(rng.randint(100, 9999)), "0"] for i in range(n)]
    new = [[f"New{d}", str(rng.randint(100, 9999)), str(rng.randint(100, 9999)), "0"] for d in range(k)]
    if where == "mid":
        body = rows[:30] + rows[50:100] + new + rows[100:]
    else:
        body = rows[20:] + new
    body = [[r[0], str(int(int(r[1]) * 1.2))] + r[2:] for r in body]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    targets = ["Mob55", "Mob61", "Mob90"]
    for i, r in enumerate(right):
        if r[2] in targets:
            right[i][5] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    for nm in targets:
        row = next(row for row in m if row[2][2] == nm)
        assert row[5][0] == "right" and row[5][3] == "7777"


# ── 29차: 이름 열보다 앞의 글자 열, 끝 백업 복사가 많을 때 속도 ─────────────────

@pytest.mark.parametrize("before,name", [(["Icon"], "Name"), (["비고"], "이름"), (["ResPath"], "MonsterName"),
                                         (["Prefab", "Code"], "Name")])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_copy_inserted_above_and_renamed_with_text_columns_before_the_name(before, name, shape):
    """Icon·비고 같은 글자 열이 Name 보다 앞에 있으면 가장 왼쪽 열(Icon)이 같다고 복사로 봐 다시 매김을 놓쳤다."""
    rng = random.Random(8)
    n = 60
    nb = len(before)
    H = ["#", "SortKey"] + before + [name, "Grade", "HP", "ATK"]
    rows = [[f"{b}_{i:03d}" for b in before] + [f"Mob{i}", str(rng.randint(1, 5)), str(rng.randint(100, 999)),
                                                str(rng.randint(1000, 4999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[nb] = r[nb] + "_Hard"
        v[nb + 1] = "5"
        v[nb + 2] = str(int(r[nb + 2]) * 2)
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        targets = [r[nb] for r in rows[n - 3:]]
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        targets = [rows[30][nb]]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    ncol = 2 + nb
    atk = len(H) - 1
    for i, r in enumerate(right):
        if r[ncol] in targets:
            right[i][atk] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, ncol)
    for nm in targets:
        row = next(row for row in m if row[ncol][2] == nm)
        assert row[atk][0] == "right" and row[atk][3] == "7777"


@pytest.mark.parametrize("n,k", [(20000, 1000), (10000, 300)])
def test_many_backup_copies_at_the_end_stay_fast(n, k):
    """끝에 백업 복사본이 많고 원본만 고치면 복사본마다 시트 전체를 다시 훑어 수십 초 걸렸다."""
    rng = random.Random(n)
    H = ["#", "MonsterID", "Name", "Desc", "Grade", "HP", "ATK"]
    rows = [["", str(i + 1), f"Mob{i}", f"D_{i:05d}", str(rng.randint(1, 5)), str(rng.randint(100, 9999)),
             str(rng.randint(100, 9999))] for i in range(n)]
    body = [r[:] for r in rows]
    at = n // 20
    for d in range(k):
        cp = rows[at + d][:]
        cp[1] = str(n + 1 + d)
        body.append(cp)
        body[at + d][5] = str(int(body[at + d][5]) + 1)
    base = [H] + rows
    left = [H] + body
    right = [r[:] for r in base]
    right[at + 5][6] = "7777"
    t0 = time.perf_counter()
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert _within(time.perf_counter() - t0, 3.0, (left, base, right))
    row = next(row for row, mt in zip(m, meta) if mt[1] == at + 5)
    assert row[1][0] == "same" and row[6][0] == "right"


# ── 30차: 이름 계열 열 여럿, 겹치는 표시 이름, 패치 + 백업 복사 + 원본 개편 ─────────────

@pytest.mark.parametrize("before,name", [(["IconName"], "Name"), (["PrefabName"], "MonsterName"), (["NameKey"], "Name"),
                                         (["ModelName", "Icon"], "Name"), (["DevName"], "이름")])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_copy_inserted_above_and_renamed_with_name_like_resource_columns(before, name, shape):
    """IconName·PrefabName 같은 이름 계열 리소스 열이 Name 앞에 있으면 그 열 하나로 복사로 봐 다시 매김을 놓쳤다."""
    test_copy_inserted_above_and_renamed_with_text_columns_before_the_name(before, name, shape)


@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_copy_inserted_above_and_renamed_with_shared_display_names(shape):
    """표시 이름이 겹치는 표(Name 비고유 + Icon 고유)에서 복사본을 위에 끼우고 이름만 바꾸면 복사로 봤다."""
    rng = random.Random(4)
    n = 60
    H = ["#", "SortKey", "Name", "Icon", "Grade", "HP", "ATK"]
    rows = [[f"슬라임{i % 7}", f"Icon_{i:03d}", str(rng.randint(1, 5)), str(rng.randint(100, 999)),
             str(rng.randint(1000, 4999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[0] = r[0] + "_Hard"
        v[2] = "5"
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        idx = list(range(n - 3, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[i + 1][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    for i in idx:
        row = next(row for row in m if row[3][2] == rows[i][1])
        assert row[3][1] == rows[i][1] and row[2][1] == rows[i][0]
        assert row[6][0] == "right" and row[6][3] == "7777"


@pytest.mark.parametrize("cols", [["Name", "Icon", "Desc"], ["Icon", "Desc"], ["Desc"], ["Name_dup", "Desc"]])
def test_last_row_copied_to_the_next_id_with_its_text_edited_in_various_tables(cols):
    """Name 이 없거나 겹치는 표에서 마지막 몹을 다음 ID 로 복사하고 원본 Desc 를 고치면 이동으로 봐, 다른 쪽이
    원본에 한 수정이 복사본에 저장됐다."""
    rng = random.Random(len(cols))
    n = 30
    H = ["#", "MonsterID"] + [("Name" if c == "Name_dup" else c) for c in cols] + ["HP", "ATK"]

    def cell(c, i):
        return {"Name": f"Mob{i}", "Name_dup": f"슬라임{i % 5}", "Icon": f"Icon_{i:03d}", "Desc": f"설명_{i:03d}"}[c]
    rows = [[cell(c, i) for c in cols] + [str(rng.randint(100, 9999)), str(rng.randint(100, 9999))]
            for i in range(n)]
    body = [r[:] for r in rows]
    body.append(rows[-1][:])
    di = cols.index("Desc")
    body[n - 1][di] = body[n - 1][di] + " (개편)"

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    yours = [H] + sheet(body)
    theirs = [r[:] for r in base]
    hp = len(H) - 2
    theirs[n][hp] = "7777"
    m, meta = compute_diff3(theirs, base, yours, 0, 0)
    row = next(row for row, mt in zip(m, meta) if mt[1] == n)
    assert row[1] == ("same", str(n), str(n), str(n))
    assert row[hp][0] == "left" and row[hp][1] == "7777" and _conf(m) == 0


@pytest.mark.parametrize("cols,edit", [(["Name", "Icon"], "rename"), (["Icon", "Desc"], "icon"), (["Desc"], "desc")])
def test_bulk_patch_plus_backup_copy_plus_reworked_original(cols, edit):
    """모든 행 HP ×1.2 + 패치 전 행을 끝에 백업 복사 + 원본 개편을 함께 하면, 원본이 백업과 짝지어졌다."""
    rng = random.Random(3)
    n = 200
    H = ["Grp", "MonsterID"] + cols + ["Grade", "HP", "ATK", "DEF"]
    nc = len(cols)

    def cell(c, i):
        return {"Name": f"Mob{i}", "Icon": f"Icon_{i:03d}", "Desc": f"설명_{i:03d}"}[c]
    rows = [["", str(i + 1)] + [cell(c, i) for c in cols] + [str(rng.randint(1, 5))]
            + [str(rng.randint(100, 9999)) for _ in range(3)] for i in range(n)]
    at = 10
    body = [r[:2 + nc + 1] + [str(int(int(r[3 + nc]) * 1.2))] + r[4 + nc:] for r in rows]
    backup = rows[at][:]
    backup[1] = str(n + 1)
    body[at][2] = body[at][2] + ("_Rework" if edit == "rename" else "_v2")
    body.append(backup)
    base = [H] + rows
    left = [H] + body
    right = [r[:] for r in base]
    atk = 4 + nc
    right[at + 1][atk] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    row = next(row for row, mt in zip(m, meta) if mt[1] == at + 1)
    assert row[1] == ("same", str(at + 1), str(at + 1), str(at + 1))
    assert row[atk][0] == "right" and row[atk][3] == "7777"


# ── 31차: 예약 ID 행에 백업, 리소스 이름 열만 고침, 범주형 이름 열 ────────────────────

@pytest.mark.parametrize("cols,edit", [(["Name", "Icon"], "Name"), (["Name", "Desc"], "Name"),
                                       (["Name", "IconName"], "IconName")])
@pytest.mark.parametrize("filler", ["blank", "dummy"])
@pytest.mark.parametrize("n", [30, 200])
def test_backup_pasted_into_a_reserved_id_row_with_a_bulk_patch(cols, edit, filler, n):
    """패치 전 백업을 이미 있는 예약 ID 행(빈칸·Dummy)에 붙여 넣고 원본을 개편하면, 원본이 백업과 짝지어졌다."""
    rng = random.Random(n)
    H = ["Grp", "MonsterID"] + cols + ["Grade", "HP", "ATK", "DEF"]
    nc = len(cols)

    def cell(c, i):
        return {"Name": f"Mob{i}", "Icon": f"Icon_{i:03d}", "Desc": f"설명_{i:03d}", "IconName": f"ico_{i:03d}"}[c]
    rows = [["", str(i + 1)] + [cell(c, i) for c in cols] + [str(rng.randint(1, 5))]
            + [str(rng.randint(100, 9999)) for _ in range(3)] for i in range(n)]
    for d in range(5):
        rows.append(["", str(n + 1 + d)] + ([""] * (nc + 4) if filler == "blank" else ["Dummy"] * nc + ["0"] * 4))
    at = 10
    body = [r[:2 + nc + 1] + ([str(int(int(r[3 + nc]) * 1.2))] if r[3 + nc] not in ("", "0") else [r[3 + nc]])
            + r[4 + nc:] for r in rows]
    backup = rows[at][:]
    backup[1] = str(n + 1)
    body[n] = backup
    ei = 2 + cols.index(edit)
    body[at][ei] = body[at][ei] + "_Rework"
    base = [H] + rows
    left = [H] + body
    right = [r[:] for r in base]
    deff = len(H) - 1
    right[at + 1][deff] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    row = next(row for row, mt in zip(m, meta) if mt[1] == at + 1)
    assert row[1] == ("same", str(at + 1), str(at + 1), str(at + 1))
    assert row[deff][0] == "right" and row[deff][3] == "7777"


@pytest.mark.parametrize("cols,edit", [(["Name", "IconName", "Desc"], "IconName"), (["Name", "SkillName"], "SkillName"),
                                       (["Name", "EffectName", "Desc"], "EffectName")])
@pytest.mark.parametrize("n", [30, 200])
def test_last_row_backup_with_a_resource_name_column_edited(cols, edit, n):
    """마지막 몹을 다음 ID 로 백업 복사하고 원본의 IconName·SkillName 만 고치면 이동으로 봐, 다른 쪽이 원본에 한
    수정이 백업에 저장됐다."""
    rng = random.Random(n + len(cols))
    H = ["#", "MonsterID"] + cols + ["HP", "ATK"]
    rows = [[f"{c}_{i:03d}" if c != "Name" else f"Mob{i}" for c in cols]
            + [str(rng.randint(100, 9999)), str(rng.randint(100, 9999))] for i in range(n)]
    body = [r[:] for r in rows]
    body.append(rows[-1][:])
    ei = cols.index(edit)
    body[n - 1][ei] = body[n - 1][ei] + "_new"

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    yours = [H] + sheet(body)
    theirs = [r[:] for r in base]
    hp = len(H) - 2
    theirs[n][hp] = "7777"
    m, meta = compute_diff3(theirs, base, yours, 0, 0)
    row = next(row for row, mt in zip(m, meta) if mt[1] == n)
    assert row[1] == ("same", str(n), str(n), str(n))
    assert row[hp][0] == "left" and row[hp][1] == "7777" and _conf(m) == 0


@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_copy_inserted_above_with_only_a_categorical_name_column(shape):
    """표시 이름 없이 ElementName(반복 값)만 있는 표에서 복사본을 위에 끼우고 Icon 을 바꾸면 복사로 봐 다시
    매김을 놓쳤다."""
    rng = random.Random(6)
    n = 60
    H = ["#", "SortKey", "Icon", "ElementName", "Grade", "HP", "ATK"]
    elems = ["Fire", "Water", "Wind", "Light", "Dark"]
    rows = [[f"Icon_{i:03d}", elems[i % 5], str(rng.randint(1, 5)), str(rng.randint(100, 999)),
             str(rng.randint(1000, 4999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[0] = r[0] + "_Hard"
        v[2] = "5"
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        idx = list(range(n - 3, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[i + 1][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not _wrong_pairs(m, meta, 2)
    for i in idx:
        row = next(row for row in m if row[2][2] == rows[i][0])
        assert row[6][0] == "right" and row[6][3] == "7777"


# ── 32차: 머리글 표기에 따른 표시 이름(스킬·클래스 표, 한글, 2단 머리글) ──────────────

@pytest.mark.parametrize("heads,edit", [
    ([["#", "SkillID", "SkillName", "IconName", "Desc", "HP", "ATK"]], 3),
    ([["#", "ClassID", "ClassName", "EffectName", "Desc", "HP", "ATK"]], 3),
    ([["#", "몬스터ID", "이름", "아이콘 이름", "설명", "체력", "공격력"]], 3),
    ([["", "몬스터 번호", "이름", "아이콘 이름", "설명", "체력", "공격력"],
      ["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"]], 3),
    ([["", "스킬 번호", "스킬 이름", "아이콘 이름", "설명", "체력", "공격력"],
      ["#", "SkillID", "SkillName", "IconName", "Desc", "HP", "ATK"]], 3),
])
@pytest.mark.parametrize("n", [30, 200])
def test_last_row_backup_with_resource_name_edited_under_various_headers(heads, edit, n):
    """스킬·클래스 표의 표시 이름(SkillName), 한글 '아이콘 이름', 2단 머리글에서 마지막 행을 다음 ID 로 백업하고
    원본 아이콘 이름을 고치면 이동으로 봐, 다른 쪽이 원본에 한 수정이 백업에 저장됐다."""
    rng = random.Random(n)
    rows = [["", str(i + 1), f"Nm{i}", f"ico_{i:03d}", f"설명_{i:03d}", str(rng.randint(100, 9999)),
             str(rng.randint(100, 9999))] for i in range(n)]
    body = [r[:] for r in rows]
    cp = rows[-1][:]
    cp[1] = str(n + 1)
    body.append(cp)
    body[n - 1][edit] = body[n - 1][edit] + "_v2"
    kr = len(heads) - 1
    base = [h[:] for h in heads] + rows
    yours = [h[:] for h in heads] + body
    theirs = [r[:] for r in base]
    theirs[kr + n][5] = "7777"
    m, meta = compute_diff3(theirs, base, yours, 0, kr)
    row = next(row for row, mt in zip(m, meta) if mt[1] == kr + n)
    assert row[1] == ("same", str(n), str(n), str(n))
    assert row[5][0] == "left" and row[5][1] == "7777" and _conf(m) == 0


# ── 33차: 머리글 설명·형식 행, 낱말 단위 리소스 판정, 목록에 없는 리소스 이름 ───────────

@pytest.mark.parametrize("heads", [
    [["", "int32", "FText", "FString", "FText", "int32", "int32"],
     ["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"]],
    [["", "몬스터 번호", "이름(텍스트 키)", "아이콘 이름", "설명", "체력", "공격력"],
     ["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"]],
    [["", "몬스터 번호", "이름 텍스트", "아이콘 이름", "설명", "체력", "공격력"],
     ["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"]],
    [["", "No", "Display name in result popup", "Icon", "Desc", "HP", "ATK"],
     ["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"]],
    [["#", "MonsterID", "Name", "PortraitName", "Desc", "HP", "ATK"]],
    [["#", "MonsterID", "Name", "FXName", "Desc", "HP", "ATK"]],
    [["#", "PresetID", "PresetName", "IconName", "Desc", "HP", "ATK"]],
    [["#", "ProfileID", "ProfileName", "IconName", "Desc", "HP", "ATK"]],
])
@pytest.mark.parametrize("n", [30, 200])
def test_last_row_backup_with_header_notes_and_unlisted_resource_names(heads, n):
    """머리글 위 설명·형식 행('이름(텍스트 키)', FText), 낱말 안에 든 리소스 단어(Preset 의 res, Profile 의 file),
    목록에 없는 리소스 이름(PortraitName)이 표시 이름 판정을 흔들어 백업 복사본을 원본과 짝지었다."""
    test_last_row_backup_with_resource_name_edited_under_various_headers(heads, 3, n)


# ── 34차: 마지막 행 백업은 머리글과 상관없이 복사, 머리글 판정 다듬기 ───────────────

@pytest.mark.parametrize("heads", [
    [["", "int32", "FText", "FName", "FText", "int32", "int32"],
     ["#", "SkillID", "SkillName", "IconName", "Desc", "HP", "ATK"]],
    [["", "int32", "FText", "FName", "FText", "int32", "int32"],
     ["#", "ClassID", "ClassName", "EffectName", "Desc", "HP", "ATK"]],
    [["#", "SkillID", "SkillName", "ResourceName", "Desc", "HP", "ATK"]],
    [["#", "SkillID", "SkillName", "AnimationName", "Desc", "HP", "ATK"]],
    [["#", "ClassID", "ClassName", "TextureName", "Desc", "HP", "ATK"]],
    [["#", "MonsterID", "Name_KR", "IconName", "Desc", "HP", "ATK"]],
    [["#", "몬스터ID", "이름(한글)", "아이콘 이름", "설명", "체력", "공격력"]],
    [["#", "MonsterID", "DisplayName_ko", "IconName", "Desc", "HP", "ATK"]],
    [["#", "SkillID", "SkillName", "MaterialName", "Desc", "HP", "ATK"]],
])
@pytest.mark.parametrize("n", [30, 200])
def test_last_row_backup_is_a_copy_whatever_the_headers(heads, n):
    """마지막 행을 다음 ID 로 백업하고 원본의 리소스 이름 열을 고치면, 머리글 표기(FName 형식 행, ResourceName,
    Name_KR …)에 따라 이동으로 봐 다른 쪽이 원본에 한 수정이 백업에 저장됐다."""
    test_last_row_backup_with_resource_name_edited_under_various_headers(heads, 3, n)


@pytest.mark.parametrize("before", [["PortraitName"], ["ResourceName"], ["AssetName"], ["초상화 이름"]])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_copy_inserted_above_with_a_resource_name_column_before_name(before, shape):
    """목록에 없거나 낱말이 긴 리소스 이름 열(PortraitName·ResourceName)이 Name 앞에 있으면 그 열이 표시 이름으로
    뽑혀 복사본 끼움을 복사로 봤다."""
    test_copy_inserted_above_and_renamed_with_text_columns_before_the_name(before, "Name", shape)


# ── 35차: F 로 시작하는 머리글, 상수 열이 많은 표, 다국어 이름 열 ─────────────────────

@pytest.mark.parametrize("prefix", ["Food", "Furniture", "Formation", "Floor"])
@pytest.mark.parametrize("k", [2, 3])
def test_backup_of_the_last_rows_with_f_prefixed_headers(prefix, k):
    """FoodName·FloorName 처럼 F 로 시작하는 머리글이 형식 칸(FName)으로 버려져, 끝 여러 행 백업 + 원본 아이콘
    수정에서 다른 쪽 수정이 백업에 저장됐다."""
    rng = random.Random(k)
    n = 30
    H = ["#", f"{prefix}ID", f"{prefix}Name", "IconName", "Desc", "HP", "ATK"]
    rows = [["", str(i + 1), f"Nm{i}", f"ico_{i:03d}", f"desc_{i:03d}", str(rng.randint(100, 9999)),
             str(rng.randint(100, 9999))] for i in range(n)]
    body = [r[:] for r in rows]
    for d in range(k):
        cp = rows[n - k + d][:]
        cp[1] = str(n + 1 + d)
        body.append(cp)
        body[n - k + d][3] = body[n - k + d][3] + "_v2"
    base = [H] + rows
    yours = [H] + body
    theirs = [r[:] for r in base]
    for d in range(k):
        theirs[n - k + d + 1][5] = "7777"
    m, meta = compute_diff3(theirs, base, yours, 0, 0)
    for d in range(k):
        row = next(row for row, mt in zip(m, meta) if mt[1] == n - k + d + 1)
        assert row[1][0] == "same" and row[5][0] == "left" and row[5][1] == "7777"
    assert _conf(m) == 0


@pytest.mark.parametrize("flags", [6, 8])
@pytest.mark.parametrize("n", [30, 200])
def test_new_row_before_the_last_with_many_constant_columns(flags, n):
    """기본값 열(모두 0)이 많으면, 마지막 행 앞에 전혀 다른 새 행을 끼우고 다시 매긴 것을 백업으로 봐 원본 수정이
    새 행에 저장됐다."""
    rng = random.Random(flags)
    H = ["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"] + [f"Flag{i}" for i in range(flags)]
    rows = [[f"Nm{i}", f"ico_{i:03d}", f"설명_{i:03d}", str(rng.randint(100, 9999)), str(rng.randint(100, 9999))]
            + ["0"] * flags for i in range(n)]
    body = [r[:] for r in rows]
    body.insert(n - 1, ["NewMob", "ico_new", "설명_new", "1234", "567"] + ["0"] * flags)

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    yours = [H] + sheet(body)
    theirs = [r[:] for r in base]
    theirs[n][7] = "1"                          # 원본 마지막 몹의 Flag0
    m, meta = compute_diff3(theirs, base, yours, 0, 0)
    row = next(row for row in m if row[2][1] == f"Nm{n - 1}" and row[2][2] == f"Nm{n - 1}")
    assert row[2][3] == f"Nm{n - 1}" and row[1][3] == str(n + 1)
    assert row[7][0] == "left" and row[7][1] == "1"


@pytest.mark.parametrize("cols,edit", [(["Name_KR", "Name"], 1), (["Name_EN", "Name_KR"], 1)])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_copy_inserted_above_with_several_language_name_columns(cols, edit, shape):
    """다국어 이름 열이 여럿일 때 왼쪽이 아닌 이름 열만 바꾼 복사본을 위에 끼우면 복사로 봐 다시 매김을 놓쳤다."""
    rng = random.Random(3)
    n = 60
    H = ["#", "SortKey"] + cols + ["Grade", "HP", "ATK"]
    rows = [[f"{c}_{i:03d}" for c in cols] + [str(rng.randint(1, 5)), str(rng.randint(100, 999)),
                                              str(rng.randint(1000, 4999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[edit] = r[edit] + "_Hard"
        v[2] = "5"
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        idx = list(range(n - 3, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[i + 1][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    for i in idx:
        row = next(row for row in m if row[2 + edit][2] == rows[i][edit] and row[6][3] == "7777")
        assert row[2 + edit][1] == rows[i][edit] and row[6][0] == "right"


# ── 36차: 숫자 TextID 이름, 드문드문 값이 있는 기본값 열, 앞에 붙는 언어 표기, 템플릿 형식 칸 ─────────

def _backup_last_k(heads, name_of, k, n):
    """끝 k행을 n+1.. 로 백업하고 원본 IconName 을 고친 쪽(yours)과 원본 HP 를 고친 쪽(theirs) — 원본 수정은
    원본 자리(1..n)에 남아야 한다."""
    rng = random.Random(n + k)
    rows = [["", str(i + 1), name_of(i), f"ico_{i:03d}", f"desc_{i:03d}", str(rng.randint(100, 9999)),
             str(rng.randint(100, 9999))] for i in range(n)]
    body = [r[:] for r in rows]
    for d in range(k):
        cp = rows[n - k + d][:]
        cp[1] = str(n + 1 + d)
        body.append(cp)
        body[n - k + d][3] = body[n - k + d][3] + "_v2"
    kr = len(heads) - 1
    base = [h[:] for h in heads] + rows
    yours = [h[:] for h in heads] + body
    theirs = [r[:] for r in base]
    for d in range(k):
        theirs[kr + n - k + d + 1][5] = "7777"
    m, meta = compute_diff3(theirs, base, yours, 0, kr)
    for d in range(k):
        num = str(n - k + d + 1)
        row = next(row for row, mt in zip(m, meta) if mt[1] == kr + n - k + d + 1)
        assert row[1] == ("same", num, num, num)
        assert row[5][0] == "left" and row[5][1] == "7777"
    assert _conf(m) == 0


@pytest.mark.parametrize("heads,name_of", [
    ([["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"]], lambda i: str(1000100 + i)),
    ([["", "int32", "FText", "FName", "FText", "int32", "int32"],
      ["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"]], lambda i: str(1000100 + i)),
    ([["#", "MonsterID", "NameTextID", "IconName", "Desc", "HP", "ATK"]], lambda i: str(1000100 + i)),
    ([["#", "MonsterID", "NameKey", "IconName", "Desc", "HP", "ATK"]], lambda i: f"TEXT_MON_{i:03d}"),
    ([["#", "MonsterID", "Name_TID", "IconName", "Desc", "HP", "ATK"]], lambda i: str(500 + i)),
])
@pytest.mark.parametrize("k", [2, 3])
@pytest.mark.parametrize("n", [30, 200])
def test_backup_of_the_last_rows_with_a_text_id_name(heads, name_of, k, n):
    """이름 열이 숫자 TextID(Name=1000123)·NameKey 이면 이름 열로 안 봐, 끝 여러 행 백업 + 원본 아이콘 수정에서
    다른 쪽이 원본에 한 수정이 백업에 저장됐다."""
    _backup_last_k(heads, name_of, k, n)


@pytest.mark.parametrize("flags", [6, 8])
@pytest.mark.parametrize("n", [30, 200])
@pytest.mark.parametrize("sparse", [1, 2])
def test_new_row_before_the_last_with_mostly_default_columns(flags, n, sparse):
    """기본값 열에 1 이 몇 칸만 있어도(대부분 0) 그 열을 '같음'으로 세어, 마지막 행 앞에 끼운 새 행을 다시
    원본과 짝지었다."""
    rng = random.Random(flags * 7 + sparse)
    H = ["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"] + [f"Flag{i}" for i in range(flags)]
    rows = [[f"Nm{i}", f"ico_{i:03d}", f"설명_{i:03d}", str(rng.randint(100, 9999)), str(rng.randint(100, 9999))]
            + ["0"] * flags for i in range(n)]
    for f in range(flags):
        for _ in range(sparse):
            rows[rng.randrange(1, n - 2)][5 + f] = "1"
    body = [r[:] for r in rows]
    body.insert(n - 1, ["NewMob", "ico_new", "설명_new", "1234", "567"] + ["0"] * flags)

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    yours = [H] + sheet(body)
    theirs = [r[:] for r in base]
    theirs[n][7] = "1"
    m, meta = compute_diff3(theirs, base, yours, 0, 0)
    row = next(row for row in m if row[2][1] == f"Nm{n - 1}" and row[2][2] == f"Nm{n - 1}")
    assert row[2][3] == f"Nm{n - 1}" and row[1][3] == str(n + 1)
    assert row[7][0] == "left" and row[7][1] == "1"


@pytest.mark.parametrize("cols", [
    ["MonsterName_KR", "MonsterName_EN"], ["KR_Name", "EN_Name"], ["KorName", "EngName"],
    ["한글 이름", "영문 이름"], ["한글이름", "영문이름"], ["Name", "EngName"], ["이름", "영문 이름"],
])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_copy_inserted_above_with_language_prefixed_name_columns(cols, shape):
    """KorName·EN_Name·'영문 이름'처럼 언어 표기가 앞에 붙거나 MonsterName_KR/_EN 처럼 …Name 이 여럿이면 왼쪽 열
    하나만 봐, 오른쪽 이름만 바꾼 복사본 끼움을 복사로 판정했다."""
    test_copy_inserted_above_with_several_language_name_columns(cols, 1, shape)


@pytest.mark.parametrize("ty", ["TArray<FName>", "TSet<FName>", "TArray<FString>"])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_copy_inserted_above_with_a_template_type_row(ty, shape):
    """형식 행의 템플릿 형식(TArray<FName>)이 name 으로 끝나 그 배열 열이 표시 이름으로 뽑혀, 이름 바꾼 복사본
    끼움을 복사로 판정했다."""
    rng = random.Random(4)
    n = 60
    T = ["", "int32", ty, "FText", "int32", "int32", "int32"]
    H = ["#", "SortKey", "Drops", "MonsterName", "Grade", "HP", "ATK"]
    rows = [[f"(Drop_{i % 7},Drop_{i % 11})", f"MonsterName_{i:03d}", str(rng.randint(1, 5)),
             str(rng.randint(100, 999)), str(rng.randint(1000, 4999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[1] = r[1] + "_Hard"
        v[2] = "5"
        v[3] = str(int(r[3]) * 2)
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        idx = list(range(n - 3, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [T, H] + sheet(rows)
    left = [T, H] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[i + 2][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 1)
    for i in idx:
        row = next(row for row in m if row[3][2] == rows[i][1] and row[6][3] == "7777")
        assert row[3][1] == rows[i][1] and row[6][0] == "right"


# ── 37차: 숫자 TextID 인 …Name, 0/1 Bool 열, 왼쪽의 범주형 …Name ─────────────────────

@pytest.mark.parametrize("head,name_of", [
    ("MonsterName", lambda i: str(1000100 + i)),
    ("HeroName", lambda i: str(1000100 + i)),
    ("DisplayName", lambda i: str(1000100 + i)),
    ("몬스터 이름", lambda i: str(1000100 + i)),
    ("HeroNameTextID", lambda i: str(1000100 + i)),
    ("MonsterNameKey", lambda i: f"TEXT_MON_{i:03d}"),
    ("Name_FR", lambda i: f"Nom{i}"),
    ("Name_en-US", lambda i: f"Nm{i}"),
])
@pytest.mark.parametrize("k", [2, 3])
@pytest.mark.parametrize("n", [30, 200])
def test_backup_of_the_last_rows_with_a_text_id_display_name(head, name_of, k, n):
    """MonsterName·DisplayName 값이 숫자 TextID 이거나 HeroNameTextID·Name_FR 같은 머리글이면 이름 열로 안 봐, 끝
    여러 행 백업 + 원본 아이콘 수정에서 다른 쪽이 원본에 한 수정이 백업에 저장됐다."""
    _backup_last_k([["#", "MonsterID", head, "IconName", "Desc", "HP", "ATK"]], name_of, k, n)


@pytest.mark.parametrize("flags,ratio", [(8, 0.3), (12, 0.3), (12, 0.2), (16, 0.4)])
@pytest.mark.parametrize("seed", range(6))
def test_new_row_before_the_last_with_many_bool_columns(flags, ratio, seed):
    """1 이 20~40% 인 0/1 열(IsBoss·CanCapture)이 많으면 서로 다른 행도 우연히 절반 넘게 같아, 마지막 행 앞에 끼운
    새 행을 원본과 짝지었다."""
    rng = random.Random(seed * 100 + flags)
    n = 60
    H = ["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"] + [f"Is{i}" for i in range(flags)]
    rows = [[f"Nm{i}", f"ico_{i:03d}", f"설명_{i:03d}", str(rng.randint(100, 9999)), str(rng.randint(100, 9999))]
            + ["1" if rng.random() < ratio else "0" for _ in range(flags)] for i in range(n)]
    rows[n - 1][5] = "0"
    body = [r[:] for r in rows]
    body.insert(n - 1, ["NewMob", "ico_new", "설명_new", "1234", "567"] + ["0"] * flags)

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    yours = [H] + sheet(body)
    theirs = [r[:] for r in base]
    theirs[n][7] = "1"
    m, meta = compute_diff3(theirs, base, yours, 0, 0)
    row = next(row for row in m if row[2][1] == f"Nm{n - 1}" and row[2][2] == f"Nm{n - 1}")
    assert row[2][3] == f"Nm{n - 1}" and row[1][3] == str(n + 1)
    assert row[7][0] == "left" and row[7][1] == "1"


@pytest.mark.parametrize("cols,edit", [
    (["StageName", "MonsterName"], 1), (["MapName", "MonsterName"], 1), (["스테이지 이름", "몬스터 이름"], 1),
    (["StageName_KR", "StageName_EN", "MonsterName_KR", "MonsterName_EN"], 2),
])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_copy_inserted_above_with_a_categorical_name_on_the_left(cols, edit, shape):
    """StageName 처럼 몇 행씩 같은 값인 …Name 이 몬스터 이름보다 왼쪽에 있으면 그 열이 표시 이름으로 뽑혀, 몬스터
    이름을 바꾼 복사본 끼움을 복사로 판정했다."""
    rng = random.Random(5)
    n = 60
    H = ["#", "SortKey"] + cols + ["Grade", "HP", "ATK"]
    nc = len(cols)
    rows = []
    for i in range(n):
        vals = [f"{c}_{i // 5:02d}" if "Stage" in c or "Map" in c or "스테이지" in c else f"{c}_{i:03d}" for c in cols]
        rows.append(vals + [str(rng.randint(1, 5)), str(rng.randint(100, 999)), str(rng.randint(1000, 4999))])
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[edit] = r[edit] + "_Hard"
        v[nc] = "5"
        v[nc + 1] = str(int(r[nc + 1]) * 2)
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        idx = list(range(n - 3, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    atk = len(H) - 1
    for i in idx:
        right[i + 1][atk] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    for i in idx:
        row = next(row for row in m if row[2 + edit][2] == rows[i][edit] and row[atk][3] == "7777")
        assert row[2 + edit][1] == rows[i][edit] and row[atk][0] == "right"


# ── 38차: 기본값 enum 열이 많은 표, 묶음 이름 열(StageName)과 되풀이되는 몬스터 이름 ─────────────

@pytest.mark.parametrize("enums,dist", [(10, (0.62, 0.28, 0.10)), (12, (0.55, 0.25, 0.12, 0.08)),
                                        (12, (0.62, 0.28, 0.10))])
@pytest.mark.parametrize("seed", range(8))
def test_new_row_before_the_last_with_many_default_enum_columns(enums, dist, seed):
    """기본값(0)이 60% 쯤인 enum 열(AIType·MoveType·Size …)이 식별 열보다 훨씬 많으면, 마지막 행 앞에 끼운 기본값
    새 행을 원본과 절반 넘게 같다고 봐 백업으로 판정했다."""
    rng = random.Random(seed * 31 + enums)
    n = 200
    vals = [str(i) for i in range(len(dist))]
    H = ["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"] + [f"Enum{i}" for i in range(enums)]
    rows = [[f"Nm{i}", f"ico_{i:03d}", f"설명_{i:03d}", str(rng.randint(100, 9999)), str(rng.randint(100, 9999))]
            + rng.choices(vals, weights=dist, k=enums) for i in range(n)]
    rows[n - 1][5:] = ["0"] * enums
    rows[n - 1][6] = "1"
    body = [r[:] for r in rows]
    body.insert(n - 1, ["NewMob", "ico_new", "설명_new", "1234", "567"] + ["0"] * enums)

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    yours = [H] + sheet(body)
    theirs = [r[:] for r in base]
    theirs[n][8] = "9"                          # 원본 마지막 몹의 Enum1
    m, meta = compute_diff3(theirs, base, yours, 0, 0)
    row = next(row for row in m if row[2][1] == f"Nm{n - 1}" and row[2][2] == f"Nm{n - 1}")
    assert row[2][3] == f"Nm{n - 1}" and row[1][3] == str(n + 1)
    assert row[8][0] == "left" and row[8][1] == "9"


def test_last_row_backup_with_many_default_enum_columns_is_still_a_backup():
    """위 판정을 좁혀도 마지막 행 백업(원본 아이콘 수정)은 여전히 백업이다."""
    rng = random.Random(1)
    n, enums = 200, 12
    H = ["#", "MonsterID", "Name", "IconName", "Desc", "HP", "ATK"] + [f"Enum{i}" for i in range(enums)]
    rows = [["", str(i + 1), f"Nm{i}", f"ico_{i:03d}", f"설명_{i:03d}", str(rng.randint(100, 9999)),
             str(rng.randint(100, 9999))] + rng.choices("012", weights=(0.6, 0.3, 0.1), k=enums) for i in range(n)]
    body = [r[:] for r in rows]
    cp = rows[-1][:]
    cp[1] = str(n + 1)
    body.append(cp)
    body[n - 1][3] = body[n - 1][3] + "_v2"
    base = [H] + rows
    yours = [H] + body
    theirs = [r[:] for r in base]
    theirs[n][5] = "7777"
    m, meta = compute_diff3(theirs, base, yours, 0, 0)
    row = next(row for row, mt in zip(m, meta) if mt[1] == n)
    assert row[1] == ("same", str(n), str(n), str(n))
    assert row[5][0] == "left" and row[5][1] == "7777" and _conf(m) == 0


@pytest.mark.parametrize("group", ["StageName", "MapName", "스테이지 이름", "GroupName"])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
@pytest.mark.parametrize("stats", [False, True])
def test_spawn_table_copy_with_another_monster_inserted_above(group, shape, stats):
    """스폰 표(StageName 5행씩 + 여러 스테이지에 되풀이되는 MonsterName)에서 몬스터만 바꾼 복사본을 위에 끼우면
    StageName 이 표시 이름으로 뽑혀 복사로 판정했다."""
    rng = random.Random(9)
    n = 60
    mon = "몬스터 이름" if group == "스테이지 이름" else "MonsterName"
    H = ["#", "SortKey", group, mon, "Grade", "HP", "ATK"]
    rows = [[f"Stage_{i // 5:02d}", f"Mon_{rng.randrange(12):02d}", str(rng.randint(1, 5)),
             str(rng.randint(100, 999)), str(rng.randint(1000, 4999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[1] = "Mon_99"
        if stats:
            v[2] = "5"
            v[3] = str(int(r[3]) * 2)
        return v
    if shape in ("last3", "last2"):
        k = int(shape[-1])
        body[n - k:n - k] = [variant(r) for r in rows[n - k:]]
        idx = list(range(n - k, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[i + 1][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    hits = [row for row in m if row[6][3] == "7777"]
    assert len(hits) == len(idx)
    for row in hits:
        assert row[3][1] == row[3][2] != "Mon_99" and row[6][0] == "right"


@pytest.mark.parametrize("group", ["StageName", "GroupName", "MapName", "RewardGroupName"])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_copy_inserted_above_with_only_a_group_name_column(group, shape):
    """31차 '범주형 이름 열만 있는 표'와 같은 모양에서 ElementName 대신 StageName·GroupName 이면 그 열이 표시
    이름이 되어 Icon 을 바꾼 복사본 끼움을 복사로 판정했다."""
    rng = random.Random(6)
    n = 60
    H = ["#", "SortKey", "Icon", group, "Grade", "HP", "ATK"]
    rows = [[f"Icon_{i:03d}", f"G_{i // 5:02d}", str(rng.randint(1, 5)), str(rng.randint(100, 999)),
             str(rng.randint(1000, 4999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[0] = r[0] + "_Hard"
        v[2] = "5"
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        idx = list(range(n - 3, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[i + 1][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    for i in idx:
        row = next(row for row in m if row[2][2] == rows[i][0])
        assert row[2][1] == rows[i][0] and row[6][0] == "right" and row[6][3] == "7777"


# ── 39차: 스테이지 표의 행마다 다른 StageName 은 표시 이름 ─────────────────────────────

@pytest.mark.parametrize("cols", [
    ["StageName", "BossName"], ["스테이지 이름", "보스 이름"], ["MapName", "BossName"],
    ["StageName", "MonsterName"], ["StageName", "RewardItemName"], ["ChapterName", "BossName"],
])
@pytest.mark.parametrize("boss", ["repeat", "unique"])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
def test_stage_table_hard_copy_inserted_above(cols, boss, shape):
    """38차에 StageName 을 머리글만으로 분류 열로 내려, 스테이지 표(행마다 다른 StageName)에서 '(어려움)' 복사본을
    위에 끼우면 오른쪽 BossName 이 표시 이름이 되어 복사로 판정했다."""
    rng = random.Random(11)
    n = 60
    H = ["#", "SortKey"] + cols + ["RecommendPower", "Stamina", "Exp"]
    rows = [[f"{i // 10 + 1}-{i % 10 + 1} 숲길_{i:02d}", f"Boss_{i % 8}" if boss == "repeat" else f"Boss_{i:03d}",
             str(rng.randint(1000, 9999)), str(rng.randint(5, 20)), str(rng.randint(100, 999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[0] = r[0] + " (어려움)"
        v[2] = str(int(r[2]) * 3)
        return v
    if shape in ("last3", "last2"):
        k = int(shape[-1])
        body[n - k:n - k] = [variant(r) for r in rows[n - k:]]
        idx = list(range(n - k, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[i + 1][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    for i in idx:
        row = next(row for row in m if row[6][3] == "7777" and row[2][2] == rows[i][0])
        assert row[2][1] == rows[i][0] and row[6][0] == "right"


# ── 40차: 스테이지당 1~3행인 스폰 표, 묶음 낱말이 수식어인 몬스터 이름 열 ───────────────────

def _spawn_copy_case(heads, group_of, shape, stats, seed):
    """키 빈 60행 스폰 표에서 몬스터만 Mon_99 로 바꾼 복사본을 위에 끼우고 다시 매김(right 는 원본 ATK 7777) —
    7777 은 원본 몬스터 행에 남아야 한다. heads 의 마지막 앞 두 칸이 [묶음 열, 몬스터 열]."""
    rng = random.Random(seed)
    n = 60
    H = ["#", "SortKey"] + heads + ["Grade", "HP", "ATK"]
    groups = group_of(rng, n)
    rows = [[groups[i], f"Mon_{rng.randrange(12):02d}", str(rng.randint(1, 5)), str(rng.randint(100, 999)),
             str(rng.randint(1000, 4999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[1] = "Mon_99"
        if stats:
            v[2] = "5"
            v[3] = str(int(r[3]) * 2)
        return v
    if shape in ("last3", "last2"):
        k = int(shape[-1])
        body[n - k:n - k] = [variant(r) for r in rows[n - k:]]
        idx = list(range(n - k, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[i + 1][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    hits = [row for row in m if row[6][3] == "7777"]
    assert len(hits) == len(idx)
    for row in hits:
        assert row[3][1] == row[3][2] != "Mon_99" and row[6][0] == "right"


def _groups_of(sizes):
    def make(rng, n):
        out = []
        g = 0
        while len(out) < n:
            out += [f"Stage_{g:02d}"] * rng.choice(sizes)
            g += 1
        return out[:n]
    return make


@pytest.mark.parametrize("sizes", [(1, 2), (1, 2, 3), (1, 1, 1, 2)])
@pytest.mark.parametrize("heads", [["StageName", "MonsterName"], ["스테이지 이름", "몬스터 이름"],
                                   ["WaveName", "MonsterName"]])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
@pytest.mark.parametrize("seed", range(3))
def test_spawn_table_with_few_rows_per_stage(sizes, heads, shape, seed):
    """스테이지당 1~3행이면 StageName 값이 절반 넘게 달라(묶음인데도) 표시 이름으로 뽑혀, 몬스터만 바꾼 복사본
    끼움을 복사로 판정했다."""
    _spawn_copy_case(heads, _groups_of(sizes), shape, seed % 2 == 0, seed)


@pytest.mark.parametrize("heads", [["StageID", "WaveMonsterName"], ["StageName", "WaveMonsterName"],
                                   ["GroupName", "StageMonsterName"], ["StageID", "StageBossName"]])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
@pytest.mark.parametrize("seed", range(3))
def test_spawn_table_with_a_group_word_modifying_the_monster_name(heads, shape, seed):
    """WaveMonsterName·StageBossName 처럼 묶음 낱말이 수식어로 붙은 몬스터 이름 열까지 묶음 열로 내려, 몬스터만
    바꾼 복사본 끼움을 복사로 판정했다."""
    def five(rng, n):
        return [f"{1000 + i // 5}" if heads[0] == "StageID" else f"Stage_{i // 5:02d}" for i in range(n)]
    _spawn_copy_case(heads, five, shape, seed % 2 == 0, seed)


@pytest.mark.parametrize("dups", [0, 2, 4])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_stage_table_with_a_few_repeated_stage_names_keeps_the_display_name(dups, shape):
    """스테이지 표에 같은 이름 스테이지가 몇 개 있어도 StageName 은 그 표의 표시 이름이다."""
    rng = random.Random(dups)
    n = 60
    H = ["#", "SortKey", "StageName", "BossName", "RecommendPower", "Stamina", "Exp"]
    names = [f"Stage_{i:03d}" for i in range(n)]
    for d in range(dups):
        names[5 + d * 7] = names[4 + d * 7]
    rows = [[names[i], f"Boss_{i % 8}", str(rng.randint(1000, 9999)), str(rng.randint(5, 20)),
             str(rng.randint(100, 999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[0] = r[0] + " (Hard)"
        v[2] = str(int(r[2]) * 2)
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        idx = list(range(n - 3, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[i + 1][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    for i in idx:
        row = next(row for row in m if row[6][3] == "7777" and row[2][2] == rows[i][0])
        assert row[2][1] == rows[i][0] and row[6][0] == "right"


# ── 41차: 붙여 쓴 한글 묶음 머리글, StageDisplayName, 장 구분 주석 행 ─────────────────────

@pytest.mark.parametrize("heads", [["스테이지이름", "몬스터이름"], ["스테이지이름", "몬스터 이름"], ["웨이브이름", "몬스터이름"],
                                   ["맵이름", "보스이름"], ["StageDisplayName", "MonsterName"]])
@pytest.mark.parametrize("sizes", [(5,), (1, 2)])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
@pytest.mark.parametrize("seed", range(2))
def test_spawn_table_with_joined_korean_group_headers(heads, sizes, shape, seed):
    """'스테이지이름'처럼 붙여 쓴 한글 머리글은 낱말 하나로 읽혀 묶음 열로 안 봐, 몬스터만 바꾼 복사본 끼움을 복사로
    판정했다(40차 머리 낱말 규칙의 회귀)."""
    _spawn_copy_case(heads, _groups_of(sizes), shape, seed % 2 == 0, seed)


@pytest.mark.parametrize("heads", [["스테이지ID", "웨이브몬스터이름"], ["StageID", "WaveDisplayMonsterName"]])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_joined_korean_monster_name_with_a_group_modifier_stays_a_name(heads, shape):
    """'웨이브몬스터이름'은 머리 낱말이 몬스터라 몬스터 이름 열이다."""
    def five(rng, n):
        return [f"{1000 + i // 5}" for i in range(n)]
    _spawn_copy_case(heads, five, shape, True, 1)


@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_spawn_table_with_a_joined_korean_note_row_under_english_headers(shape):
    """영문 머리글 아래 붙여 쓴 한글 설명 행('스테이지이름')이 StageName 을 표시 이름으로 끌어올렸다."""
    rng = random.Random(3)
    n = 60
    note = ["", "정렬", "스테이지이름", "몬스터이름", "등급", "체력", "공격력"]
    H = ["#", "SortKey", "StageName", "MonsterName", "Grade", "HP", "ATK"]
    rows = [[f"Stage_{i // 5:02d}", f"Mon_{rng.randrange(12):02d}", str(rng.randint(1, 5)),
             str(rng.randint(100, 999)), str(rng.randint(1000, 4999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[1] = "Mon_99"
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        idx = list(range(n - 3, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [note, H] + sheet(rows)
    left = [note, H] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[i + 2][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 1)
    hits = [row for row in m if row[6][3] == "7777"]
    assert len(hits) == len(idx)
    for row in hits:
        assert row[3][1] == row[3][2] != "Mon_99" and row[6][0] == "right"


@pytest.mark.parametrize("every", [4, 6])
@pytest.mark.parametrize("heads", [["StageName", "BossName"], ["스테이지 이름", "보스 이름"]])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_stage_table_with_chapter_comment_rows_keeps_the_display_name(every, heads, shape):
    """'// 챕터 N' 주석 행의 빈칸까지 값으로 세어 StageName 이 90% 고유에 못 미쳐 묶음 열로 내려가, '(어려움)'
    복사본 끼움을 복사로 판정했다."""
    rng = random.Random(every)
    n = 60
    H = ["#", "SortKey"] + heads + ["RecommendPower", "Stamina", "Exp"]
    rows = [[f"{i // 10 + 1}-{i % 10 + 1} 숲길_{i:02d}", f"Boss_{i % 8}", str(rng.randint(1000, 9999)),
             str(rng.randint(5, 20)), str(rng.randint(100, 999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[0] = r[0] + " (어려움)"
        v[2] = str(int(r[2]) * 3)
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        idx = list(range(n - 3, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        out = []
        for i, r in enumerate(rs):
            if i % every == 0:
                out.append([f"// 챕터 {i // every + 1}", "", "", "", "", "", ""])
            out.append(["", str(i + 1)] + r)
        return out
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    for r in right:
        if r[2] in [rows[i][0] for i in idx]:
            r[6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    for i in idx:
        row = next(row for row in m if row[6][3] == "7777" and row[2][2] == rows[i][0])
        assert row[2][1] == rows[i][0] and row[6][0] == "right"


# ── 42차: 묶음 첫 행에만 이름을 적은 스폰 표·통째로 빈 묶음 열, '…명' 머리글 ─────────────────

def _first_row_only(sizes):
    def make(rng, n):
        out = []
        g = 0
        while len(out) < n:
            s = rng.choice(sizes)
            out += [f"Stage_{g:02d}"] + [""] * (s - 1)
            g += 1
        return out[:n]
    return make


@pytest.mark.parametrize("heads", [["StageName", "MonsterName"], ["스테이지 이름", "몬스터 이름"],
                                   ["스테이지이름", "몬스터이름"], ["WaveName", "MonsterName"]])
@pytest.mark.parametrize("sizes", [(5,), (1, 2), (2, 3, 4)])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
@pytest.mark.parametrize("seed", range(2))
def test_spawn_table_with_the_stage_name_only_on_the_first_row(heads, sizes, shape, seed):
    """묶음 첫 행에만 StageName 을 적으면(나머지 빈칸) 빈칸을 뺀 값이 모두 달라 '고유'로 봐, 몬스터만 바꾼 복사본
    끼움을 복사로 판정했다(41차 회귀)."""
    _spawn_copy_case(heads, _first_row_only(sizes), shape, seed % 2 == 0, seed)


@pytest.mark.parametrize("heads", [["StageName", "MonsterName"], ["스테이지 이름", "몬스터 이름"]])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
def test_spawn_table_with_an_empty_stage_name_column(heads, shape):
    """아직 채우지 않아 통째로 빈 StageName 열을 '고유'로 봤다."""
    _spawn_copy_case(heads, lambda rng, n: [""] * n, shape, True, 2)


@pytest.mark.parametrize("heads", [["스테이지ID", "몬스터명"], ["그룹ID", "아이템명"], ["스테이지ID", "몬스터 명칭"],
                                   ["스테이지 명", "몬스터 명"], ["스테이지명", "몬스터명"], ["영문스테이지명", "영문몬스터명"]])
@pytest.mark.parametrize("sizes", [(5,), (1, 2)])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
def test_spawn_table_with_korean_myeong_headers(heads, sizes, shape):
    """'몬스터명'·'아이템명'·'몬스터 명칭' 같은 '…명' 머리글을 이름 열로 안 봐, 몬스터만 바꾼 복사본 끼움을 복사로
    판정했다."""
    def groups(rng, n):
        if heads[0].endswith("ID"):
            return [f"{1000 + i // 5}" for i in range(n)]
        return _groups_of(sizes)(rng, n)
    _spawn_copy_case(heads, groups, shape, True, 3)


@pytest.mark.parametrize("n", [30, 200])
def test_last_row_backup_with_a_korean_myeong_name_and_description(n):
    """'몬스터명'·'아이콘명'·'설명' 머리글 표의 마지막 행 백업(원본 아이콘명 수정)은 백업이다 — '설명' 은 이름 열이
    아니다."""
    test_last_row_backup_with_resource_name_edited_under_various_headers(
        [["#", "몬스터ID", "몬스터명", "아이콘명", "설명", "체력", "공격력"]], 3, n)


# ── 43차: 인원 단위 '명'·'(명)', 끝이나 한 장에 몰린 미작성 StageName ─────────────────────────

def _stage_hard_copy(H, rows, name_col, unit=None, kr=0, shape="mid_copy_del_next"):
    """스테이지 표에서 '(어려움)' 복사본을 위에 끼우고 다시 매김(right 는 원본 마지막 열 7777) — 7777 은 원본
    행에 남아야 한다. rows 는 번호 뒤 값들, name_col 은 그 안의 이름 칸."""
    n = len(rows)
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[name_col] = r[name_col] + " (어려움)"
        v[name_col + 2] = str(int(r[name_col + 2]) * 3)
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        idx = list(range(n - 3, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    top = [H] + ([unit] if unit else [])
    base = top + sheet(rows)
    left = top + sheet(body)
    right = [r[:] for r in base]
    last = len(H) - 1
    for i in idx:
        right[len(top) + i][last] = "7777"
    m, meta = compute_diff3(left, base, right, 0, kr)
    c = 2 + name_col
    for i in idx:
        row = next(row for row in m if row[last][3] == "7777" and row[c][2] == rows[i][name_col])
        assert row[c][1] == rows[i][name_col] and row[last][0] == "right"


@pytest.mark.parametrize("unit", ["명", "(명)", "[명]", "단위: 명"])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_stage_table_with_a_unit_row_of_myeong(unit, shape):
    """머리글 아래 단위 행의 '명' 칸을 딱 '이름'으로 봐, 숫자 인원 열이 표시 이름이 되어 '(어려움)' 복사본 끼움을
    복사로 판정했다(42차 회귀)."""
    rng = random.Random(4)
    H = ["#", "SortKey", "StageName", "BossName", "RecommendPower", "MaxHero", "Stamina", "Exp"]
    U = ["", "", "", "", "", unit, "", ""]
    rows = [[f"Stage_{i:03d}", f"Boss_{i % 8}", str(rng.randint(1000, 9999)), str(rng.choice((3, 4, 5))),
             str(rng.randint(5, 20)), str(rng.randint(100, 999))] for i in range(60)]
    _stage_hard_copy(H, rows, 0, unit=U, kr=0, shape=shape)


@pytest.mark.parametrize("member", ["출전 인원(명)", "최대 인원(명)", "인원 [명]"])
@pytest.mark.parametrize("hero", ["HeroName", "영웅명"])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
def test_spawn_table_with_a_member_count_column_in_myeong(member, hero, shape):
    """'출전 인원(명)' 숫자 열이 '…명' 규칙으로 표시 이름이 되어, 영웅만 바꾼 복사본 끼움을 복사로 판정했다."""
    rng = random.Random(7)
    n = 60
    H = ["#", "SortKey", "StageID", member, hero, "Grade", "HP", "ATK"]
    rows = [[str(1000 + i // 5), str(rng.choice((3, 4, 5))), f"Hero_{rng.randrange(12):02d}", str(rng.randint(1, 5)),
             str(rng.randint(100, 999)), str(rng.randint(1000, 4999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[2] = "Hero_99"
        return v
    if shape in ("last3", "last2"):
        k = int(shape[-1])
        body[n - k:n - k] = [variant(r) for r in rows[n - k:]]
        idx = list(range(n - k, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[i + 1][7] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    hits = [row for row in m if row[7][3] == "7777"]
    assert len(hits) == len(idx)
    for row in hits:
        assert row[4][1] == row[4][2] != "Hero_99" and row[7][0] == "right"


@pytest.mark.parametrize("blank", ["tail10", "tail12", "tail20", "mid10", "scatter4"])
@pytest.mark.parametrize("heads", [["StageName", "BossName"], ["스테이지 이름", "보스 이름"]])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_stage_table_with_unwritten_stage_names_keeps_the_display_name(blank, heads, shape):
    """다음 장 행만 먼저 만들고 StageName 을 비워 두면(끝이나 한 장에 10행 이상) 그 빈칸을 되풀이 값으로 세어
    StageName 을 묶음 열로 내려, '(어려움)' 복사본 끼움을 복사로 판정했다(42차 회귀)."""
    rng = random.Random(5)
    n = 60
    names = [f"{i // 10 + 1}-{i % 10 + 1} 숲길_{i:02d}" for i in range(n)]
    if blank.startswith("tail"):
        k = int(blank[4:])
        empty = range(n - k, n)
    elif blank == "mid10":
        empty = range(40, 50)
    else:
        empty = (5, 17, 44, 52)
    for i in empty:
        names[i] = ""
    H = ["#", "SortKey"] + heads + ["RecommendPower", "Stamina", "Exp"]
    rows = [[names[i], f"Boss_{i % 8}", str(rng.randint(1000, 9999)), str(rng.randint(5, 20)),
             str(rng.randint(100, 999))] for i in range(n)]
    if shape == "last3" and blank.startswith("tail"):
        pytest.skip("끝 행이 이름 없음 — 끼움 대상이 아님")
    _stage_hard_copy(H, rows, 0, shape=shape)


# ── 44차: 2단 머리글을 키 행 0 으로 연 표, 묶음 1~2개, 숫자에 붙은 인원 단위 ─────────────────────

@pytest.mark.parametrize("order", ["en_ko", "ko_en"])
@pytest.mark.parametrize("num", ["SortKey", "No"])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
@pytest.mark.parametrize("seed", range(2))
def test_two_row_header_opened_with_key_row_0(order, num, shape, seed):
    """영문 + 한글 설명 2단 머리글을 기본 키 행 0 으로 열면 설명 행의 '정렬' 이 번호 열 첫 값이 되어 다시 매김
    감지가 꺼지고, '(어려움)' 복사본 끼움을 원본과 짝지었다."""
    rng = random.Random(seed)
    n = 60
    en = ["#", num, "StageName", "BossName", "RecommendPower", "Stamina", "Exp"]
    ko = ["", "정렬", "스테이지 이름", "보스 이름", "추천 전투력", "스태미나", "경험치"]
    top = [en, ko] if order == "en_ko" else [ko, en]
    rows = [[f"Stage_{i:03d}", f"Boss_{i % 8}", str(rng.randint(1000, 99999)), str(rng.randint(5, 20)),
             str(rng.randint(100, 999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[0] = r[0] + " (어려움)"
        v[2] = str(int(r[2]) * 3)
        return v
    if shape in ("last3", "last2"):
        k = int(shape[-1])
        body[n - k:n - k] = [variant(r) for r in rows[n - k:]]
        idx = list(range(n - k, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = top + sheet(rows)
    left = top + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[2 + i][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    for i in idx:
        row = next(row for row in m if row[6][3] == "7777" and row[2][2] == rows[i][0])
        assert row[2][1] == rows[i][0] and row[6][0] == "right"
    # 설명 행은 그 자리끼리 그대로 짝지어진다
    assert meta[1] == (1, 1, 1) and all(c[0] == "same" for c in m[1])


def test_text_table_first_row_is_not_taken_for_a_description_row():
    """모든 칸이 글자인 문자열 표의 본문 첫 행은 설명 행으로 빼지 않는다(내용으로 짝짓고 '불확실' 표시)."""
    rows = [["ID", "N"]] + [["", f"v{i % 3}"] for i in range(50)]
    left = rows[:20] + [["", "new"]] + rows[20:]
    amb = set()
    m, _meta = compute_diff3(left, rows, rows, 0, 0, None, amb)
    assert len(amb) == len(m) - 1


@pytest.mark.parametrize("heads", [["StageName", "MonsterName"], ["WaveName", "MonsterName"],
                                   ["스테이지 이름", "몬스터 이름"]])
@pytest.mark.parametrize("sizes", [(30,), (60,), (20, 40)])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
def test_spawn_table_with_one_or_two_groups_named_on_the_first_row(heads, sizes, shape):
    """묶음이 1~2개뿐인 '첫 행에만 이름' 표에서 빈칸을 다시 빼고 세어 묶음 열이 고유(표시 이름)로 올라갔다."""
    _spawn_copy_case(heads, _first_row_only(sizes), shape, True, 1)


@pytest.mark.parametrize("member", ["출전 인원(최대 5명)", "출전 인원(3~5명)", "최대 5명", "인원 (최대 5명)"])
@pytest.mark.parametrize("hero", ["HeroName", "영웅명"])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_spawn_table_with_a_member_count_written_after_a_number(member, hero, shape):
    """'최대 5명'·'(3~5명)' 처럼 숫자 뒤에 붙은 인원 단위를 이름으로 봐 숫자 인원 열이 표시 이름이 됐다."""
    test_spawn_table_with_a_member_count_column_in_myeong(member, hero, shape)


# ── 45차: 설명 행이 세 파일에서 조금 다른 2단 머리글, 3단 머리글, 이름을 절반 넘게 안 적은 표 ───────────────

_EN = ["#", "SortKey", "StageName", "BossName", "RecommendPower", "Stamina", "Exp"]
_KO = ["", "정렬", "스테이지 이름", "보스 이름", "추천 전투력", "스태미나", "경험치"]
_TY = ["", "int32", "FName", "FName", "int32", "int32", "int32"]


def _multi_head_case(top, shape, seed, edit=None):
    """머리글 여러 줄(top)을 키 행 0 으로 연 스테이지 표의 '(어려움)' 복사본 끼움 + 다시 매김(right 는 원본 Exp
    7777). edit: 'left_adds_col'·'right_adds_col'·'left_edits_desc'·'right_edits_desc'."""
    rng = random.Random(seed)
    n = 60
    rows = [[f"Stage_{i:03d}", f"Boss_{i % 8}", str(rng.randint(1000, 99999)), str(rng.randint(5, 20)),
             str(rng.randint(100, 999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[0] = r[0] + " (어려움)"
        v[2] = str(int(r[2]) * 3)
        return v
    if shape in ("last3", "last2"):
        k = int(shape[-1])
        body[n - k:n - k] = [variant(r) for r in rows[n - k:]]
        idx = list(range(n - k, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    h = len(top)
    base = [t[:] for t in top] + sheet(rows)
    left = [t[:] for t in top] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[h + i][6] = "7777"
    ko = next((j for j, t in enumerate(top) if t is _KO), None)
    if edit in ("left_adds_col", "right_adds_col"):
        side = left if edit == "left_adds_col" else right
        for j, r in enumerate(side):
            if j < h:
                r.append("Gold" if top[j] is _EN else "골드" if top[j] is _KO else "int32")
            else:
                r.append("0")
    elif edit == "left_edits_desc" and ko is not None:
        left[ko][6] = "획득 경험치"
    elif edit == "right_edits_desc" and ko is not None:
        right[ko][4] = "권장 전투력"
    m, meta = compute_diff3(left, base, right, 0, 0)
    for i in idx:
        row = next(row for row in m if row[6][3] == "7777" and row[2][2] == rows[i][0])
        assert row[2][1] == rows[i][0] and row[6][0] == "right"
    return m


@pytest.mark.parametrize("edit", ["left_adds_col", "right_adds_col", "left_edits_desc", "right_edits_desc"])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
@pytest.mark.parametrize("seed", range(2))
def test_two_row_header_with_a_changed_description_row(edit, shape, seed):
    """설명 행이 세 파일에서 칸까지 똑같을 때만 머리글로 봐, 한쪽이 열을 더하거나 설명 칸을 고치면 키 행 0 으로
    연 2단 머리글 표의 '(어려움)' 복사본 끼움을 다시 원본과 짝지었다."""
    m = _multi_head_case([_EN, _KO], shape, seed, edit)
    if edit == "left_edits_desc":
        assert m[1][6] == ("left", "획득 경험치", "경험치", "경험치")


@pytest.mark.parametrize("order", ["en_ko_ty", "ko_en_ty", "en_ty_ko"])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
@pytest.mark.parametrize("seed", range(2))
def test_three_row_header_opened_with_key_row_0(order, shape, seed):
    """이름·설명·형식(int32) 3단 머리글을 키 행 0 으로 열면 설명 행 처리가 걸리지 않아 복사본 끼움을 원본과
    짝지었다."""
    top = {"en_ko_ty": [_EN, _KO, _TY], "ko_en_ty": [_KO, _EN, _TY], "en_ty_ko": [_EN, _TY, _KO]}[order]
    m = _multi_head_case(top, shape, seed)
    assert all(c[0] == "same" for r in m[1:3] for c in r)


@pytest.mark.parametrize("blank", ["tail30", "tail35", "head35", "mid30"])
@pytest.mark.parametrize("heads", [["StageName", "BossName"], ["스테이지 이름", "보스 이름"]])
def test_stage_table_with_most_stage_names_unwritten(blank, heads):
    """이름을 절반 넘게 아직 안 적은 스테이지 표(한 덩어리 빈칸)를 '묶음 첫 행에만 이름' 표로 봐 StageName 을 묶음
    열로 내렸다(44차 회귀)."""
    rng = random.Random(9)
    n = 60
    names = [f"{i // 10 + 1}-{i % 10 + 1} 숲길_{i:02d}" for i in range(n)]
    empty = {"tail30": range(30, 60), "tail35": range(25, 60), "head35": range(0, 35),
             "mid30": range(10, 40)}[blank]
    for i in empty:
        names[i] = ""
    H = ["#", "SortKey"] + heads + ["RecommendPower", "Stamina", "Exp"]
    rows = [[names[i], f"Boss_{i % 8}", str(rng.randint(1000, 9999)), str(rng.randint(5, 20)),
             str(rng.randint(100, 999))] for i in range(n)]
    target = next(i for i in (5, 45, 50, 15) if names[i] and i + 1 < n and names[i + 1])
    body = [r[:] for r in rows]
    v = rows[target][:]
    v[0] = v[0] + " (어려움)"
    v[2] = str(int(v[2]) * 3)
    body.insert(target, v)
    del body[target + 2]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [H] + sheet(rows)
    left = [H] + sheet(body)
    right = [r[:] for r in base]
    right[target + 1][6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    row = next(row for row in m if row[6][3] == "7777")
    assert row[2][1] == row[2][2] == rows[target][0] and row[6][0] == "right"


# ── 46차: 설명 행 묶음은 키 칸이 같은 앞쪽 줄까지만, 주석 행은 똑같을 때만 ─────────────────────────

def _chapter_table(first, n_per=22):
    """1단 머리글 고유 StageID 표 — 장마다 첫 행이 A열 '// N장' 주석."""
    rows = [["StageID", "Name", "Exp"]]
    for ch in first:
        rows.append([f"// {ch}장", "", ""])
        rows += [[str(ch * 1000 + i), f"S{ch}-{i}", str(100 + i)] for i in range(1, n_per + 1)]
    return rows


def test_new_chapter_inserted_on_top_does_not_relabel_the_first_chapter_comment():
    """1단 머리글 표의 첫 본문 행이 '// 1장' 이면 그 행을 자리끼리 짝지어, 맨 위에 새 장을 넣으면 '// 1장' 을
    '// 9장' 으로 바꿔 쓴 것으로 봤다."""
    base = _chapter_table([1, 2])
    left = _chapter_table([9, 1, 2])
    right = [r[:] for r in base]
    right[5][2] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert not any(r[0][1] == "// 9장" and r[0][2] == "// 1장" for r in m)
    row = next(r for r in m if r[0][2] == "// 1장")
    assert row[0][0] == "same"
    assert next(r for r in m if r[2][3] == "7777")[0][2] == "1004"


def test_swapped_chapters_do_not_rewrite_the_comment_rows():
    """장 순서만 바꾼 쪽은 주석 행 라벨을 바꾼 것으로 보이지 않는다."""
    base = _chapter_table([1, 2])
    left = _chapter_table([2, 1])
    m, meta = compute_diff3(left, base, base, 0, 0)
    for r in m:
        if r[0][2].startswith("//"):
            assert r[0][1] == r[0][2]


@pytest.mark.parametrize("change", ["swap", "replace"])
def test_constant_table_top_rows_are_paired_by_key(change):
    """맨 위 1~2행이 글자만 있는 상수 표에서 그 행을 자리끼리 짝지어, 한쪽이 순서를 바꾸거나 다른 키로 바꿔 넣으면
    다른 쪽의 설명 수정이 엉뚱한 키 행에 병합됐다."""
    base = [["Key", "Value", "Desc"],
            ["StartHeroID", "Hero_A", "영웅"],
            ["StartMapID", "Map_A", "맵"]] + [[f"Const{i}", str(10 + i), f"설명{i}"] for i in range(25)]
    left = [r[:] for r in base]
    if change == "swap":
        left[1], left[2] = left[2], left[1]
    else:
        left[1] = ["StartGachaID", "Gacha_A", "뽑기"]
    right = [r[:] for r in base]
    right[1][2] = "시작 영웅"
    m, meta = compute_diff3(left, base, right, 0, 0)
    hits = [r for r in m if r[2][3] == "시작 영웅"]
    assert len(hits) == 1
    assert hits[0][0][2] == "StartHeroID" and hits[0][0][3] == "StartHeroID"
    assert hits[0][0][1] in ("StartHeroID", "")


@pytest.mark.parametrize("who", ["left", "right"])
@pytest.mark.parametrize("shape", ["last3", "last2", "mid_copy_del_next"])
def test_two_row_header_when_one_side_removes_the_top_chapter_comment(who, shape):
    """2단 머리글 아래 '// 1장' 주석 행을 한쪽이 지우면 설명 행 묶음 줄 수가 달라 설명 행 처리가 꺼지고, 복사본
    끼움을 원본과 짝지었다."""
    rng = random.Random(3)
    n = 60
    rows = [[f"Stage_{i:03d}", f"Boss_{i % 8}", str(rng.randint(1000, 99999)), str(rng.randint(5, 20)),
             str(rng.randint(100, 999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[0] = r[0] + " (어려움)"
        v[2] = str(int(r[2]) * 3)
        return v
    if shape in ("last3", "last2"):
        k = int(shape[-1])
        body[n - k:n - k] = [variant(r) for r in rows[n - k:]]
        idx = list(range(n - k, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs, top_comment=True):
        out = [["// 1장", "", "", "", "", "", ""]] if top_comment else []
        for i, r in enumerate(rs):
            if i and i % 20 == 0:
                out.append([f"// {i // 20 + 1}장", "", "", "", "", "", ""])
            out.append(["", str(i + 1)] + r)
        return out
    base = [_EN, _KO] + sheet(rows)
    left = [_EN, _KO] + sheet(body, who != "left")
    right = [_EN, _KO] + sheet(rows, who != "right")
    for r in right:
        if r[2] in [rows[i][0] for i in idx]:
            r[6] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    for i in idx:
        row = next(row for row in m if row[6][3] == "7777" and row[2][2] == rows[i][0])
        assert row[2][1] == rows[i][0] and row[6][0] == "right"


# ── 47차: 자리 짝짓는 설명 행은 칸을 넓게 채우고 키 아닌 칸이 절반 넘게 같은 행만 ─────────────────────

def _chapter_blank_key(chapters, n_per=25):
    rows = [["#", "SortKey", "StageName", "Boss", "Power", "Stamina", "Exp"]]
    titles = {0: "프롤로그", 1: "숲의 입구", 2: "동굴"}
    k = 1
    for ch in chapters:
        rows.append(["", f"{ch}장", titles[ch], "", "", "", ""])
        for i in range(n_per):
            rows.append(["", str(k), f"{ch}-{i + 1} {titles[ch]}_{i}", f"Boss_{i % 4}", str(1000 + i), "5",
                         str(100 + i)])
            k += 1
    return rows


@pytest.mark.parametrize("change", ["insert_top", "swap"])
def test_sparse_chapter_rows_are_not_paired_by_position(change):
    """키가 빈 표의 장 구분 행('1장' + 제목, 칸 2개)을 자리끼리 짝지어, 맨 위에 새 장을 넣거나 장 순서를 바꾸면
    장 라벨을 바꿔 쓴 것으로 자동 병합했다."""
    base = _chapter_blank_key([1, 2])
    left = _chapter_blank_key([0, 1, 2] if change == "insert_top" else [2, 1])
    m, meta = compute_diff3(left, base, base, 0, 0)
    assert not any(r[1][2] == "1장" and r[1][1] not in ("1장", "") for r in m)


@pytest.mark.parametrize("who", ["left", "right"])
def test_key_column_description_cell_edit_keeps_the_row_in_place(who):
    """2단 머리글 고유 키 표에서 키 열의 설명 칸('아이디')만 고치면 설명 행이 지움 + 새 행이 되어 시트 끝으로
    옮겨졌다."""
    rng = random.Random(1)
    base = [["ID", "SortKey", "Name", "Power", "Exp"], ["아이디", "정렬", "이름", "전투력", "경험치"]] + \
        [[str(1000 + i), str(i + 1), f"Mon_{i}", str(rng.randint(1000, 9999)), str(rng.randint(100, 999))]
         for i in range(40)]
    left = [r[:] for r in base]
    right = [r[:] for r in base]
    (left if who == "left" else right)[1][0] = "고유 ID"
    (right if who == "left" else left)[10][4] = "7777"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert meta[1] == (1, 1, 1)
    want = ("left", "고유 ID", "아이디", "아이디") if who == "left" else ("right", "아이디", "아이디", "고유 ID")
    assert m[1][0] == want
    assert len(m) == len(base)


@pytest.mark.parametrize("edit", ["right_writes", "left_writes", "both_write"])
@pytest.mark.parametrize("shape", ["last3", "mid_copy_del_next"])
def test_description_row_key_cell_written_on_one_side(edit, shape):
    """설명 행의 키 열(#) 칸에 한쪽만 글을 적으면 설명 행 처리가 꺼져 복사본 끼움을 원본과 짝지었다."""
    rng = random.Random(2)
    n = 60
    rows = [[f"Stage_{i:03d}", f"Boss_{i % 8}", str(rng.randint(1000, 99999)), str(rng.randint(5, 20)),
             str(rng.randint(100, 999))] for i in range(n)]
    body = [r[:] for r in rows]

    def variant(r):
        v = r[:]
        v[0] = r[0] + " (어려움)"
        v[2] = str(int(r[2]) * 3)
        return v
    if shape == "last3":
        body[n - 3:n - 3] = [variant(r) for r in rows[n - 3:]]
        idx = list(range(n - 3, n))
    else:
        body.insert(30, variant(rows[30]))
        del body[32]
        idx = [30]

    def sheet(rs):
        return [["", str(i + 1)] + r for i, r in enumerate(rs)]
    base = [_EN[:], _KO[:]] + sheet(rows)
    left = [_EN[:], _KO[:]] + sheet(body)
    right = [r[:] for r in base]
    for i in idx:
        right[2 + i][6] = "7777"
    if edit in ("right_writes", "both_write"):
        right[1][0] = "비고"
    if edit in ("left_writes", "both_write"):
        left[1][0] = "메모"
    m, meta = compute_diff3(left, base, right, 0, 0)
    assert meta[1] == (1, 1, 1)
    for i in idx:
        row = next(row for row in m if row[6][3] == "7777" and row[2][2] == rows[i][0])
        assert row[2][1] == rows[i][0] and row[6][0] == "right"


@pytest.mark.parametrize("change", ["insert_top", "swap"])
def test_text_only_group_top_rows_are_not_paired_by_position(change):
    """키가 빈(또는 겹친) 묶음의 맨 위 글자 행(영웅 보상)을 자리끼리 짝지어, 맨 위에 행을 끼우거나 순서를 바꾸면
    다른 쪽 수정이 엉뚱한 행에 표시 없이 자동 병합됐다."""
    H = ["#", "Name", "Type", "Memo"]
    top = [["", "Rudy", "Hero", ""], ["", "Eileen", "Hero", ""]]
    rest = [["", str(1001 + i), "Item", ""] for i in range(25)]
    base = [H] + [r[:] for r in top] + [r[:] for r in rest]
    left = [r[:] for r in base]
    if change == "insert_top":
        left.insert(1, ["", "Kris", "Hero", ""])
    else:
        left[1], left[2] = left[2], left[1]
    right = [r[:] for r in base]
    right[2][2] = "HeroPiece"                   # Eileen
    amb = set()
    m, meta = compute_diff3(left, base, right, 0, 0, None, amb)
    hits = [i for i, r in enumerate(m) if r[2][3] == "HeroPiece"]
    assert len(hits) == 1
    r = m[hits[0]]
    assert r[1][2] == "Eileen" and r[1][3] == "Eileen" and r[1][1] in ("Eileen", "")


# ── 48차: 자리로 짝지은 다른 줄은 '불확실'(자동 병합 제외), 본문 값·옮긴 흔적이면 끊음 ───────────────────

@pytest.mark.parametrize("change", ["insert_top", "swap"])
@pytest.mark.parametrize("extra", [["Grade"], ["Element", "Grade"]])
def test_reward_group_top_rows_with_shared_columns(change, extra):
    """보상 묶음 맨 위 글자 행들이 이름만 다르고 Type·Grade 가 같으면 47차 '절반 넘게 같음'을 통과해 자리로
    짝지어져, 다른 쪽 수정이 엉뚱한 행에 표시 없이 자동 병합되고 행 하나가 사라졌다."""
    H = ["#", "Name", "Type"] + extra + ["Memo"]
    pad = ["Normal"] * len(extra) if extra == ["Grade"] else ["Fire", "Normal"]
    top = [["", "Rudy", "Hero"] + pad + [""], ["", "Eileen", "Hero"] + pad + [""]]
    rest = [["", str(1001 + i), "Item"] + pad + [""] for i in range(25)]
    base = [H] + [r[:] for r in top] + [r[:] for r in rest]
    left = [r[:] for r in base]
    if change == "insert_top":
        left.insert(1, ["", "Kris", "Hero"] + pad + [""])
    else:
        left[1], left[2] = left[2], left[1]
    right = [r[:] for r in base]
    right[2][2] = "HeroPiece"                   # Eileen
    amb = set()
    m, meta = compute_diff3(left, base, right, 0, 0, None, amb)
    plan = auto_merge_plan(m, meta, amb)
    for (r, c), _side in plan.items():
        # 자동 병합 계획에 엉뚱한 행의 이름·종류 변경이 들어가지 않는다
        assert not (m[r][1][2] in ("Rudy", "Eileen") and m[r][1][0] != "same")
    hit = [i for i, r in enumerate(m) if r[2][3] == "HeroPiece"]
    assert len(hit) == 1 and m[hit[0]][1][2] == "Eileen"


@pytest.mark.parametrize("change", ["swap", "insert_top", "replace"])
def test_constant_table_with_shared_columns_is_paired_by_key(change):
    """고유 키 상수 표의 맨 위 글자 행이 Type·Category 를 같이 쓰면 47차 규칙이 자리로 짝지어, 맞바꿈·끼움·바꿔
    넣기에서 다른 쪽 설명 수정이 다른 키 행에 표시 없이 자동 병합됐다."""
    H = ["Key", "Type", "Value", "Category", "Desc"]
    base = [H, ["StartHeroID", "String", "Hero_A", "Start", ""],
            ["StartMapID", "String", "Map_A", "Start", ""]] + \
        [[f"Const{i}", "Int", str(10 + i), "Misc", ""] for i in range(25)]
    left = [r[:] for r in base]
    if change == "swap":
        left[1], left[2] = left[2], left[1]
    elif change == "insert_top":
        left.insert(1, ["StartGachaID", "String", "Gacha_A", "Start", ""])
    else:
        left[1] = ["StartGachaID", "String", "Gacha_A", "Start", ""]
    right = [r[:] for r in base]
    right[1][4] = "HERO"
    amb = set()
    m, meta = compute_diff3(left, base, right, 0, 0, None, amb)
    plan = auto_merge_plan(m, meta, amb)
    for (r, c), _side in plan.items():
        if c == 4:
            assert m[r][0][2] == "StartHeroID"
    hit = [r for r in m if r[4][3] == "HERO"]
    assert len(hit) == 1 and hit[0][0][2] == "StartHeroID"
    assert hit[0][0][1] in ("StartHeroID", "")


def test_changed_description_row_is_not_auto_merged_silently():
    """세 파일에서 다른 설명 행은 자리로 짝짓되 '불확실'로 표시해 자동 병합에서 뺀다."""
    base = [_EN[:], _KO[:]] + [["", str(i + 1), f"Stage_{i:03d}", f"Boss_{i % 8}", str(1000 + i), "5",
                                str(100 + i)] for i in range(30)]
    left = [r[:] for r in base]
    left[1][6] = "획득 경험치"
    amb = set()
    m, meta = compute_diff3(left, base, base, 0, 0, None, amb)
    assert meta[1] == (1, 1, 1) and m[1][6] == ("left", "획득 경험치", "경험치", "경험치")
    assert 1 in amb
