# -*- coding: utf-8 -*-
"""같은 화면에서 두 번 저장해도 **맞는 행**에 써야 한다.

저장 뒤 화면 값만 '같음'으로 고치고 행 대응(row_meta)·원본 데이터를 그대로 두던 때는,
두 번째 저장이 옛 행 번호로 썼다. 실측: B 에만 있는 ID=2 행을 지우는 저장을 한 뒤
ID=3 의 V 를 저장하자, 이미 없어진 4행 자리에 ['', 'a3-new'] 행이 새로 생겼다.
저장 검증은 '계획한 패치가 들어갔나' 만 봐서 그대로 통과했다.

이제 저장 뒤 파일을 다시 읽어 다시 비교하고, 병합 표시(파랑)와 다른 쪽 병합 준비는
키·열로 다시 찾아 되살린다. 전부 실제 화면·워커·파일 경로로 돈다(화면 없이).
"""
import os
import time

import openpyxl
import pytest
from PyQt5.QtWidgets import QApplication

from excelmerge.constants import DIR_A2B, DIR_B2A, STATUS_SAME
from excelmerge.loaders import clear_values_cache, load_values_any


def _xlsx(path, rows):
    wb = openpyxl.Workbook()
    for r in rows:
        wb.active.append(r)
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


def _values(path):
    clear_values_cache()
    return load_values_any(path)


@pytest.fixture
def opened(qapp, tmp_path, monkeypatch):
    """A: ID=2 가 없고 ID=3 값이 다르다. B: ID=2 가 있다."""
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information", staticmethod(lambda *a, **k: None))
    a = _xlsx(tmp_path / "a.xlsx", [["ID", "V"], ["1", "a1"], ["3", "a3-new"], ["5", "a5"]])
    b = _xlsx(tmp_path / "b.xlsx", [["ID", "V"], ["1", "a1"], ["2", "b2"], ["3", "b3"], ["5", "b5"]])
    win = MainWindow()
    win.show()
    view = win.open_file_compare(a, b)
    assert _wait(lambda: bool(view._diff_matrix) and view.status.currentMessage().startswith("비교 완료"))
    yield view, a, b
    win.close(); win.deleteLater(); qapp.processEvents()


def _row_of(view, key):
    for r, row in enumerate(view._diff_matrix):
        if row[0][1] == key or row[0][2] == key:
            return r
    raise AssertionError(f"{key} 행이 없다")


def _save(view, side):
    view._set_save_btn_state()
    view._save_staged(side)
    # 저장 워커가 끝나고, 다시 비교가 있다면 그것까지 끝날 때까지. (다시 비교가 없던 예전
    # 코드에서도 이 조건은 '저장 완료' 에서 풀린다 — 그래야 버그 자체로 실패하는 걸 본다.)
    assert _wait(lambda: not view._saving_now()
                 and getattr(view, "_post_save_carry", None) is None
                 and view.status.currentMessage().startswith(("비교 완료", "저장 완료"))),         view.status.currentMessage()


def test_a_second_save_writes_to_the_right_row(opened):
    view, a, b = opened
    r2 = _row_of(view, "2")
    view._staged[(r2, 0)] = DIR_A2B          # B 에만 있는 행 → B 에서 지운다
    view._staged[(r2, 1)] = DIR_A2B
    _save(view, "b")
    assert _values(b) == [["ID", "V"], ["1", "a1"], ["3", "b3"], ["5", "b5"]]

    r3 = _row_of(view, "3")                  # 같은 화면에서 한 번 더
    view._staged[(r3, 1)] = DIR_A2B
    _save(view, "b")
    assert _values(b) == [["ID", "V"], ["1", "a1"], ["3", "a3-new"], ["5", "b5"]], \
        "두 번째 저장이 옛 행 번호로 엉뚱한 행에 썼다"


def test_the_merged_mark_survives_the_reload(opened):
    view, a, b = opened
    r3 = _row_of(view, "3")
    view._staged[(r3, 1)] = DIR_A2B
    _save(view, "b")
    r3 = _row_of(view, "3")
    assert (r3, 1) in view._merged_cells, "저장한 칸의 병합 표시(파랑)가 사라졌다"
    assert view._diff_matrix[r3][1][0] == STATUS_SAME


def test_the_other_sides_staged_cells_survive(opened):
    """B 를 저장해도 A 쪽으로 준비해 둔 병합은 남아야 한다 — 다시 비교하며 날리면 안 된다."""
    view, a, b = opened
    r3, r5 = _row_of(view, "3"), _row_of(view, "5")
    view._staged[(r3, 1)] = DIR_A2B          # 이번에 저장할 쪽(B)
    view._staged[(r5, 1)] = DIR_B2A          # 남겨 둘 쪽(A)
    _save(view, "b")
    r5 = _row_of(view, "5")
    assert view._staged == {(r5, 1): DIR_B2A}
    assert "다시 비교함" in view.status.currentMessage()
    _save(view, "a")                         # 되살린 준비로 A 를 저장해도 맞는 행에
    assert _values(a) == [["ID", "V"], ["1", "a1"], ["3", "a3-new"], ["5", "b5"]]


def test_the_filter_state_is_kept(opened):
    """다시 비교가 사용자가 꺼 둔 '변경점만 보기' 를 켜 버리면 안 된다."""
    view, a, b = opened
    view.diff_only_btn.setChecked(False)
    r3 = _row_of(view, "3")
    view._staged[(r3, 1)] = DIR_A2B
    _save(view, "b")
    assert not view.diff_only_btn.isChecked()


def test_row_order_mode_drops_staged_cells_and_says_so(opened):
    """키 없는 행 순서 모드는 저장으로 행이 밀릴 수 있다 — 엉뚱한 칸에 다시 걸지 않고 알린다."""
    view, a, b = opened
    view._key_col = -1
    view._recompute_diff()
    assert _wait(lambda: view.status.currentMessage().startswith("키 헤더"))
    view._staged[(1, 1)] = DIR_A2B           # 저장할 쪽
    view._staged[(2, 1)] = DIR_B2A           # 남겨 두려던 쪽
    _save(view, "b")
    assert view._staged == {}
    assert "1개는 다시 찾지 못해 해제됨" in view.status.currentMessage()


def test_new_rows_are_appended_in_screen_order():
    """새 행은 화면 순서대로 붙어야 한다 — 집합 순서로 돌면 k5000, k4097, k10, k3 순으로 붙었다."""
    from excelmerge.merge_build import build_side_patches
    shown = [3, 10, 4097, 5000]
    n = max(shown) + 1
    matrix = [[("same", "", ""), ("same", "", "")] for _ in range(n)]
    meta = [(i, i) for i in range(n)]
    staged = {}
    for r in shown:
        matrix[r] = [("added", f"k{r}", ""), ("added", f"v{r}", "")]
        meta[r] = (r, None)                  # A 에만 있는 행 → B 에 새로 넣는다
        staged[(r, 0)] = DIR_A2B
        staged[(r, 1)] = DIR_A2B
    _p, inserts, _s = build_side_patches(DIR_A2B, matrix, meta, staged, None)
    assert list(inserts.keys()) == shown
