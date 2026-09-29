# -*- coding: utf-8 -*-
"""비교에서 빠진 행을 보이게 한다 — 원인 진단 + 쓸 수 있는 키 열 추천.

키 열에 중복·빈 값이 있으면 그 행은 비교에서 빠진다(compute_diff 는 같은 키의 첫 행만
쓴다). 지금까지는 상태바 한 줄로만 알려서, 14,659행이 빠진 비교가 '차이 없음'처럼
보일 수 있었다.

실측하면 현장 데이터의 원인은 거의 전부 **키 열을 잘못 잡은 것**이라, 이 기능의 핵심은
목록이 아니라 '이 열을 키로 바꾸기'다.
"""
import pytest

from excelmerge.diff_engine import (count_dropped_key_rows, dropped_key_rows,
                                    unique_key_candidates)

HEADER = ["ID", "GRP", "NAME"]
#           키가 될 A열은 유니크, B열(GRP)은 중복투성이 — 실제 데이터와 같은 모양
A_DATA = [HEADER] + [[f"k{i}", "10", f"a{i}"] for i in range(1, 6)]
B_DATA = [HEADER] + [[f"k{i}", "10", f"b{i}"] for i in range(1, 6)]


# ── 진단 데이터 ──────────────────────────────────────────────────────────────

def test_counts_match_the_existing_counter():
    """기존 카운터(count_dropped_key_rows)와 합계가 어긋나면 안 된다."""
    info = dropped_key_rows(A_DATA, B_DATA, key_col=1)
    assert info["blank"] + info["dup"] == count_dropped_key_rows(A_DATA, B_DATA, 1)


def test_duplicate_rows_are_reported_per_side():
    info = dropped_key_rows(A_DATA, B_DATA, key_col=1)
    assert info["dup"] == 8 and info["blank"] == 0      # 각 파일 5행 중 4행이 중복
    assert info["per_side"] == {"A": 4, "B": 4}


def test_blank_keys_are_counted_separately():
    a = [HEADER, ["k1", "", "x"], ["k2", "", "y"], ["k3", "g", "z"]]
    info = dropped_key_rows(a, [], key_col=1)
    assert info["blank"] == 2 and info["dup"] == 0
    assert all(reason == "빈 키" for _s, _r, _k, reason in info["rows"])


def test_rows_carry_original_row_numbers():
    info = dropped_key_rows(A_DATA, [], key_col=1)
    side, row_no, key, reason = info["rows"][0]
    assert side == "A" and key == "10" and reason == "중복 키"
    assert row_no == 3, f"원본 행 번호(1-based)여야 한다: {row_no}"


def test_top_duplicate_keys_are_summarised():
    a = [HEADER] + [["k", "0", "x"] for _ in range(5)] + [["k", "9", "y"]] * 2
    info = dropped_key_rows(a, [], key_col=1)
    keys = dict((k, na) for k, na, _nb in info["dup_keys"])
    assert keys["0"] == 4 and keys["9"] == 1
    assert info["dup_keys"][0][0] == "0", "많이 겹친 키가 앞에 와야 한다"


def test_row_list_is_capped_but_counts_are_not():
    a = [HEADER] + [["k", "0", "x"] for _ in range(300)]
    info = dropped_key_rows(a, [], key_col=1, limit=10)
    assert len(info["rows"]) == 10
    assert info["dup"] == 299, "목록은 잘라도 집계는 전부 세야 한다"


def test_row_order_mode_drops_nothing():
    assert dropped_key_rows(A_DATA, B_DATA, key_col=-1)["dup"] == 0


# ── 키 후보 ──────────────────────────────────────────────────────────────────

def test_candidate_must_be_unique_in_both_files():
    """한쪽만 유니크한 열을 고르면 반대쪽에서 다시 행이 빠진다."""
    a = [HEADER, ["1", "x", "n"], ["2", "y", "n"]]
    b = [HEADER, ["1", "x", "n"], ["1", "z", "n"]]      # B 의 A열은 중복
    assert 0 not in [c for c, _h in unique_key_candidates(a, b)]


def test_candidate_rejects_columns_with_blanks():
    a = [HEADER, ["1", "", "n"], ["2", "g", "n"]]
    assert 1 not in [c for c, _h in unique_key_candidates(a, a)]


def test_candidate_returns_header_text():
    cands = unique_key_candidates(A_DATA, B_DATA)
    assert (0, "ID") in cands, cands


def test_candidate_respects_the_header_row():
    """key_row 아래(본문)만 본다 — 헤더 자체는 유니크 판정에 넣지 않는다."""
    a = [["title", "", ""], HEADER, ["1", "g", "n"], ["2", "g", "n"]]
    cands = [c for c, _h in unique_key_candidates(a, a, key_row=1)]
    assert 0 in cands and 1 not in cands


# ── 화면 연결 ────────────────────────────────────────────────────────────────

@pytest.fixture
def dv(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    monkeypatch.setenv("APPDATA", str(tmp_path))
    win = MainWindow()
    win.show()
    try:
        view = win.tabs.currentWidget()
        view.panel_a.set_path("a.xlsx")
        view.panel_b.set_path("b.xlsx")
        view._key_col = 1                      # 중복투성이 열을 키로 — 행이 빠진다
        for p in (view.panel_a, view.panel_b):
            p.table.set_key_col(1)
        view._on_loaded(A_DATA, B_DATA)
        w = getattr(view, "_diff_worker", None)
        if w is not None:
            w.wait(8000)
        for _ in range(20):
            qapp.processEvents()
        yield view
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def test_warning_appears_when_rows_are_dropped(dv):
    assert dv._dropped_count > 0, "전제: 빠진 행이 있어야 한다"
    assert dv.status._warn.isVisible()
    assert "제외됨" in dv.status._warn.text()


def test_warning_disappears_after_switching_to_a_good_key(dv, qapp):
    dv._on_key_col_changed(0)                  # 유니크한 A열로 교체
    w = getattr(dv, "_diff_worker", None)
    if w is not None:
        w.wait(8000)
    for _ in range(20):
        qapp.processEvents()
    assert dv._dropped_count == 0, "좋은 키로 바꿨는데도 빠진 행이 있다"
    assert not dv.status._warn.isVisible(), "경고가 남아 있다"


def test_clicking_the_warning_opens_the_dialog_for_the_visible_tab(dv, monkeypatch):
    opened = []
    monkeypatch.setattr(dv, "_open_dropped_dialog", lambda: opened.append(1))
    dv.status.warning_clicked.emit()
    assert opened == [1]


def test_hidden_tab_ignores_the_shared_warning(dv, monkeypatch, qapp):
    """상태바는 모든 탭이 공유한다 — 화면에 없는 탭이 창을 띄우면 안 된다."""
    opened = []
    monkeypatch.setattr(dv, "_open_dropped_dialog", lambda: opened.append(1))
    dv.hide()
    qapp.processEvents()
    dv.status.warning_clicked.emit()
    assert opened == []


def test_dialog_offers_the_unique_column_and_applies_it(dv, qapp, monkeypatch):
    """창의 '이 열을 키로 바꾸기' 는 기존 키 변경 경로를 그대로 탄다."""
    from excelmerge.dropped_rows_dialog import DroppedRowsDialog
    shown = {}

    def fake_exec(self):
        shown["dlg"] = self
        return 0

    monkeypatch.setattr(DroppedRowsDialog, "exec_", fake_exec)
    dv._open_dropped_dialog()
    dlg = shown["dlg"]
    picked = []
    dlg._on_pick_key = picked.append
    dlg._pick(0)
    assert picked == [0], "A열로 바꾸라는 요청이 전달되지 않았다"


def test_status_warning_does_not_change_the_bar_height(qapp):
    """경고가 나타났다 사라질 때 상태바 높이가 변하면 그 아래 표가 들썩인다."""
    from excelmerge.statusbar import StatusBar
    sb = StatusBar()
    sb.resize(600, sb.sizeHint().height())
    sb.show()
    qapp.processEvents()
    before = sb.sizeHint().height()
    sb.show_warning("⚠ 1,234행 제외됨 — 자세히")
    qapp.processEvents()
    assert sb.sizeHint().height() == before, "경고 때문에 상태바 높이가 변했다"
    sb.clear_warning()
    qapp.processEvents()
    assert sb.sizeHint().height() == before
    sb.deleteLater()
