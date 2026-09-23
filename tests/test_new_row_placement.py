# -*- coding: utf-8 -*-
"""신규(B 전용) 행은 B 에서의 제자리에 끼워 넣는다.

예전에는 A 에 없는 키를 전부 표 맨 아래에 몰아 붙였다. B 중간에 새로 넣은 행이
이웃과 떨어져 맨 끝에 나타나니, '어디에 추가됐는지'가 표에서 사라졌다.
이제 'B 에서 그 행 다음에 나오는 공통 키 바로 앞'에 넣는다(고전적 diff 정렬).
"""
from excelmerge.diff_engine import compute_diff

H = ["ID", "V"]


def _keys(dm):
    """각 행의 (A쪽 키, B쪽 키) — 빈 문자열은 그쪽 파일에 행이 없다는 뜻."""
    return [(r[0][1], r[0][2]) for r in dm[1:]]     # 0행은 헤더


def _run(a_keys, b_keys):
    a = [H] + [[k, "a"] for k in a_keys]
    b = [H] + [[k, "b" if k.startswith("NEW") else "a"] for k in b_keys]
    dm, meta = compute_diff(a, b, key_col=0, key_row=0)
    return _keys(dm), meta[1:]


def test_new_row_lands_between_its_b_neighbours():
    """보고된 요청 — B 중간에 추가된 행이 그 자리에 보인다."""
    keys, _ = _run(["k1", "k2", "k3"], ["k1", "k2", "NEW", "k3"])
    assert keys == [("k1", "k1"), ("k2", "k2"), ("", "NEW"), ("k3", "k3")]


def test_consecutive_new_rows_keep_their_b_order():
    keys, _ = _run(["k1", "k2"], ["k1", "NEW1", "NEW2", "NEW3", "k2"])
    assert keys == [("k1", "k1"), ("", "NEW1"), ("", "NEW2"), ("", "NEW3"),
                    ("k2", "k2")]


def test_new_row_at_the_top_of_b_goes_to_the_top():
    keys, _ = _run(["k1", "k2"], ["NEW", "k1", "k2"])
    assert keys[0] == ("", "NEW")


def test_new_row_at_the_end_of_b_still_goes_last():
    """B 뒤에 공통 키가 없으면 붙일 자리가 '맨 끝'뿐이다 — 예전 동작 그대로."""
    keys, _ = _run(["k1", "k2"], ["k1", "k2", "NEW"])
    assert keys == [("k1", "k1"), ("k2", "k2"), ("", "NEW")]


def test_deleted_row_comes_before_the_new_one_at_the_same_spot():
    """A 에만 있는 행(삭제)과 자리가 겹치면 삭제 → 신규 순서(diff 관례)."""
    keys, _ = _run(["k1", "DEL", "k3"], ["k1", "NEW", "k3"])
    assert keys == [("k1", "k1"), ("DEL", ""), ("", "NEW"), ("k3", "k3")]


def test_a_order_stays_the_backbone_when_common_keys_are_reordered():
    """공통 키가 B 에서 뒤섞여 있어도 표의 뼈대는 A 순서다."""
    keys, _ = _run(["k1", "k2", "k3"], ["k3", "NEW", "k1", "k2"])
    assert [a for a, _b in keys if a] == ["k1", "k2", "k3"]
    assert ("", "NEW") in keys


def test_row_meta_points_at_the_original_file_rows():
    """병합 저장 좌표 — 끼워 넣어도 원본 행 번호가 정확해야 한다."""
    _keys_, meta = _run(["k1", "k2", "k3"], ["k1", "k2", "NEW", "k3"])
    #               k1      k2      NEW        k3(B에선 4행)
    assert meta == [(1, 1), (2, 2), (None, 3), (3, 4)]


def test_only_new_rows_at_all():
    keys, _ = _run([], ["NEW1", "NEW2"])
    assert keys == [("", "NEW1"), ("", "NEW2")]


def test_new_rows_survive_the_preamble_header_split():
    """key_row 가 1이면 0~1행은 위치 1:1, 그 아래 본문에만 끼워 넣기가 적용된다."""
    a = [["TITLE", ""], ["ID", "V"], ["k1", "a"], ["k2", "a"]]
    b = [["TITLE", ""], ["ID", "V"], ["k1", "a"], ["NEW", "b"], ["k2", "a"]]
    dm, meta = compute_diff(a, b, key_col=0, key_row=1)
    assert [(r[0][1], r[0][2]) for r in dm] == [
        ("TITLE", "TITLE"), ("ID", "ID"), ("k1", "k1"), ("", "NEW"), ("k2", "k2")]
    assert meta == [(0, 0), (1, 1), (2, 2), (None, 3), (3, 4)]


def test_row_order_mode_is_untouched():
    """키 열 해제(ROW 순서 비교)는 예전 그대로 — 행 번호 1:1."""
    a = [H, ["k1", "a"], ["k2", "a"]]
    b = [H, ["k1", "a"], ["NEW", "b"], ["k2", "a"]]
    dm, meta = compute_diff(a, b, key_col=-1)
    assert [(r[0][1], r[0][2]) for r in dm] == [
        ("ID", "ID"), ("k1", "k1"), ("k2", "NEW"), ("", "k2")]
    assert meta == [(0, 0), (1, 1), (2, 2), (None, 3)]
