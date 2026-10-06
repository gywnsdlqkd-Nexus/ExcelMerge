# -*- coding: utf-8 -*-
"""열을 헤더 이름으로 맞춘다 — 1단계: col_meta 계산과 compute_diff 연결.

행은 키로 맞추면서 열은 번호로만 비교했다. 그래서 누가 B 중간에 열 하나를 끼우면 그
오른쪽 전부가 '변경' 으로 뜨고, 그 상태로 병합하면 A 의 값이 B 의 **엉뚱한 열**에
들어간다. 저장 검증도 통과한다 — 의도한 자리에 의도한 값을 썼으니까.

여기서 고정하는 것 셋:
  1. 맞출 수 있으면 맞춘다 — 그리고 표시 순서는 행 규칙(A 가 뼈대, B 전용은 제자리)과 같다.
  2. **맞출 수 없으면 손대지 않는다** — 애매하면 예전 동작(위치 기준)을 유지한다.
  3. col_meta 를 주지 않으면 결과가 예전과 **완전히 같다**(기존 비교가 안 바뀐다).
"""
import pytest

from excelmerge.constants import STATUS_ADDED, STATUS_SAME
from excelmerge.diff_engine import compute_diff, match_columns

A = [["ID", "NAME", "VALUE"],
     ["k1", "칼", "10"],
     ["k2", "방패", "20"]]
# B 중간에 GRADE 가 끼었다 — 기획자가 필드 하나 추가한 흔한 모양
B = [["ID", "GRADE", "NAME", "VALUE"],
     ["k1", "A", "칼", "10"],
     ["k2", "S", "방패", "20"]]


# ── 맞추기 ───────────────────────────────────────────────────────────────────

def test_inserted_column_is_placed_where_b_has_it():
    """B 전용 열은 B 에서의 제자리 — 신규 행을 B 위치에 끼우는 규칙과 같다."""
    assert match_columns(A, B) == [(0, 0), (None, 1), (1, 2), (2, 3)]


def test_a_only_column_is_kept_with_no_b_side():
    a = [["ID", "MEMO", "VALUE"], ["k1", "메모", "10"]]
    b = [["ID", "VALUE"], ["k1", "10"]]
    assert match_columns(a, b) == [(0, 0), (1, None), (2, 1)]


def test_reordered_columns_follow_a():
    """공통 열의 순서가 다르면 A 순서를 따른다(뼈대가 A)."""
    a = [["ID", "NAME", "VALUE"]]
    b = [["VALUE", "ID", "NAME"]]
    assert match_columns(a, b) == [(0, 1), (1, 2), (2, 0)]


def test_identical_headers_give_the_identity_mapping():
    a = [["ID", "NAME"], ["k1", "x"]]
    assert match_columns(a, a) == [(0, 0), (1, 1)]


def test_b_only_columns_with_no_common_after_them_go_last():
    a = [["ID", "NAME"]]
    b = [["ID", "NAME", "NEW1", "NEW2"]]
    assert match_columns(a, b) == [(0, 0), (1, 1), (None, 2), (None, 3)]


def test_the_header_row_is_respected():
    a = [["제목", "", ""], ["ID", "NAME", "VALUE"], ["k1", "칼", "10"]]
    b = [["제목", "", "", ""], ["ID", "GRADE", "NAME", "VALUE"], ["k1", "A", "칼", "10"]]
    assert match_columns(a, b, key_row=1) == [(0, 0), (None, 1), (1, 2), (2, 3)]


def test_trailing_blank_headers_are_not_columns():
    a = [["ID", "NAME", "", ""], ["k1", "x", "", ""]]
    b = [["ID", "NAME"], ["k1", "x"]]
    assert match_columns(a, b) == [(0, 0), (1, 1)]


# ── 맞출 수 없으면 손대지 않는다 ────────────────────────────────────────────
# 판정은 보수적이다. 애매하면 None 을 돌려 호출부가 위치 기준을 쓰게 한다.

@pytest.mark.parametrize("a,b,why", [
    ([["ID", "", "VALUE"]], [["ID", "X", "VALUE"]], "가운데 빈 헤더"),
    ([], [["ID"]], "A 가 비었다"),
    ([["ID"]], [], "B 가 비었다"),
    ([[]], [["ID"]], "헤더가 빈 행"),
])
def test_unmatchable_returns_none(a, b, why):
    assert match_columns(a, b) is None, why


# ── 같은 이름이 여럿일 때 — '몇 번째냐' 로 구분한다 ─────────────────────────
# 예전엔 이것도 포기 조건이었다. 그런데 실제 데이터에서 열 매칭이 꺼지는 79개 파일 중
# **78개가 이 이유** 하나였다 — '#X' 는 바로 왼쪽 열의 주석 열이라 한 시트에 '#' 가
# 아홉 개씩 있다(57개 파일이 '#', 26개가 '#Desc'). 그래서 n번째끼리 맞춘다.

def test_duplicate_names_pair_in_order():
    a = [["UID", "#", "Count", "#"]]
    b = [["UID", "#", "Count", "#"]]
    assert match_columns(a, b) == [(0, 0), (1, 1), (2, 2), (3, 3)]


def test_extra_duplicates_on_one_side_are_one_sided():
    """40 빌드가 주석 열을 뒤에 더 붙인 모양(실측: Data_MiniGameRoulette_CS.xlsx).

    위치 기준으로는 이 아홉 열에 59개 셀을 써 넣으려 했다 — 40 에만 있는 주석 열이다.
    """
    a = [["UID", "ItemID", "#ItemID", "Count"]]
    b = [["UID", "ItemID", "#ItemID", "Count", "#", "#"]]
    assert match_columns(a, b) == [(0, 0), (1, 1), (2, 2), (3, 3),
                                   (None, 4), (None, 5)]


def test_a_missing_duplicate_is_one_sided_too():
    """반대쪽 — A 에만 있는 꼬리 주석 열(실측: Data_ShopGoodsRewardSelect_CS.xlsx).

    위치 기준으로는 B 에 **없는 열**에 쓰려 해서 고스트 열이 생길 뻔했다.
    """
    a = [["UID", "#", "#", "#"]]
    b = [["UID", "#", "#"]]
    assert match_columns(a, b) == [(0, 0), (1, 1), (2, 2), (3, None)]


def test_duplicates_do_not_disturb_a_real_insertion():
    """중복 이름이 있어도 가운데 끼인 **다른** 열은 제대로 한쪽 전용으로 잡힌다."""
    a = [["ID", "#", "NAME", "#"]]
    b = [["ID", "#", "GRADE", "NAME", "#"]]
    assert match_columns(a, b) == [(0, 0), (1, 1), (None, 2), (2, 3), (3, 4)]


def test_a_duplicate_inserted_in_the_middle_mispairs_only_its_own_kind():
    """같은 이름이 중간에 끼면 그 뒤의 같은 이름끼리 한 칸씩 밀려 짝이 틀린다.

    숨기지 않고 한계를 적어 둔다. 다만 **이름이 있는 열은 전부 제자리를 찾고**,
    틀리는 것은 주석 열끼리의 짝뿐이다. 위치 기준은 그 오른쪽의 멀쩡한 열까지 전부
    밀린다 — 더 나빠지는 경우가 아니다.
    """
    a = [["ID", "X", "#", "Y", "#"]]
    b = [["ID", "X", "#", "NEW", "#", "Y", "#"]]
    cm = match_columns(a, b)
    assert (0, 0) in cm and (1, 1) in cm, f"ID·X 가 제자리가 아니다: {cm}"
    assert (3, 5) in cm, f"Y 가 B 의 5열을 못 찾았다(위치 기준이면 3열로 밀린다): {cm}"
    assert (4, 4) in cm, f"한계 — A 의 2번째 '#' 는 B 의 2번째 '#' 에 붙는다: {cm}"


def test_key_row_out_of_range_returns_none():
    assert match_columns([["ID"]], [["ID"]], key_row=5) is None


def test_whitespace_around_header_is_ignored():
    a = [["ID", " NAME "], ["k1", "x"]]
    b = [["ID", "NAME"], ["k1", "y"]]
    assert match_columns(a, b) == [(0, 0), (1, 1)]


# ── compute_diff 연결 ────────────────────────────────────────────────────────

def _cells(m):
    return [[(st, av, bv) for st, av, bv in row] for row in m]


def test_without_col_meta_nothing_changes():
    """기존 호출은 한 글자도 달라지면 안 된다 — 위치 기준 그대로."""
    m, meta = compute_diff(A, B, key_col=0, key_row=0)
    assert [c[0] for c in m[1]] == [STATUS_SAME, "modified", "modified", STATUS_ADDED]


def test_with_col_meta_only_the_new_column_differs():
    cm = match_columns(A, B)
    m, meta = compute_diff(A, B, key_col=0, key_row=0, col_meta=cm)
    body = m[1]
    assert len(body) == 4
    assert body[0][0] == STATUS_SAME                      # ID
    assert body[1] == (STATUS_ADDED, "", "A")             # GRADE — B 에만
    assert body[2][0] == STATUS_SAME, body[2]             # NAME 끼리 맞았다
    assert body[3][0] == STATUS_SAME, body[3]             # VALUE 끼리 맞았다


def test_matching_collapses_the_false_changes():
    """열 하나 끼었을 뿐인데 9개가 변경으로 잡히던 것이 GRADE 열만 남아야 한다."""
    def changed(col_meta):
        m, _ = compute_diff(A, B, key_col=0, key_row=0, col_meta=col_meta)
        return sum(1 for row in m for cell in row if cell[0] != STATUS_SAME)
    assert changed(None) == 9
    assert changed(match_columns(A, B)) == 3, "헤더행 + 본문 2행의 GRADE 열만"


def test_key_column_is_found_on_each_side_separately():
    """키 열이 두 파일에서 다른 자리에 있어도 행이 맞아야 한다."""
    a = [["ID", "NAME"], ["k1", "칼"], ["k2", "방패"]]
    b = [["NAME", "ID"], ["칼", "k1"], ["방패", "k2"]]
    cm = match_columns(a, b)
    assert cm == [(0, 1), (1, 0)]
    m, meta = compute_diff(a, b, key_col=0, key_row=0, col_meta=cm)
    assert len(m) == 3, "행이 쪼개졌다"
    assert all(cell[0] == STATUS_SAME for row in m for cell in row), _cells(m)
    assert meta[1] == (1, 1) and meta[2] == (2, 2)


def test_key_column_missing_on_one_side_does_not_swallow_that_file():
    """키로 고른 열이 A 에만 있으면 B 쪽 키를 읽을 자리가 없다.

    그대로 두면 B 의 **모든 행이 빈 키로 탈락**해 비교에서 통째로 사라진다 — 화면엔
    'B 에 아무것도 없음' 처럼 보인다. 그래서 항등 매핑으로 물러선다.
    """
    a = [["ID", "NAME"], ["k1", "a"]]
    b = [["ID", "GRADE"], ["k1", "A"]]
    cm = match_columns(a, b)
    assert cm == [(0, 0), (1, None), (None, 1)], cm
    m, meta = compute_diff(a, b, key_col=1, key_row=0, col_meta=cm)
    body = meta[1:]
    assert any(b_row is not None for _a_row, b_row in body), (
        f"B 의 행이 통째로 사라졌다: {meta}")


def test_row_order_mode_ignores_col_meta():
    """키 없음(ROW 순서)은 헤더 개념이 없다 — 예전 그대로."""
    m, meta = compute_diff(A, B, key_col=-1, col_meta=match_columns(A, B))
    assert len(m[0]) == 4 and meta[0] == (0, 0)


def test_new_rows_still_land_in_place_with_matching():
    """열 매칭을 켜도 신규 행 끼우기 규칙은 그대로여야 한다."""
    a = [["ID", "NAME"], ["k1", "a"], ["k3", "c"]]
    b = [["ID", "GRADE", "NAME"], ["k1", "A", "a"], ["k2", "B", "b"], ["k3", "C", "c"]]
    cm = match_columns(a, b)
    m, meta = compute_diff(a, b, key_col=0, key_row=0, col_meta=cm)
    assert [row[0][2] for row in m] == ["ID", "k1", "k2", "k3"], "신규 행이 제자리가 아니다"
    assert meta[2] == (None, 2), "k2 는 B 에만 있는 행"
