# -*- coding: utf-8 -*-
"""'빈 열' 판정은 **새로 넣는 행까지** 보고 해야 한다.

머지가 어떤 열을 전부 빈값으로 덮으면 그 열은 '삭제'로 승격된다(화면에서 비운다).
그 판정이 기존 행만 보고 있었다. 그런데 같은 머지가 **새 행**을 넣고 그 행이 그 열에
값을 가지면, 그 열은 비어 있지 않다.

그러면 저장 검증이 막는다 — 모델(_expected_after)은 열을 비운 **뒤** 새 행을 붙이므로
"새 행은 그 값을 갖는다" 로 계산하고, 쓰기는 새 행을 붙인 **뒤** 열을 지워 그 값을
걷어낸다. 실측: `C7: 예상 '새값' → 파일에는 ''` 로 저장 거부.

원인은 순서가 아니라 판단이다. 그 열은 애초에 '빈 열' 이 아니다. 아래 테스트는
판정·저장·결과를 각각 못 박는다.
"""
import os

import openpyxl
import pytest

from excelmerge.loaders import clear_values_cache, load_values_any
from excelmerge.xlsx_writer import (_expected_after,
                                    _promote_empty_cols_to_delete,
                                    _write_patches_to_file)


@pytest.fixture
def sheet(tmp_path):
    """A=ID, B=KEEP, C=GONE 의 5행짜리 표."""
    p = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["ID", "KEEP", "GONE"])
    for i in range(1, 6):
        ws.append([f"k{i}", f"v{i}", f"g{i}"])
    wb.save(str(p))
    clear_values_cache()
    return str(p)


def _blank_c():
    """C열 전체를 빈값으로 덮는 패치 — 이것만 보면 'C는 빈 열' 이다."""
    return {f"C{r}": "" for r in range(1, 7)}


NEW_WITH_C = [[(0, "k9"), (1, "v9"), (2, "새값")]]      # 새 행이 C에 값을 갖는다
NEW_WITHOUT_C = [[(0, "k9"), (1, "v9")]]               # 새 행도 C가 비었다


# ── 판정 ────────────────────────────────────────────────────────────────────

def test_a_column_the_new_rows_fill_is_not_empty(sheet):
    """새 행이 채우는 열은 지우면 안 된다."""
    _p, _r, cols = _promote_empty_cols_to_delete(
        _blank_c(), set(), sheet, None, insert_rows=NEW_WITH_C)
    assert "C" not in cols, f"새 행이 값을 갖는 C 를 지우려 한다: {sorted(cols)}"


def test_a_genuinely_empty_column_is_still_promoted(sheet):
    """고치면서 원래 동작을 잃으면 안 된다 — 아무도 안 채우면 그대로 지운다."""
    _p, _r, cols = _promote_empty_cols_to_delete(
        _blank_c(), set(), sheet, None, insert_rows=NEW_WITHOUT_C)
    assert "C" in cols, "비어 있는데도 안 지운다"


def test_no_inserts_behaves_as_before(sheet):
    """insert_rows 를 안 넘기는 호출(옛 서명)도 그대로 돌아야 한다."""
    _p, _r, cols = _promote_empty_cols_to_delete(_blank_c(), set(), sheet, None)
    assert "C" in cols


def test_an_empty_value_in_a_new_row_does_not_save_the_column(sheet):
    """새 행이 그 열에 **빈값**을 갖는 건 채우는 게 아니다."""
    inserts = [[(0, "k9"), (1, "v9"), (2, "")]]
    _p, _r, cols = _promote_empty_cols_to_delete(
        _blank_c(), set(), sheet, None, insert_rows=inserts)
    assert "C" in cols


# ── 저장이 실제로 되는가 ────────────────────────────────────────────────────

def test_the_save_goes_through(sheet):
    """고치기 전에는 저장 검증이 막았다 — C7: 예상 '새값' → 파일에는 ''."""
    patches, del_rows, del_cols = _promote_empty_cols_to_delete(
        _blank_c(), set(), sheet, None, insert_rows=NEW_WITH_C)
    _write_patches_to_file(sheet, patches, NEW_WITH_C, del_rows, del_cols)


def test_the_result_is_what_the_model_says(sheet):
    """모델과 파일이 같은 그림이어야 한다 — 그 둘이 어긋난 게 이 버그였다."""
    clear_values_cache()
    before = load_values_any(sheet)
    patches, del_rows, del_cols = _promote_empty_cols_to_delete(
        _blank_c(), set(), sheet, None, insert_rows=NEW_WITH_C)
    expected = _expected_after(before, patches, NEW_WITH_C, del_rows, del_cols)
    _write_patches_to_file(sheet, patches, NEW_WITH_C, del_rows, del_cols)
    clear_values_cache()
    after = load_values_any(sheet)
    assert [list(r) for r in after] == [list(r) for r in expected]


def test_the_old_rows_are_blanked_and_the_new_one_keeps_its_value(sheet):
    """사람이 읽는 형태로 한 번 더 — 기존 행의 C는 비고, 새 행의 C는 남는다."""
    patches, del_rows, del_cols = _promote_empty_cols_to_delete(
        _blank_c(), set(), sheet, None, insert_rows=NEW_WITH_C)
    _write_patches_to_file(sheet, patches, NEW_WITH_C, del_rows, del_cols)
    clear_values_cache()
    after = load_values_any(sheet)
    assert [r[2] for r in after[:6]] == [""] * 6, "기존 행의 C가 안 비었다"
    assert after[6][0] == "k9" and after[6][2] == "새값", f"새 행: {after[6]}"


# ── 호출부가 실제로 넘기는가 ────────────────────────────────────────────────

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@pytest.mark.parametrize("name", ["excelmerge/workers.py", "merge_check.py"])
def test_callers_pass_the_inserted_rows(name):
    """고쳐 놔도 호출부가 안 넘기면 아무 소용이 없다."""
    with open(os.path.join(ROOT, name), encoding="utf-8") as f:
        src = f.read()
    i = src.index("_promote_empty_cols_to_delete(")
    call = src[i:i + 300]
    assert "insert_rows=" in call, f"{name} 가 insert_rows 를 안 넘긴다"
