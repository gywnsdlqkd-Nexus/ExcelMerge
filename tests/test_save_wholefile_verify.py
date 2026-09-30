# -*- coding: utf-8 -*-
"""저장 결과 전수 대조 — '자기가 쓴 셀만' 보는 검사로는 못 잡는 것들.

기존 검사(_patch_mismatches)는 **덮어쓴 셀만** 되읽어 본다. 의도한 자리에 의도한 값을
썼으면 통과한다. 그래서 이런 사고를 못 잡았다.

  · v206 — 지웠어야 할 행이 '빈 행' 으로 살아남았다. 패치는 전부 의도대로 들어갔다.
  · 앞으로 들어올 열 매칭 — 엉뚱한 열에 써도 '그 좌표에 그 값' 이라 통과한다.

그래서 저장 직전 임시 파일을 되읽어 **시트 전체**를 기대 격자와 대조한다. 여기서
고정하는 것은 '전수로 본다' 와 '멀쩡한 저장을 막지 않는다' 두 가지다.
"""
import openpyxl
import pytest

from excelmerge.loaders import load_values_any, clear_values_cache
from excelmerge.xlsx_writer import (
    _expected_after, _grid_mismatches, _trim_grid, _write_patches_to_file,
)


def _make_xlsx(path, rows, sheet="Sheet1"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for r in rows:
        ws.append(r)
    wb.save(str(path))
    return str(path)


def _values(path, sheet=None):
    clear_values_cache()
    return load_values_any(str(path), sheet_name=sheet)


ROWS = [["k1", "v1", "n1"], ["k2", "v2", "n2"], ["k3", "v3", "n3"]]


# ── 기대 격자 계산 (순수) ────────────────────────────────────────────────────
# _patch_sheet_xml 의 순서를 흉내 낸다: 덮어쓰기(저장 전 좌표) → 행 삭제 → 열 삭제
# → 신규 행 꼬리 추가. 이 순서가 틀리면 멀쩡한 저장이 막힌다.

def test_patch_lands_at_the_pre_delete_coordinate():
    got = _expected_after(ROWS, {"B3": "새값"})
    assert got[2] == ["k3", "새값", "n3"]


def test_deleting_a_row_pulls_the_rest_up():
    got = _expected_after(ROWS, {}, delete_row_nums={2})
    assert [r[0] for r in got] == ["k1", "k3"]


def test_patch_coordinates_are_pre_delete_not_post():
    """B3 은 '저장 전' 3행이다 — 행 삭제로 자리가 밀려도 그 행을 따라가야 한다."""
    got = _expected_after(ROWS, {"B3": "새값"}, delete_row_nums={1})
    assert got == [["k2", "v2", "n2"], ["k3", "새값", "n3"]]


def test_deleted_column_becomes_empty_not_removed():
    """열 삭제는 <c> 만 지운다 — 열 자체는 남고 비어 있다(열 문자가 밀리지 않는다)."""
    got = _expected_after(ROWS, {}, delete_col_letters={"B"})
    assert got[0] == ["k1", "", "n1"]


def test_inserted_rows_go_to_the_tail():
    got = _expected_after(ROWS, {}, insert_rows=[[(0, "k9"), (2, "n9")]])
    assert got[-1] == ["k9", "", "n9"], got[-1]


def test_a_written_formula_reads_back_empty():
    """수식은 캐시값 없이 쓰이므로 되읽으면 빈 칸 — 기대값도 빈 칸이어야 한다."""
    got = _expected_after(ROWS, {"B1": "=SUM(A1:A2)"})
    assert got[0][1] == ""


# ── 대조 (순수) ──────────────────────────────────────────────────────────────

def test_trailing_blanks_are_not_a_difference():
    """로더는 파일에 없는 꼬리를 돌려주지 않는다 — 모양 차이로 막으면 안 된다."""
    assert _grid_mismatches([["a"]], [["a", "", ""], []]) == []


def test_a_blank_in_the_middle_is_a_real_difference():
    assert _grid_mismatches([["a", "", "c"]], [["a", "b", "c"]]) != []


def test_mismatch_names_the_cell():
    bad = _grid_mismatches([["a", "틀린값"]], [["a", "맞는값"]])
    assert bad and bad[0].startswith("B1:"), bad


def test_row_count_difference_is_reported():
    bad = _grid_mismatches([["a"]], [["a"], ["b"]])
    assert any("행 수" in b for b in bad), bad


def test_trim_keeps_inner_blanks():
    assert _trim_grid([["a", "", "b", "", ""], ["", ""]]) == [["a", "", "b"]]


# ── 실제 저장 경로 ───────────────────────────────────────────────────────────

def test_a_good_save_is_not_blocked(tmp_path):
    """전수 대조를 켠 뒤에도 멀쩡한 저장은 그대로 통과해야 한다."""
    p = _make_xlsx(tmp_path / "a.xlsx", ROWS)
    _write_patches_to_file(p, {"B2": "새값"}, delete_row_nums={1})
    assert _values(p) == [["k2", "새값", "n2"], ["k3", "v3", "n3"]]


def test_a_row_that_should_vanish_but_survives_is_caught(tmp_path, monkeypatch):
    """v206 그 버그 — 지워야 할 행이 빈 행으로 남는다. 패치는 전부 의도대로 들어갔다.

    옛 검사는 '자기가 쓴 셀' 만 봐서 통과시켰다. 전수 대조는 행 수에서 걸린다.
    """
    import excelmerge.xlsx_writer as xw
    p = _make_xlsx(tmp_path / "a.xlsx", ROWS)
    before = open(p, "rb").read()
    real = xw._patch_sheet_xml

    def drop_the_delete(data, patches, insert_rows=None, delete_row_nums=None,
                        *a, **k):
        # 삭제 지시를 무시 → 행이 빈 채로 살아남는다(정확히 v206 의 모양)
        return real(data, patches, insert_rows, set(), *a, **k)

    monkeypatch.setattr(xw, "_patch_sheet_xml", drop_the_delete)
    with pytest.raises(ValueError, match="저장 검증 실패"):
        _write_patches_to_file(p, {"A2": "", "B2": "", "C2": ""},
                               delete_row_nums={2})
    assert open(p, "rb").read() == before, "검증 실패인데 원본이 바뀌었다"


def test_an_untouched_cell_that_changed_is_caught(tmp_path, monkeypatch):
    """건드리지 않기로 한 셀이 바뀌면 잡아야 한다 — 옛 검사는 볼 생각조차 안 했다."""
    import excelmerge.xlsx_writer as xw
    p = _make_xlsx(tmp_path / "a.xlsx", ROWS)
    before = open(p, "rb").read()
    real = xw._patch_sheet_xml

    def also_touch_c3(data, patches, *a, **k):
        return real(data, {**patches, "C3": "몰래 바꾼 값"}, *a, **k)

    monkeypatch.setattr(xw, "_patch_sheet_xml", also_touch_c3)
    with pytest.raises(ValueError, match="저장 검증 실패"):
        _write_patches_to_file(p, {"B1": "새값"})
    assert open(p, "rb").read() == before


def test_a_value_written_to_the_wrong_column_is_caught(tmp_path, monkeypatch):
    """열 매칭이 들어오면 생길 수 있는 사고 — 의도한 값을 **옆 열**에 쓴다.

    셀만 보는 검사로는 못 잡는다(그 좌표엔 그 값이 없으니 잡히긴 하나, 여기서는
    옆 열로 밀린 흔적까지 전수로 드러나는지를 본다).
    """
    import excelmerge.xlsx_writer as xw
    from openpyxl.utils import column_index_from_string, get_column_letter
    p = _make_xlsx(tmp_path / "a.xlsx", ROWS)
    before = open(p, "rb").read()
    real = xw._patch_sheet_xml

    def shift_one_column(data, patches, *a, **k):
        moved = {}
        for ref, v in patches.items():
            m = xw._COL_RE.match(ref)
            col = column_index_from_string(m.group(1))
            moved[f"{get_column_letter(col + 1)}{m.group(2)}"] = v
        return real(data, moved, *a, **k)

    monkeypatch.setattr(xw, "_patch_sheet_xml", shift_one_column)
    with pytest.raises(ValueError, match="저장 검증 실패"):
        _write_patches_to_file(p, {"B2": "새값"})
    assert open(p, "rb").read() == before


def test_inserted_row_is_verified_too(tmp_path):
    p = _make_xlsx(tmp_path / "a.xlsx", ROWS)
    _write_patches_to_file(p, {}, insert_rows=[[(0, "k9"), (1, "v9"), (2, "n9")]])
    rows = _values(p)
    assert rows[-1] == ["k9", "v9", "n9"] and len(rows) == 4


def test_unreadable_original_does_not_block_the_save(tmp_path, monkeypatch):
    """원본을 못 읽는다고 멀쩡한 저장을 막지는 않는다 — 셀 단위 검사는 그대로 돈다."""
    import excelmerge.xlsx_writer as xw
    p = _make_xlsx(tmp_path / "a.xlsx", ROWS)
    monkeypatch.setattr(xw, "_before_values", lambda *a, **k: None)
    _write_patches_to_file(p, {"B2": "새값"})
    assert _values(p)[1] == ["k2", "새값", "n2"]


def test_the_original_is_read_before_it_is_replaced(tmp_path):
    """기준 격자는 **교체 전** 원본이어야 한다 — 교체 뒤를 읽으면 대조가 무의미해진다."""
    import inspect
    import excelmerge.xlsx_writer as xw
    src = inspect.getsource(xw._write_patches_to_file)
    assert src.index("_verify_saved(") < src.index("os.replace(")
