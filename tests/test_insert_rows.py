# -*- coding: utf-8 -*-
"""행 삽입이 시트를 깨뜨리지 않는다 — 꼬리에 남은 빈 행·서식 전용 셀을 만났을 때.

실기에서 드러났다. 합성 조작("셀 하나 고치기", "행 몇 개 지우기")으로는 멀쩡했는데,
앱과 **같은 경로**로 367쌍을 실제 머지해 보니 행 삽입에서 셋이 터졌다.

  1. **행 번호 중복** — Data_RewardGroup_CS.xlsx 는 셀이 있는 행이 423개인데 <row>
     요소는 1207개다(424~1207 은 빈 행). 새 행을 424부터 만들어 같은 번호의 <row> 가
     둘 생기고 순서도 뒤집혔다(…1207, 424, 425…). 엑셀이 파일을 거부했다.
  2. **빈 행 아래에 붙음** — Data_TextPCSkillTable_CS.xlsm 은 값이 있는 마지막 행이
     4166인데 서식만 준 <c> 가 4560행까지 있다. 새 행이 빈 행 394개 **아래**에 붙었다.
  3. **셀 좌표 중복** — 1을 고쳐 기존 행을 재사용했더니, 그 행에 이미 있던 서식 전용
     <c> 와 좌표가 겹쳤다.

셋 다 **값 대조로는 안 잡힌다.** 로더는 행·열 번호로 격자를 채우므로 순서가 어긋나도,
좌표가 겹쳐도 같은 값이 나온다. 그래서 저장 검증에 순서·중복 검사를 넣었고, 아래
'가드' 절이 그게 실제로 막는지 본다.
"""
import os
import zipfile

import openpyxl
import pytest
from lxml import etree

from excelmerge.loaders import clear_values_cache, load_values_any
from excelmerge.xlsx_writer import (_sheet_order_mismatches,
                                    _write_patches_to_file)

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
HEAD = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/'
        'spreadsheetml/2006/main"><sheetData>')
TAIL = "</sheetData></worksheet>"


def _data_rows(n):
    return "".join(
        f'<row r="{r}"><c r="A{r}" t="str"><v>k{r}</v></c>'
        f'<c r="B{r}" t="str"><v>v{r}</v></c></row>' for r in range(1, n + 1))


def _make(tmp_path, body, name="a.xlsx"):
    p = tmp_path / name
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.active.append(["", ""])
    wb.save(str(p))
    tmp = str(p) + ".b"
    with zipfile.ZipFile(str(p)) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            zout.writestr(n, (HEAD + body + TAIL).encode("utf-8")
                          if n.endswith("worksheets/sheet1.xml") else zin.read(n))
    os.replace(tmp, str(p))
    assert _sheet_order_mismatches(str(p)) == [], "전제: 저장 전은 성하다"
    clear_values_cache()
    return p


def _rows(path):
    """[(행번호, [셀좌표...])] — XML 에 적힌 **순서 그대로**."""
    out = []
    with zipfile.ZipFile(str(path)) as z:
        sd = etree.fromstring(
            z.read("xl/worksheets/sheet1.xml")).find(NS + "sheetData")
        for row_el in sd:
            out.append((int(row_el.get("r")),
                        [c.get("r") for c in row_el]))
    return out


def _insert(path, rows):
    clear_values_cache()
    _write_patches_to_file(str(path), {}, rows)


NEW = [[(0, "새키1"), (1, "새값1")], [(0, "새키2"), (1, "새값2")]]


# ── 1. 꼬리에 빈 <row> 가 있을 때 ───────────────────────────────────────────

@pytest.fixture
def blank_tail(tmp_path):
    """데이터 3행 + 빈 <row> 4~8행 — 실기의 RewardGroup 과 같은 모양."""
    return _make(tmp_path, _data_rows(3) + "".join(
        f'<row r="{r}"/>' for r in range(4, 9)))


def test_no_duplicate_row_numbers(blank_tail):
    """이 테스트가 그 복구 창이다 — 같은 번호의 <row> 가 둘이면 엑셀이 거부한다."""
    _insert(blank_tail, NEW)
    nums = [n for n, _ in _rows(blank_tail)]
    assert len(nums) == len(set(nums)), f"행 번호가 겹쳤다: {nums}"


def test_rows_stay_in_order(blank_tail):
    _insert(blank_tail, NEW)
    nums = [n for n, _ in _rows(blank_tail)]
    assert nums == sorted(nums), f"행 순서가 뒤집혔다: {nums}"


def test_the_new_rows_land_right_after_the_data(blank_tail):
    _insert(blank_tail, NEW)
    clear_values_cache()
    g = load_values_any(str(blank_tail))
    assert g[3][0] == "새키1" and g[4][0] == "새키2", g[:6]


def test_the_existing_blank_row_is_reused(blank_tail):
    """빈 행을 새로 만들지 않고 그 자리를 채운다 — 행 요소 수가 늘지 않는다."""
    before = len(_rows(blank_tail))
    _insert(blank_tail, NEW)
    assert len(_rows(blank_tail)) == before, "빈 행을 두고 새 행을 또 만들었다"


def test_the_saved_sheet_is_well_ordered(blank_tail):
    _insert(blank_tail, NEW)
    assert _sheet_order_mismatches(str(blank_tail)) == []


# ── 2. 꼬리에 서식만 준 <c> 가 있을 때 ──────────────────────────────────────

@pytest.fixture
def styled_tail(tmp_path):
    """데이터 3행 + 값 없이 <c> 만 있는 4~6행 — 실기의 TextPCSkillTable 과 같은 모양."""
    styled = "".join(f'<row r="{r}"><c r="A{r}" s="1"/><c r="B{r}" s="1"/></row>'
                     for r in range(4, 7))
    return _make(tmp_path, _data_rows(3) + styled)


def test_new_rows_do_not_land_below_empty_rows(styled_tail):
    """<c> 가 있다고 데이터인 건 아니다 — 비교에 쓰는 로더와 같은 기준으로 세야 한다.

    실기에서는 이 어긋남 때문에 새 행이 빈 행 **394개 아래**에 붙을 뻔했다.
    """
    _insert(styled_tail, NEW)
    clear_values_cache()
    g = load_values_any(str(styled_tail))
    assert g[3][0] == "새키1", f"새 행이 빈 행 아래에 붙었다: {[r[0] for r in g]}"


def test_no_duplicate_cell_refs(styled_tail):
    """기존 행을 다시 쓸 때 그 행의 <c> 도 다시 써야 한다 — 새로 만들면 좌표가 겹친다."""
    _insert(styled_tail, NEW)
    for num, refs in _rows(styled_tail):
        assert len(refs) == len(set(refs)), f"{num}행에 셀 좌표가 겹쳤다: {refs}"
    assert _sheet_order_mismatches(str(styled_tail)) == []


def test_cells_stay_in_column_order(styled_tail):
    _insert(styled_tail, NEW)
    from openpyxl.utils import column_index_from_string
    import re
    for num, refs in _rows(styled_tail):
        cols = [column_index_from_string(re.match(r"([A-Z]+)", r).group(1))
                for r in refs]
        assert cols == sorted(cols), f"{num}행의 셀 순서가 뒤집혔다: {refs}"


# ── 3. 평범한 시트는 그대로 ────────────────────────────────────────────────

def test_a_plain_sheet_still_appends(tmp_path):
    p = _make(tmp_path, _data_rows(3))
    _insert(p, NEW)
    clear_values_cache()
    g = load_values_any(str(p))
    assert len(g) == 5 and g[3][0] == "새키1" and g[4][1] == "새값2", g


def test_inserting_nothing_changes_nothing(blank_tail):
    before = _rows(blank_tail)
    _insert(blank_tail, [])
    assert _rows(blank_tail) == before


# ── 4. 가드가 살아 있나 ─────────────────────────────────────────────────────
# 값 대조는 이 사고를 통과시킨다 — 로더가 번호로 격자를 채우므로 순서가 어긋나도,
# 좌표가 겹쳐도 같은 값이 나온다. 그래서 순서·중복 검사를 넣었다. 버그를 되살려
# 그게 실제로 막는지 본다.

# ── 4. 같은 저장에서 행을 지우면서 넣을 때 ─────────────────────────────────
#
# 실기에서 드러났다(Data_MailBox_CS.xlsx, 367쌍 전수 머지). 1~276 행에서 **끝의**
# 275·276 을 지우고 2행을 넣으면 저장이 거부됐다 — "275행이 두 번 들어 있습니다".
#
# 왜: 번호 당기기(_renumber_after_delete)가 삽입 **뒤**에 돌았다. _append_rows 는
# 살아남은 마지막 행 다음 번호를 붙이는데 그건 아직 '지우기 전' 번호 공간이고,
# 뒤이어 당기면 새 행도 같이 밀린다. 당기는 양은 '자기보다 작은 삭제 행 수'라
# 275 는 0칸, 276 은 1칸 밀려 **둘 다 275** 가 됐다.
#
# 가운데를 지울 때는 새 행 번호가 삭제 행보다 전부 위라 밀림이 모두 같아
# 우연히 맞아떨어졌다. 그래서 삽입·삭제를 각각 시험한 테스트로는 오래 안 보였다.


def _save(path, inserts, delete_rows=None):
    clear_values_cache()
    _write_patches_to_file(str(path), {}, inserts, set(delete_rows or ()))


def test_deleting_the_last_rows_while_inserting_keeps_numbers_unique(tmp_path):
    """끝 2행을 지우고 2행을 넣는다 — 번호가 겹치면 엑셀이 파일을 거부한다."""
    p = _make(tmp_path, _data_rows(10))
    _save(p, NEW, {9, 10})
    nums = [n for n, _ in _rows(p)]
    assert len(nums) == len(set(nums)), f"번호 중복: {nums}"
    assert nums == sorted(nums), f"순서가 뒤집혔다: {nums}"
    assert _sheet_order_mismatches(str(p)) == []


def test_deleting_the_last_rows_while_inserting_keeps_them_contiguous(tmp_path):
    """8행이 남고 2행이 붙으니 1~10 이 빈틈없이 이어져야 한다."""
    p = _make(tmp_path, _data_rows(10))
    _save(p, NEW, {9, 10})
    assert [n for n, _ in _rows(p)] == list(range(1, 11))


def test_the_inserted_values_land_after_the_survivors(tmp_path):
    """번호만 맞고 값이 엉뚱한 자리에 가면 소용이 없다."""
    p = _make(tmp_path, _data_rows(10))
    _save(p, NEW, {9, 10})
    clear_values_cache()
    grid = load_values_any(str(p))
    assert [r[0] for r in grid[:8]] == [f"k{i}" for i in range(1, 9)]
    assert [r[0] for r in grid[8:10]] == ["새키1", "새키2"]


@pytest.mark.parametrize("dead", [{5, 6}, {1, 2}, {9, 10}, {1, 10}, {4, 7}])
def test_insert_with_deletion_anywhere_is_sound(tmp_path, dead):
    """지운 자리가 앞이든 가운데든 끝이든 결과는 같은 모양이어야 한다.

    가운데만 맞고 끝에서 틀렸던 것이 이 버그였다 — 자리를 바꿔 가며 못 박는다.
    """
    p = _make(tmp_path, _data_rows(10), name=f"d{min(dead)}_{max(dead)}.xlsx")
    _save(p, NEW, dead)
    nums = [n for n, _ in _rows(p)]
    assert nums == list(range(1, 11)), f"{sorted(dead)} 삭제 후: {nums}"
    assert _sheet_order_mismatches(str(p)) == []
    clear_values_cache()
    grid = load_values_any(str(p))
    kept = [f"k{i}" for i in range(1, 11) if i not in dead]
    assert [r[0] for r in grid[:8]] == kept
    assert [r[0] for r in grid[8:10]] == ["새키1", "새키2"]


def _buggy_append(sheetdata, insert_rows):
    """예전 _append_rows — 빈 <row> 를 무시하고 새로 만든다."""
    from excelmerge.xlsx_writer import _TAG_C, _TAG_ROW, _cell_ref, _set_cell_value
    if not insert_rows:
        return
    last = max((int(r.get("r", 0)) for r in sheetdata if list(r)), default=0)
    nxt = last + 1 if last > 0 else 1
    for cells in insert_rows:
        row_el = etree.SubElement(sheetdata, _TAG_ROW)
        row_el.set("r", str(nxt))
        for col_idx, val in [(c[0], c[1]) for c in cells]:
            if val == "":
                continue
            c_el = etree.SubElement(row_el, _TAG_C)
            c_el.set("r", _cell_ref(nxt - 1, col_idx))
            _set_cell_value(c_el, val)
        nxt += 1


def test_the_verifier_blocks_duplicate_rows(blank_tail, monkeypatch):
    import excelmerge.xlsx_writer as W
    monkeypatch.setattr(W, "_append_rows", _buggy_append)
    before = _rows(blank_tail)
    with pytest.raises(ValueError) as e:
        _insert(blank_tail, NEW)
    msg = str(e.value)
    assert "두 번 들어 있습니다" in msg or "뒤집혔습니다" in msg, msg
    assert _rows(blank_tail) == before, "원본이 덮였다"
    assert not os.path.exists(str(blank_tail) + ".tmp_merge"), "임시 파일이 남았다"


def test_the_verifier_blocks_rows_below_blank_ones(styled_tail, monkeypatch):
    """빈 행 아래에 붙으면 전수 대조가 행 수 차이로 막는다."""
    import excelmerge.xlsx_writer as W
    monkeypatch.setattr(W, "_append_rows", _buggy_append)
    with pytest.raises(ValueError) as e:
        _insert(styled_tail, NEW)
    assert "행 수가 다릅니다" in str(e.value), str(e.value)


@pytest.mark.parametrize("body, why", [
    ('<row r="1"><c r="A1"><v>1</v></c></row><row r="1"><c r="A1"><v>2</v></c></row>',
     "행 번호 중복"),
    ('<row r="2"><c r="A2"><v>1</v></c></row><row r="1"><c r="A1"><v>2</v></c></row>',
     "행 순서 역전"),
    ('<row r="1"><c r="A1"><v>1</v></c><c r="A1"><v>2</v></c></row>', "셀 좌표 중복"),
    ('<row r="1"><c r="B1"><v>1</v></c><c r="A1"><v>2</v></c></row>', "셀 순서 역전"),
    ('<row r="1"><c r="A2"><v>1</v></c></row>', "셀이 다른 행에 있다"),
])
def test_the_order_check_sees_each_kind(tmp_path, body, why):
    """검사기 자체가 각 모양을 알아보는가 — 못 보면 가드가 통과만 시킨다."""
    p = tmp_path / "bad.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["", ""])
    wb.save(str(p))
    tmp = str(p) + ".b"
    with zipfile.ZipFile(str(p)) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            zout.writestr(n, (HEAD + body + TAIL).encode("utf-8")
                          if n.endswith("worksheets/sheet1.xml") else zin.read(n))
    os.replace(tmp, str(p))
    assert _sheet_order_mismatches(str(p)), why


def test_the_order_check_passes_a_good_sheet(tmp_path):
    assert _sheet_order_mismatches(str(_make(tmp_path, _data_rows(5)))) == []
