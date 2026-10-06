# -*- coding: utf-8 -*-
"""행을 지우면 수식의 셀 참조도 함께 다시 쓴다 — 엑셀이 하는 그대로.

예전엔 좌표(<c r>)만 당기고 수식 본문은 그대로 뒀다. 그래서 'B3: =B2+1' 이 B2 로
당겨지면 '=B2+1' 그대로라 **자기 자신을 가리키는 순환 참조**가 됐다. 파일은 정상적으로
열리고 값도 들어 있어서(캐시 값이 그대로다) 눈에 띄지 않는다 — 기존 저장 검증을 그냥
통과하던 종류다.

실측: 38↔40 의 367쌍 중 행을 지우는 머지는 15개뿐이지만, 그중 6개 파일에서 935개
수식이 틀어진다. 한 행만 지워도 그 아래 수백 칸이 조용히 어긋난다.

**규칙은 추측으로 정하지 않았다.** 모든 경우를 한 시트에 깔고 엑셀에게 행을 지우게 한
뒤 결과를 읽어 정했다(탐침 52건). 아래 표가 그 정답지 그대로다. 그리고 실제 빌드 파일
25개로 'Excel 이 지운 결과'와 'ExcelMerge 가 지운 결과'를 수식 단위 전수 비교해
전부 일치하는 것을 확인했다.
"""
import os
import zipfile

import openpyxl
import pytest
from lxml import etree

from excelmerge.loaders import clear_values_cache, load_values_any
from excelmerge.xlsx_writer import (_has_self_row_ref, _rewrite_formula_rows,
                                    _self_referencing, _write_patches_to_file)

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# ── 엑셀에게 물어서 받아 적은 정답지 ────────────────────────────────────────
# (지울 행, 원래 수식, 엑셀이 만든 수식)

ORACLE_5_8 = [
    ("B3", "B3", "삭제 위 — 그대로"),
    ("B9", "B7", "삭제 아래 — 두 칸 당김"),
    ("B5", "#REF!", "지워진 행 — 참조 토큰 전체가 #REF!"),
    ("B6", "B5", "두 삭제 사이"),
    ("B8", "#REF!", "지워진 행(둘째)"),
    ("$B$9", "$B$7", "절대 참조도 행 번호는 조정된다"),
    ("B$9", "B$7", "혼합(행 절대)"),
    ("$B9", "$B7", "혼합(열 절대)"),
    ("B17", "B15", "자기보다 아래"),
    ("SUM(B2:B4)", "SUM(B2:B4)", "범위가 전부 삭제 위"),
    ("SUM(B2:B9)", "SUM(B2:B7)", "범위가 삭제를 가로지름"),
    ("SUM(B5:B5)", "SUM(#REF!)", "범위가 통째로 삭제"),
    ("SUM(B5:B8)", "SUM(B5:B6)", "★ 양 끝이 삭제 행이어도 가운데가 남으면 산다"),
    ("SUM(B4:B6)", "SUM(B4:B5)", "범위 가운데가 삭제"),
    ("SUM(B9:B11)", "SUM(B7:B9)", "범위가 전부 삭제 아래"),
    ("SUM(B:B)", "SUM(B:B)", "열 전체 — 행이 없다"),
    ("SUM($B$2:$B$9)", "SUM($B$2:$B$7)", "절대 범위"),
    ("Sheet2!B9", "Sheet2!B9", "다른 시트 — 손대지 않는다"),
    ("SUM(Sheet2!B5:B9)", "SUM(Sheet2!B5:B9)", "다른 시트 범위"),
    ('INDIRECT("B9")', 'INDIRECT("B9")', "문자열 안의 좌표"),
    ('IF(B9=0,"B5","B8")', 'IF(B7=0,"B5","B8")', "문자열은 그대로, 참조만 조정"),
    ("ROW()", "ROW()", "자리 무관"),
    ("ROW()-1", "ROW()-1", "자리 무관(실제 데이터에 많다)"),
    ("B9+B2", "B7+B2", "참조 둘"),
    ('COUNTIF(B2:B9,">1")', 'COUNTIF(B2:B7,">1")', "범위 + 문자열 인자"),
    ("VLOOKUP(B9,B2:B11,2,0)", "VLOOKUP(B7,B2:B9,2,0)", "참조 + 범위"),
    ("OFFSET(B2,1,0)", "OFFSET(B2,1,0)", "함수 인자"),
]

ORACLE_5_6_7 = [
    ("Sheet1!B9", "Sheet1!B6", "★ 자기 시트를 이름으로 써도 조정된다"),
    ("SUM(Sheet1!B2:B9)", "SUM(Sheet1!B2:B6)", "자기 시트 이름 + 범위"),
    ("Sheet1!B5", "Sheet1!#REF!", "★ 시트 접두는 남고 참조만 #REF!"),
    ("SUM(Sheet1!B5:B7)", "SUM(Sheet1!#REF!)", "자기 시트 + 통째 삭제"),
    ("SUM(Sheet1!B4:B8)", "SUM(Sheet1!B4:B5)", "자기 시트 + 일부 생존"),
    ("Sheet2!B5", "Sheet2!B5", "다른 시트는 같은 행번호여도 그대로"),
    ("B2:B9", "B2:B6", "범위를 그대로 돌려주는 식"),
    ("SUM(2:2)", "SUM(2:2)", "행 전체 참조(위)"),
    ("SUM(9:9)", "SUM(6:6)", "★ 행 전체 참조도 조정된다"),
    ("SUM(5:5)", "SUM(#REF!)", "행 전체 참조가 삭제됨"),
    ("SUM(B2:D9)", "SUM(B2:D6)", "여러 열에 걸친 범위"),
    ("IFERROR(B9,0)", "IFERROR(B6,0)", "함수로 감싼 참조"),
    ('B9&"-"&B2', 'B6&"-"&B2', "문자열 연결"),
    ('SUMIF(B2:B9,">1",B2:B9)', 'SUMIF(B2:B6,">1",B2:B6)', "같은 범위 두 번"),
    ("B$9+$B$5", "B$6+#REF!", "혼합 + 삭제된 절대 참조"),
    ("[ext.xlsx]Data!$A$9", "[ext.xlsx]Data!$A$9", "외부 통합문서"),
    ("VLOOKUP(B9,[ext.xlsx]Data!$A$2:$B$30,2,0)",
     "VLOOKUP(B6,[ext.xlsx]Data!$A$2:$B$30,2,0)", "외부는 그대로, 내부만 조정"),
    ("SUM([ext.xlsx]Data!$A$5:$A$7)", "SUM([ext.xlsx]Data!$A$5:$A$7)",
     "외부는 같은 행번호여도 그대로"),
    ("[1]TextUITable!$D$2:$E$9982", "[1]TextUITable!$D$2:$E$9982",
     "실제 데이터가 쓰는 외부 참조 꼴"),
    ("VLOOKUP(E104,[7]메뉴얼!$L$8:$M$15,2,FALSE)",
     "VLOOKUP(E101,[7]메뉴얼!$L$8:$M$15,2,FALSE)",
     "★ 시트 이름이 한글이어도 외부 참조로 알아본다"),
    ("SUM('[7]내 시트'!$L$8:$M$15)", "SUM('[7]내 시트'!$L$8:$M$15)",
     "따옴표로 감싼(공백 포함) 외부 시트명"),
    ("SUM(메뉴얼!$L$8:$M$15)", "SUM(메뉴얼!$L$8:$M$15)", "한글 이름의 다른 시트"),
]


@pytest.mark.parametrize("src, want, why", ORACLE_5_8)
def test_oracle_rows_5_and_8(src, want, why):
    """떨어진 두 행(5, 8)을 지웠을 때 — 엑셀이 만든 결과와 같아야 한다."""
    assert _rewrite_formula_rows(src, {5, 8}, self_sheet="Sheet1") == want, why


@pytest.mark.parametrize("src, want, why", ORACLE_5_6_7)
def test_oracle_rows_5_to_7(src, want, why):
    """붙은 세 행(5~7)을 지웠을 때."""
    assert _rewrite_formula_rows(src, {5, 6, 7}, self_sheet="Sheet1") == want, why


def test_nothing_changes_without_deletions():
    for src, _want, _why in ORACLE_5_8:
        assert _rewrite_formula_rows(src, set(), self_sheet="Sheet1") == src


@pytest.mark.parametrize("formula, want", [
    ("ROW()-1", False),
    ("[1]X!$D:$E", False),
    ("SUM($A:$W)", False),
    ('VLOOKUP(E2,[1]TextUITable!$D:$E,2,0)', True),      # E2 가 자기 시트 참조
    ("Sheet2!B9", False),
    ("Sheet1!B9", True),
    ("B9", True),
    ("", False),
])
def test_has_self_row_ref(formula, want):
    """자리에 무관한 식을 먼저 걸러 내는 판정 — 큰 공유 그룹에서 일을 크게 줄인다."""
    assert _has_self_row_ref(formula, "Sheet1") is want


# ── 저장 경로 ───────────────────────────────────────────────────────────────

SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:C8"/><sheetData>
<row r="1"><c r="A1" t="str"><v>ID</v></c><c r="B1"><v>100</v></c></row>
<row r="2"><c r="A2" t="str"><v>r2</v></c><c r="B2"><v>200</v></c></row>
<row r="3"><c r="A3" t="str"><v>r3</v></c><c r="B3"><f>B2+1</f><v>201</v></c></row>
<row r="4"><c r="A4" t="str"><v>r4</v></c><c r="B4"><f>SUM(B1:B3)</f><v>501</v></c></row>
<row r="5"><c r="A5" t="str"><v>r5</v></c><c r="B5"><f>Data!B2+1</f><v>201</v></c></row>
<row r="6"><c r="A6" t="str"><v>r6</v></c><c r="B6"><f>Other!B2+1</f><v>9</v></c></row>
<row r="7"><c r="A7" t="str"><v>r7</v></c><c r="B7"><f>ROW()-1</f><v>6</v></c></row>
<row r="8"><c r="A8" t="str"><v>r8</v></c><c r="B8"><f>[1]Ext!$B$2</f><v>7</v></c></row>
</sheetData></worksheet>"""


@pytest.fixture
def book(tmp_path):
    """시트 이름이 'Data' 인 통합문서 — 자기 시트를 이름으로 쓴 수식을 담는다."""
    p = tmp_path / "a.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.create_sheet("Other")
    for _ in range(8):
        wb["Data"].append(["", ""])
    wb.save(str(p))
    tmp = str(p) + ".build"
    with zipfile.ZipFile(str(p)) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            zout.writestr(n, SHEET.encode("utf-8")
                          if n.endswith("worksheets/sheet1.xml") else zin.read(n))
    os.replace(tmp, str(p))
    clear_values_cache()
    return p


def _formulas(path):
    out = {}
    with zipfile.ZipFile(str(path)) as z:
        for c in etree.fromstring(
                z.read("xl/worksheets/sheet1.xml")).iter(NS + "c"):
            f = c.find(NS + "f")
            if f is not None:
                out[c.get("r")] = (f.text or "").strip()
    return out


def _save(path, **kw):
    clear_values_cache()
    _write_patches_to_file(str(path), kw.pop("patches", {}), **kw)


def test_a_plain_formula_follows_the_deletion(book):
    """신고된 사고 그 자체 — 'B3: =B2+1' 이 B2 로 당겨지면 순환 참조가 됐었다."""
    _save(book, delete_row_nums={1})
    f = _formulas(book)
    assert f["B2"] == "B1+1", f"본문이 안 따라왔다: {f}"


def test_a_range_follows_the_deletion(book):
    _save(book, delete_row_nums={1})
    assert _formulas(book)["B3"] == "SUM(B1:B2)", _formulas(book)


def test_the_own_sheet_name_is_recognised(book):
    """자기 시트를 이름으로 쓴 참조도 조정된다 — 실측 18개 파일 85,429건이 이 꼴이다."""
    _save(book, delete_row_nums={1})
    assert _formulas(book)["B4"] == "Data!B1+1", _formulas(book)


def test_other_sheets_and_external_books_are_left_alone(book):
    _save(book, delete_row_nums={1})
    f = _formulas(book)
    assert f["B5"] == "Other!B2+1", f
    assert f["B7"] == "[1]Ext!$B$2", f


def test_a_position_independent_formula_is_untouched(book):
    _save(book, delete_row_nums={1})
    assert _formulas(book)["B6"] == "ROW()-1", _formulas(book)


def test_a_reference_to_a_deleted_row_becomes_ref_error(book):
    _save(book, delete_row_nums={2})
    f = _formulas(book)
    assert f["B2"] == "#REF!+1", f


def test_the_values_are_untouched(book):
    clear_values_cache()
    before = load_values_any(str(book))
    _save(book, delete_row_nums={1})
    clear_values_cache()
    after = load_values_any(str(book))
    assert [r[0] for r in after] == [r[0] for r in before[1:]], "값이 밀렸다"


# ── 가드가 살아 있나 ────────────────────────────────────────────────────────
# 값만 보는 검증이 이 사고를 통과시켰다. 재작성을 꺼 놓고 저장을 시도한다.

def test_the_verifier_blocks_a_new_circular_reference(book, monkeypatch):
    import excelmerge.xlsx_writer as W
    monkeypatch.setattr(W, "_rewrite_rows_in_formulas", lambda *a, **k: None)
    clear_summary = _formulas(book)
    with pytest.raises(ValueError) as e:
        _save(book, delete_row_nums={1})
    assert "순환 참조" in str(e.value), str(e.value)
    assert _formulas(book) == clear_summary, "원본이 덮였다"
    assert not os.path.exists(str(book) + ".tmp_merge"), "임시 파일이 남았다"


def test_a_same_row_neighbour_is_not_circular(tmp_path):
    """같은 **행**의 옆 칸을 쓰는 건 순환이 아니다 — 자기 **셀**일 때만 순환이다.

    처음엔 '자기 행을 가리키면 순환' 으로 짰다가 실제 데이터에서 걸렸다.
    D36 의 '"UITable.Currency_"&A36' 같은 꼴이 흔해서, 그렇게 보면 멀쩡한 저장이
    줄줄이 막힌다(실측: Data_CurrencyETC_CS.xlsx, Data_EventMain_CS.xlsx).
    """
    p = tmp_path / "neighbour.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    for _ in range(5):
        wb.active.append(["", "", ""])
    wb.save(str(p))
    sheet = ("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"""
             """<worksheet xmlns="http://schemas.openxmlformats.org/"""
             """spreadsheetml/2006/main"><sheetData>"""
             """<row r="1"><c r="A1" t="str"><v>x</v></c></row>"""
             """<row r="2"><c r="A2" t="str"><v>k</v></c>"""
             """<c r="D2" t="str"><f>"UITable."&amp;A2</f><v>UITable.k</v></c></row>"""
             """<row r="3"><c r="A3" t="str"><v>k3</v></c>"""
             """<c r="D3" t="str"><f>"UITable."&amp;A3</f><v>UITable.k3</v></c></row>"""
             """</sheetData></worksheet>""")
    tmp = str(p) + ".b"
    with zipfile.ZipFile(str(p)) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            zout.writestr(n, sheet.encode("utf-8")
                          if n.endswith("worksheets/sheet1.xml") else zin.read(n))
    os.replace(tmp, str(p))
    assert _self_referencing(str(p), "Data") == set(), "같은 행 옆 칸을 순환으로 봤다"
    _save(p, delete_row_nums={1})                 # 막히지 않고 저장돼야 한다
    assert _formulas(p)["D1"] == '"UITable."&A1', _formulas(p)


def test_a_whole_row_reference_in_its_own_row_is_circular():
    """=SUM(5:5) 가 5행에 있으면 자기를 포함한다 — 행 전체 참조도 봐야 한다."""
    from excelmerge.xlsx_writer import _ref_covers
    assert _ref_covers("5:5", "B5") is True
    assert _ref_covers("5:5", "B6") is False
    assert _ref_covers("B2:D9", "C5") is True
    assert _ref_covers("B2:D9", "E5") is False
    assert _ref_covers("B5", "B5") is True
    assert _ref_covers("A5", "B5") is False


def test_an_existing_circular_reference_is_not_blamed(tmp_path):
    """원래 있던 순환은 막지 않는다 — 성한 파일만 들어온다고 가정하면 안 된다."""
    p = tmp_path / "circ.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    for _ in range(4):
        wb.active.append(["", ""])
    wb.save(str(p))
    sheet = SHEET.replace("<f>B2+1</f>", "<f>B3+1</f>")      # B3 가 자기를 가리킨다
    tmp = str(p) + ".b"
    with zipfile.ZipFile(str(p)) as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            zout.writestr(n, sheet.encode("utf-8")
                          if n.endswith("worksheets/sheet1.xml") else zin.read(n))
    os.replace(tmp, str(p))
    assert _self_referencing(str(p), "Data"), "전제: 원래 순환이 있다"
    _save(p, delete_row_nums={8})                 # 예외 없이 저장돼야 한다
    assert _formulas(p)["B3"] == "B3+1", "원래 있던 순환까지 건드렸다"
