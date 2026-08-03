# -*- coding: utf-8 -*-
"""OOXML 시트 경로 해석(단일 출처 ooxml 모듈) 회귀 테스트.

핵심 관심사: **읽기와 쓰기가 같은 해석을 쓰는가**, 그리고 쓰기 경로가 스펙 위반
워크북(<sheets> 밖의 떠돌이 <sheet>)을 따라가 **엉뚱한 시트에 저장하지 않는가**.
"""
import os
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest

from excelmerge import ooxml, loaders, xlsx_writer

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WS = REL + "/worksheet"


def _wb(sheets, active=None, stray=None):
    bv = f'<bookViews><bookView activeTab="{active}"/></bookViews>' if active is not None else ""
    inner = "".join(f'<sheet name="{n}" sheetId="{i+1}" r:id="{r}"/>'
                    for i, (n, r) in enumerate(sheets))
    out = f'<extra><sheet name="{stray[0]}" sheetId="9" r:id="{stray[1]}"/></extra>' if stray else ""
    return (f'<?xml version="1.0"?><workbook xmlns="{MAIN}" xmlns:r="{REL}">'
            f'{bv}{out}<sheets>{inner}</sheets></workbook>').encode()


def _rels(items):
    r = "".join(f'<Relationship Id="{i}" Target="{t}" Type="{ty}"/>' for i, t, ty in items)
    return f'<?xml version="1.0"?><Relationships xmlns="{RELS_NS}">{r}</Relationships>'.encode()


def _book(tmp_path, wb, rels, name="b.xlsx"):
    p = tmp_path / name
    with zipfile.ZipFile(p, "w") as z:
        if wb is not None:
            z.writestr("xl/workbook.xml", wb)
        if rels is not None:
            z.writestr("xl/_rels/workbook.xml.rels", rels)
        z.writestr("xl/worksheets/sheet1.xml", b"<worksheet/>")
        z.writestr("xl/worksheets/sheet2.xml", b"<worksheet/>")
    return str(p)


@pytest.mark.parametrize("target,expect", [
    ("worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"),     # 워크북 기준 상대
    ("/xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"),  # 패키지 루트 절대
    ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"),   # 이미 정규형
])
def test_rel_target_normalization(target, expect):
    assert ooxml.rel_target_to_path(target) == expect


def test_read_and_write_resolve_the_same_sheet(tmp_path):
    """같은 이름에 대해 읽기와 쓰기가 동일 경로로 해석돼야 한다(갈라지면 오저장)."""
    p = _book(tmp_path,
              _wb([("Sheet1", "rId1"), ("Data", "rId2")]),
              _rels([("rId1", "worksheets/sheet1.xml", WS),
                     ("rId2", "worksheets/sheet2.xml", WS)]))
    with zipfile.ZipFile(p) as z:
        for name in ("Sheet1", "Data"):
            assert (loaders._sheet_xml_path_in_zip(z, name)
                    == xlsx_writer._resolve_sheet_path(z, name))
        assert loaders._sheet_xml_path_in_zip(z, "Data") == "xl/worksheets/sheet2.xml"


def test_write_refuses_missing_sheet_read_falls_back(tmp_path):
    """정책 차이는 **의도된 설계**: 읽기는 첫 시트 폴백, 쓰기는 raise(오저장 방지)."""
    p = _book(tmp_path, _wb([("Sheet1", "rId1")]),
              _rels([("rId1", "worksheets/sheet1.xml", WS)]))
    with zipfile.ZipFile(p) as z:
        assert loaders._sheet_xml_path_in_zip(z, "NoSuch") == "xl/worksheets/sheet1.xml"
        with pytest.raises(ValueError):
            xlsx_writer._resolve_sheet_path(z, "NoSuch")


def test_stray_sheet_outside_container_is_not_writable(tmp_path):
    """<sheets> 밖의 떠돌이 <sheet> 는 쓰기 대상이 될 수 없다.

    통합 전에는 쓰기 경로가 문서 전체를 iter() 로 훑어 이 떠돌이 항목을 주웠고,
    activeTab/이름 해석이 그것을 가리키면 **엉뚱한 파트에 저장**될 수 있었다.
    """
    p = _book(tmp_path,
              _wb([("Sheet1", "rId1")], active=0, stray=("Hidden", "rId2")),
              _rels([("rId1", "worksheets/sheet1.xml", WS),
                     ("rId2", "worksheets/sheet2.xml", WS)]))
    with zipfile.ZipFile(p) as z:
        # 떠돌이 이름으로는 저장할 수 없다
        assert xlsx_writer._find_sheet_path_by_name(z, "Hidden") is None
        with pytest.raises(ValueError):
            xlsx_writer._resolve_sheet_path(z, "Hidden")
        # 이름 미지정(activeTab) 경로도 컨테이너 안의 시트로 해석돼야 한다
        assert xlsx_writer._resolve_sheet_path(z, None) == "xl/worksheets/sheet1.xml"


def test_write_requires_worksheet_type_read_is_lenient(tmp_path):
    """쓰기는 worksheet 타입 rel 만 인정(엄격), 읽기는 기존 관용 동작 유지."""
    chart = REL + "/chartsheet"
    p = _book(tmp_path, _wb([("Sheet1", "rId1")]),
              _rels([("rId1", "worksheets/sheet1.xml", chart)]))
    with zipfile.ZipFile(p) as z:
        assert xlsx_writer._find_sheet_path_by_name(z, "Sheet1") is None
        assert loaders._sheet_xml_path_in_zip(z, "Sheet1") == "xl/worksheets/sheet1.xml"


def test_active_tab_selects_and_clamps(tmp_path):
    rels = _rels([("rId1", "worksheets/sheet1.xml", WS),
                  ("rId2", "worksheets/sheet2.xml", WS)])
    sheets = [("Sheet1", "rId1"), ("Data", "rId2")]
    with zipfile.ZipFile(_book(tmp_path, _wb(sheets, active=1), rels, "a1.xlsx")) as z:
        assert ooxml.active_sheet_path(z) == "xl/worksheets/sheet2.xml"
    # 범위를 넘으면 마지막 시트로 clamp
    with zipfile.ZipFile(_book(tmp_path, _wb(sheets, active=9), rels, "a9.xlsx")) as z:
        assert ooxml.active_sheet_path(z) == "xl/worksheets/sheet2.xml"


@pytest.mark.parametrize("wb,rels", [
    (None, _rels([("rId1", "worksheets/sheet1.xml", WS)])),   # workbook.xml 없음
    (b"<workbook not xml", None),                              # 깨진 XML
    (_wb([]), _rels([])),                                      # 시트 0개
    (_wb([("Sheet1", "rId1")]), None),                         # rels 없음
])
def test_degenerate_books_do_not_raise_on_read(tmp_path, wb, rels):
    """손상/비정상 워크북에서도 읽기 경로는 예외 없이 관례적 경로로 폴백해야 한다."""
    p = _book(tmp_path, wb, rels, "deg.xlsx")
    with zipfile.ZipFile(p) as z:
        assert loaders._sheet_xml_path_in_zip(z, "Sheet1") == ooxml.DEFAULT_SHEET_PATH
        assert xlsx_writer._find_active_sheet_path(z) == ooxml.DEFAULT_SHEET_PATH
