# -*- coding: utf-8 -*-
"""안전 수정 회귀 테스트 (Qt 불필요 — 순수 로직/저장 경로).

실행: python tests/test_robustness.py  (또는 pytest)
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

import json

from excelmerge.xlsx_writer import _write_patches_to_file
from excelmerge.diff_engine import count_dropped_key_rows, count_changed
from excelmerge.folder_compare import content_equal
from excelmerge.loaders import (
    load_formula_flags_any, load_values_any, load_json_as_matrix,
)


def _make_xlsx(path, rows, sheet="Sheet1"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for r in rows:
        ws.append(r)
    wb.save(path)


def test_save_roundtrip_and_bak():
    """값 패치 저장 → 재로드 시 반영 + 원본 .bak 백업 생성 (A2)."""
    d = tempfile.mkdtemp()
    p = os.path.join(d, "t.xlsx")
    _make_xlsx(p, [["ID", "V"], ["1", "a1"], ["2", "a2"]], sheet="Data")
    _write_patches_to_file(p, {"B2": "PATCHED"}, sheet_name="Data")
    vals = load_values_any(p, None, "Data")
    assert vals[1][1] == "PATCHED", vals[1]
    assert os.path.exists(p + ".bak"), ".bak 백업 미생성"
    # 백업은 패치 전 원본이어야 함
    bak = load_values_any(p + ".bak", None, "Data")
    assert bak[1][1] == "a1", bak[1]
    print("PASS test_save_roundtrip_and_bak")


def test_save_missing_sheet_raises():
    """존재하지 않는 시트명으로 저장하면 조용히 다른 시트에 쓰지 않고 raise (A1)."""
    d = tempfile.mkdtemp()
    p = os.path.join(d, "t.xlsx")
    _make_xlsx(p, [["ID", "V"], ["1", "a1"]], sheet="Data")
    raised = False
    try:
        _write_patches_to_file(p, {"B2": "x"}, sheet_name="NoSuchSheet")
    except ValueError:
        raised = True
    assert raised, "없는 시트인데 raise 안 함 (엉뚱한 시트에 쓸 위험)"
    print("PASS test_save_missing_sheet_raises")


def test_count_dropped_key_rows():
    """공백 키 + 중복 키 본문 행 카운트 (A5/H2)."""
    a = [["ID", "V"], ["1", "a"], ["1", "dup"], ["", "blank"], ["2", "b"]]
    b = [["ID", "V"], ["1", "a"], ["2", "b"]]
    # a: dup('1') 1 + blank 1 = 2, b: 0
    assert count_dropped_key_rows(a, b, 0) == 2, count_dropped_key_rows(a, b, 0)
    # ROW 순서(-1)는 드롭 없음
    assert count_dropped_key_rows(a, b, -1) == 0
    # 유니크 키만 있으면 0
    assert count_dropped_key_rows(b, b, 0) == 0
    print("PASS test_count_dropped_key_rows")


def test_count_changed():
    """변경 셀 수 카운트 — 'same' 제외, 제외 열도 제외."""
    matrix = [
        [("same",), ("modified",)],
        [("added",), ("same",)],
    ]
    assert count_changed(matrix) == 2, count_changed(matrix)            # modified + added
    assert count_changed(matrix, excluded_cols={1}) == 1               # col1(modified) 제외 → added만
    assert count_changed([]) == 0
    print("PASS test_count_changed")


def test_content_equal_xlsx():
    """바이트가 달라도 셀 값이 같으면 content_equal True, 값이 다르면 False."""
    d = tempfile.mkdtemp()
    p1 = os.path.join(d, "a.xlsx")
    p2 = os.path.join(d, "b.xlsx")
    p3 = os.path.join(d, "c.xlsx")
    _make_xlsx(p1, [["ID", "V"], ["1", "x"]], sheet="S")
    _make_xlsx(p2, [["ID", "V"], ["1", "x"]], sheet="S")   # 같은 값, 다른 바이트(별도 저장)
    _make_xlsx(p3, [["ID", "V"], ["1", "y"]], sheet="S")   # 값 다름
    assert content_equal(p1, p2) is True, "동일 값인데 다름으로 판정"
    assert content_equal(p1, p3) is False, "다른 값인데 동일로 판정"
    print("PASS test_content_equal_xlsx")


def test_formula_flags_xml():
    """수식 셀 좌표를 시트 XML <f> 스캔으로 정확히 반환."""
    d = tempfile.mkdtemp()
    p = os.path.join(d, "f.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "V", "Calc"])
    ws.append([1, 10, "=B2*2"])   # C2 = (row0=1, col0=2)
    ws.append([2, 20, 30])        # 리터럴
    wb.save(p)
    flags = load_formula_flags_any(p)
    assert (1, 2) in flags, f"C2 수식 미검출: {flags}"
    assert (2, 2) not in flags, "리터럴을 수식으로 오검출"
    # json/uasset은 빈 집합
    assert load_formula_flags_any(os.path.join(d, "x.json")) == set()
    print("PASS test_formula_flags_xml")


def test_json_wrapper_renders_as_table():
    """래퍼 객체({"K":[{...}]})도 최상위 객체 배열([{...}])과 '동일한 표'로 로딩되고,
    비표형 JSON은 [path,value] 폴백을 유지한다."""
    d = tempfile.mkdtemp()
    records = [{"ID": 1, "Name": "a"}, {"ID": 2, "Name": "b"}]
    p_arr = os.path.join(d, "arr.json")     # 최상위 배열 (AbyssAttack 형)
    p_wrap = os.path.join(d, "wrap.json")   # 래퍼 객체 (TablePackage 형)
    p_scalar = os.path.join(d, "scalar.json")
    with open(p_arr, "w", encoding="utf-8") as f:
        json.dump(records, f)
    with open(p_wrap, "w", encoding="utf-8") as f:
        json.dump({"TablePackage": records}, f)
    with open(p_scalar, "w", encoding="utf-8") as f:
        json.dump({"meta": {"v": 1}}, f)   # 객체 배열 아님 → 폴백

    m_arr = load_json_as_matrix(p_arr)
    m_wrap = load_json_as_matrix(p_wrap)
    assert m_arr == [["ID", "Name"], ["1", "a"], ["2", "b"]], m_arr
    assert m_wrap == m_arr, f"래퍼 JSON이 배열 JSON과 다르게 로딩됨: {m_wrap}"

    m_scalar = load_json_as_matrix(p_scalar)
    assert m_scalar[0] == ["path", "value"], f"비표형 폴백 깨짐: {m_scalar[0]}"
    print("PASS test_json_wrapper_renders_as_table")


def main():
    test_save_roundtrip_and_bak()
    test_json_wrapper_renders_as_table()
    test_save_missing_sheet_raises()
    test_count_dropped_key_rows()
    test_count_changed()
    test_content_equal_xlsx()
    test_formula_flags_xml()
    print("ALL ROBUSTNESS TESTS PASS")


if __name__ == "__main__":
    main()
