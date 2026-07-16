# -*- coding: utf-8 -*-
"""CSV/TSV 네이티브 비교(A안: 로딩·비교 전용) — 로더 동작 검증."""
from excelmerge.loaders import (
    load_csv_as_matrix, load_values_any, list_sheet_names,
    load_formula_flags_any, clear_values_cache, SUPPORTED_EXTS,
)
from excelmerge.diff_engine import compute_diff


def test_supported_exts_include_csv_tsv():
    assert ".csv" in SUPPORTED_EXTS and ".tsv" in SUPPORTED_EXTS


def test_csv_comma(tmp_path):
    p = tmp_path / "a.csv"
    p.write_text("ID,Name\n1,alpha\n2,beta\n", encoding="utf-8")
    m = load_csv_as_matrix(str(p))
    assert m == [["ID", "Name"], ["1", "alpha"], ["2", "beta"]]


def test_tsv_tab(tmp_path):
    p = tmp_path / "a.tsv"
    p.write_text("ID\tName\n1\talpha\n", encoding="utf-8")
    assert load_csv_as_matrix(str(p)) == [["ID", "Name"], ["1", "alpha"]]


def test_cp949_encoding(tmp_path):
    p = tmp_path / "kr.csv"
    p.write_bytes("ID,이름\n1,가나다\n".encode("cp949"))   # 한국어 cp949
    m = load_csv_as_matrix(str(p))
    assert m == [["ID", "이름"], ["1", "가나다"]]


def test_quoted_and_ragged(tmp_path):
    p = tmp_path / "q.csv"
    # 따옴표 안의 쉼표/개행 + 열 수 불균일 (개행 변환 없이 정확히 쓰려고 bytes 사용)
    p.write_bytes(b'a,b,c\n"x,y","line1\nline2"\nsolo\n')
    m = load_csv_as_matrix(str(p))
    assert m[0] == ["a", "b", "c"]
    assert m[1] == ["x,y", "line1\nline2", ""]      # 3열로 패딩
    assert m[2] == ["solo", "", ""]


def test_load_values_any_dispatches_csv(tmp_path):
    p = tmp_path / "v.csv"
    p.write_text("K,V\n1,a\n", encoding="utf-8")
    clear_values_cache()
    assert load_values_any(str(p)) == [["K", "V"], ["1", "a"]]
    assert list_sheet_names(str(p)) == []          # 단일 시트 → 탭 숨김
    assert load_formula_flags_any(str(p)) == set()  # 수식 개념 없음


def test_compare_two_csv(tmp_path):
    a = tmp_path / "a.csv"; a.write_text("ID,V\n1,x\n2,y\n", encoding="utf-8")
    b = tmp_path / "b.csv"; b.write_text("ID,V\n1,x\n2,Z\n", encoding="utf-8")
    ma = load_values_any(str(a)); mb = load_values_any(str(b))
    m, meta = compute_diff(ma, mb, key_col=0, key_row=0)
    # 2행(ID=2)의 V열이 modified
    assert m[2][1][0] == "modified"
