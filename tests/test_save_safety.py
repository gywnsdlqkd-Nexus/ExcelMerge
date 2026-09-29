# -*- coding: utf-8 -*-
"""저장이 값을 바꾸지 않는다 — 숫자/문자 판정 + 교체 전 검증.

두 가지를 고정한다.

  1. **숫자 판정**. 예전엔 float() 로 파싱만 되면 숫자 셀로 썼다. 그래서 '01'·'007'·
     끝에 탭이 붙은 '210000\\t'·'nan' 이 숫자가 되어, 저장하면 값이 조용히 바뀌었다
     (실측: 실제 빌드 파일 사본에 '01' 을 저장하면 '1' 이 남았다).
  2. **교체 전 검증**. 백업(.bak)을 만들지 않으므로, 잘못 쓴 파일로 원본을 덮으면
     되돌릴 수 없다. 임시 파일을 되읽어 확인한 뒤에만 교체한다.
"""
import openpyxl
import pytest

from excelmerge.loaders import load_values_any, clear_values_cache, _cell_to_str
from excelmerge.xlsx_writer import (
    _is_file_locked, _is_numeric, _write_patches_to_file,
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


# ── 1. 숫자로 써도 되는 값 / 문자로 지켜야 하는 값 ───────────────────────────

@pytest.mark.parametrize("val", ["0", "1", "123", "-4", "1.5", "-2.75", "0.1",
                                 "1000000", "-0.5"])
def test_plain_numbers_stay_numeric(val):
    assert _is_numeric(val), f"{val!r} 은 숫자로 써도 글자가 그대로다"


@pytest.mark.parametrize("val", [
    "01",          # 앞자리 0 — 숫자로 쓰면 '1' 이 된다(실제 데이터에 49개)
    "007",
    "00",
    "1.0",         # 되읽으면 '1' — 보이는 글자가 달라진다
    "1e5",         # 되읽으면 '100000'
    "210000\t",    # 끝에 탭 — 실제 데이터에 4개
    " 1",
    "1 ",
    "nan", "inf", "-inf",
    "1_0",         # 파이썬만 숫자로 보는 표기
    "-0",          # 되읽으면 '0'
    "+1",          # 되읽으면 '1'
    "",
])
def test_ambiguous_values_are_kept_as_text(val):
    assert not _is_numeric(val), f"{val!r} 을 숫자로 쓰면 값이 바뀐다"


def test_numeric_rule_matches_the_loader(tmp_path):
    """판정 기준(되읽기 표기)이 로더의 표기 규칙과 어긋나지 않아야 한다.

    _is_numeric 은 loaders._cell_to_str 과 같은 규칙을 가정한다 — 한쪽만 바뀌면
    '숫자로 썼는데 되읽으니 다른 글자'가 되어 저장이 값을 바꾼다.
    """
    for raw in (0, 1, 123, -4, 1.5, -2.75, 0.1, 1000000.0):
        text = _cell_to_str(raw)
        assert _is_numeric(text), f"{raw!r} → {text!r} 는 숫자로 쓸 수 있어야 한다"


def test_leading_zero_survives_a_real_save(tmp_path):
    """보고된 실제 사례 — '01' 을 저장하면 '1' 이 되던 버그."""
    p = _make_xlsx(tmp_path / "a.xlsx", [["ID", "CNT"], ["k1", "01"]])
    _write_patches_to_file(p, {"B2": "01"})
    assert _values(p)[1][1] == "01"


def test_tab_suffix_survives_a_real_save(tmp_path):
    p = _make_xlsx(tmp_path / "a.xlsx", [["ID", "V"], ["k1", "x"]])
    _write_patches_to_file(p, {"B2": "210000\t"})
    assert _values(p)[1][1] == "210000\t"


def test_plain_number_is_still_written_as_a_number(tmp_path):
    """비회귀 — 평범한 숫자는 예전처럼 숫자 셀이어야 한다(표시형식·수식 계산 유지)."""
    p = _make_xlsx(tmp_path / "a.xlsx", [["ID", "V"], ["k1", "x"]])
    _write_patches_to_file(p, {"B2": "1234"})
    ws = openpyxl.load_workbook(p)["Sheet1"]
    assert ws["B2"].value == 1234 and isinstance(ws["B2"].value, int)
    assert _values(p)[1][1] == "1234"


# ── 2. 교체 전 검증 ──────────────────────────────────────────────────────────

def test_bad_write_does_not_touch_the_original(tmp_path, monkeypatch):
    """쓰기가 어긋나면 원본을 건드리지 않는다 — 검증이 교체를 막아야 한다."""
    import excelmerge.xlsx_writer as xw
    p = _make_xlsx(tmp_path / "a.xlsx", [["ID", "V"], ["k1", "원래값"]])
    before = open(p, "rb").read()

    real = xw._patch_sheet_xml

    def sabotage(data, patches, *a, **k):
        # 의도한 값 대신 엉뚱한 값을 쓴다(과거 _is_numeric 버그와 같은 상황).
        return real(data, {ref: "엉뚱한값" for ref in patches}, *a, **k)

    monkeypatch.setattr(xw, "_patch_sheet_xml", sabotage)
    with pytest.raises(ValueError, match="저장 검증 실패"):
        _write_patches_to_file(p, {"B2": "새값"})

    assert open(p, "rb").read() == before, "검증에 실패했는데 원본이 바뀌었다"
    assert not (tmp_path / "a.xlsx.tmp_merge").exists(), "임시 파일이 남았다"


def test_verification_reports_the_cell_and_both_values(tmp_path, monkeypatch):
    import excelmerge.xlsx_writer as xw
    p = _make_xlsx(tmp_path / "a.xlsx", [["ID", "V"], ["k1", "x"]])
    real = xw._patch_sheet_xml
    monkeypatch.setattr(
        xw, "_patch_sheet_xml",
        lambda data, patches, *a, **k: real(data, {r: "틀린값" for r in patches}, *a, **k))
    with pytest.raises(ValueError) as e:
        _write_patches_to_file(p, {"B2": "맞는값"})
    msg = str(e.value)
    assert "B2" in msg and "맞는값" in msg and "틀린값" in msg, msg


def test_good_write_passes_verification(tmp_path):
    p = _make_xlsx(tmp_path / "a.xlsx", [["ID", "V"], ["k1", "x"], ["k2", "y"]])
    _write_patches_to_file(p, {"B2": "새값", "B3": "12"})
    rows = _values(p)
    assert rows[1][1] == "새값" and rows[2][1] == "12"


def test_formula_patch_is_not_flagged(tmp_path):
    """수식은 캐시값이 없어 되읽으면 빈 칸 — 검증이 거짓 실패를 내면 안 된다."""
    p = _make_xlsx(tmp_path / "a.xlsx", [["ID", "V"], ["k1", "x"]])
    _write_patches_to_file(p, {"B2": "=1+2"})
    ws = openpyxl.load_workbook(p)["Sheet1"]
    assert ws["B2"].value == "=1+2"


def test_cleared_cell_passes_verification(tmp_path):
    """빈 값 패치는 셀 제거 — 되읽으면 빈 칸이므로 통과해야 한다."""
    p = _make_xlsx(tmp_path / "a.xlsx", [["ID", "V"], ["k1", "지울값"]])
    _write_patches_to_file(p, {"B2": ""})
    assert _values(p)[1][1] == ""


def test_verification_runs_on_the_named_sheet(tmp_path):
    """시트를 지정해 저장하면 검증도 그 시트를 봐야 한다."""
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "첫번째"
    ws1.append(["ID", "V"]); ws1.append(["k1", "A쪽"])
    ws2 = wb.create_sheet("두번째")
    ws2.append(["ID", "V"]); ws2.append(["k1", "B쪽"])
    p = str(tmp_path / "multi.xlsx")
    wb.save(p)

    _write_patches_to_file(p, {"B2": "바뀐값"}, sheet_name="두번째")
    assert _values(p, "두번째")[1][1] == "바뀐값"
    assert _values(p, "첫번째")[1][1] == "A쪽", "다른 시트가 바뀌었다"


# ── 3. 잠금 확인이 파일을 만들지 않는다 ──────────────────────────────────────

def test_lock_check_does_not_create_a_missing_file(tmp_path):
    missing = tmp_path / "없는파일.xlsx"
    assert _is_file_locked(str(missing)) is False
    assert not missing.exists(), "잠금을 확인만 했는데 빈 파일이 생겼다"


def test_lock_check_reports_a_writable_file_as_free(tmp_path):
    p = _make_xlsx(tmp_path / "a.xlsx", [["ID"], ["k1"]])
    assert _is_file_locked(p) is False
