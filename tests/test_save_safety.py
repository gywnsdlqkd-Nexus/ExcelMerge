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
    _is_file_locked, _is_numeric, _promote_empty_cols_to_delete,
    _write_patches_to_file,
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


# ── 행 삭제가 섞인 저장 — 자리가 밀려도 검증이 오작동하면 안 된다 ────────────
# 실사용 신고: 병합 저장을 누르니 "저장 검증 실패" 로 막혔다. 원인은 저장 자체가 아니라
# 검증이었다 — _renumber_after_delete 가 VBA .Delete 처럼 삭제 행 아래를 위로 당기는데,
# 검증이 저장 전 좌표를 그대로 보고 한 칸 아래 값과 비교했다.

def test_patch_below_a_deleted_row_is_verified_at_its_new_place(tmp_path):
    p = _make_xlsx(tmp_path / "a.xlsx",
                   [[f"k{i}", f"v{i}", f"n{i}"] for i in range(1, 7)])
    _write_patches_to_file(p, {"B4": "새값"}, delete_row_nums={2})
    rows = _values(p)
    assert len(rows) == 5, "행이 하나 지워져야 한다"
    assert rows[2] == ["k4", "새값", "n4"], f"패치가 엉뚱한 자리에 갔다: {rows}"


def test_many_deleted_rows_shift_correctly(tmp_path):
    p = _make_xlsx(tmp_path / "a.xlsx",
                   [[f"k{i}", f"v{i}"] for i in range(1, 9)])
    _write_patches_to_file(p, {"B7": "일곱", "B8": "여덟"}, delete_row_nums={2, 5})
    rows = _values(p)
    assert [r[0] for r in rows] == ["k1", "k3", "k4", "k6", "k7", "k8"]
    assert rows[4] == ["k7", "일곱"] and rows[5] == ["k8", "여덟"]


def test_patch_on_a_deleted_row_is_not_flagged(tmp_path):
    """지워진 행에 남은 패치는 확인할 자리가 없다 — 거짓 실패를 내면 안 된다."""
    p = _make_xlsx(tmp_path / "a.xlsx", [[f"k{i}", f"v{i}"] for i in range(1, 5)])
    _write_patches_to_file(p, {"B2": "사라질값"}, delete_row_nums={2})
    rows = _values(p)
    assert [r[0] for r in rows] == ["k1", "k3", "k4"]


def test_verification_still_catches_a_bad_write_with_deletes(tmp_path, monkeypatch):
    """자리 보정을 넣어도 진짜 오기록은 여전히 잡아야 한다."""
    import excelmerge.xlsx_writer as xw
    p = _make_xlsx(tmp_path / "a.xlsx", [[f"k{i}", f"v{i}"] for i in range(1, 6)])
    before = open(p, "rb").read()
    real = xw._patch_sheet_xml
    monkeypatch.setattr(
        xw, "_patch_sheet_xml",
        lambda data, patches, *a, **k: real(data, {r: "엉뚱" for r in patches}, *a, **k))
    with pytest.raises(ValueError, match="저장 검증 실패"):
        _write_patches_to_file(p, {"B4": "새값"}, delete_row_nums={2})
    assert open(p, "rb").read() == before, "검증 실패인데 원본이 바뀌었다"


def test_deleted_column_does_not_shift_verification(tmp_path):
    """열 삭제는 <c> 만 지우고 열 문자는 그대로라 좌표가 밀리지 않는다."""
    p = _make_xlsx(tmp_path / "a.xlsx", [["k1", "v1", "n1"], ["k2", "v2", "n2"]])
    _write_patches_to_file(p, {"C2": "끝값"}, delete_col_letters={"B"})
    rows = _values(p)
    assert rows[1][0] == "k2" and rows[1][2] == "끝값"
    assert rows[1][1] == "", "지운 열 자리는 비어야 한다"


# ── 한쪽에만 있는 행 지우기 — '결과가 빈 수식' 이 행 삭제를 막으면 안 된다 ────
# 실사용 신고: Data_HelpPopUp_C.xlsx 에서 TID 245101~245103 을 A→B 로 지웠더니
# 245102 자리에 **빈 행**이 남았다. 흐름은 이렇다.
#   · A 에 없는 행이라 모든 셀이 '추가'로 잡히지만, 양쪽 다 빈 셀은 '같음'이라
#     병합 준비에서 빠진다(staging.stageable_cells) → 그 셀엔 패치가 안 붙는다.
#   · 그 셀이 마침 결과가 빈 수식(`<f>…</f><v/>`)이면, 수식 원문을 값으로 치는
#     행-비었나 판정이 '아직 내용 있음'으로 읽어 행 삭제 승격을 막는다.
#   · 결국 나머지 열만 빈값으로 덮여, 눈에는 아무것도 없는 행이 남는다.
# 그래서 행 판정은 **보이는 값**(수식이면 계산 결과)으로 한다. 열 판정은 그대로다.

def _formula_cell_xlsx(path, rows, formulas):
    """formulas: {"C2": "IF(1=1,\"\",\"x\")"} — <f> 만 있고 캐시값이 없는 셀."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    for ref, f in formulas.items():
        ws[ref] = "=" + f
    wb.save(str(path))
    return str(path)


def test_empty_formula_does_not_block_deleting_the_row(tmp_path):
    p = _formula_cell_xlsx(tmp_path / "a.xlsx",
                           [["k1", "v1", None], ["k2", "v2", None]],
                           {"C2": 'IF(1=1,"","x")'})
    # A 에 없는 행이라 A·B 열은 빈값으로 덮이고, C 열(빈 수식)은 패치가 없다.
    patches = {"A2": "", "B2": ""}
    new_patches, deletes, _cols = _promote_empty_cols_to_delete(patches, set(), p)
    assert deletes == {2}, "결과가 빈 수식 때문에 행 삭제가 막혔다"
    assert new_patches == {}, "삭제된 행의 패치가 남아 있다"


def test_formula_with_a_cached_value_still_keeps_the_row(tmp_path):
    """수식이 실제로 값을 내고 있으면 그 행은 비어 있지 않다 — 지우면 안 된다."""
    p = _make_xlsx(tmp_path / "a.xlsx", [["k1", "v1", "x"], ["k2", "v2", "남을값"]])
    new_patches, deletes, _cols = _promote_empty_cols_to_delete(
        {"A2": "", "B2": ""}, set(), p)
    assert deletes == set(), "내용이 남아 있는 행을 지웠다"
    assert new_patches == {"A2": "", "B2": ""}


def test_empty_formula_column_is_still_not_deleted(tmp_path):
    """열 판정은 수식 원문을 본다 — 지금 결과가 비었다고 수식 열을 지우면 안 된다."""
    p = _formula_cell_xlsx(tmp_path / "a.xlsx",
                           [["k1", "v1", None], ["k2", "v2", None]],
                           {"C1": 'IF(1=1,"","x")', "C2": 'IF(1=1,"","x")'})
    _new, _del, cols = _promote_empty_cols_to_delete({"A2": ""}, set(), p)
    assert "C" not in cols, "수식이 든 열을 통째로 지우려 했다"


def test_cached_empty_value_element_counts_as_empty(tmp_path):
    """실제 파일의 모양은 `<f>…</f><v/>` 였다 — <v> 가 있되 내용이 없는 경우."""
    import zipfile
    src = _formula_cell_xlsx(tmp_path / "src.xlsx",
                             [["k1", "v1", None], ["k2", "v2", None]],
                             {"C2": 'IF(1=1,"","x")'})
    dst = str(tmp_path / "a.xlsx")
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                text = text.replace("<v></v>", "<v/>")
                assert "<v/>" in text, "테스트 준비 실패: <v/> 를 넣지 못했다"
                data = text.encode("utf-8")
            zout.writestr(item, data)
    _new, deletes, _cols = _promote_empty_cols_to_delete({"A2": "", "B2": ""}, set(), dst)
    assert deletes == {2}


def test_the_whole_row_really_disappears_after_the_promotion(tmp_path):
    """승격 → 실제 저장까지: 빈 행이 남지 않고 아래 행이 위로 당겨져야 한다."""
    p = _formula_cell_xlsx(
        tmp_path / "a.xlsx",
        [["k1", "v1", None], ["k2", "v2", None], ["k3", "v3", None]],
        {"C2": 'IF(1=1,"","x")'})
    new_patches, deletes, cols = _promote_empty_cols_to_delete(
        {"A2": "", "B2": ""}, set(), p)
    _write_patches_to_file(p, new_patches, delete_row_nums=deletes,
                           delete_col_letters=cols)
    rows = _values(p)
    assert [r[0] for r in rows] == ["k1", "k3"], f"빈 행이 남았다: {rows}"
    assert all(any(str(c).strip() for c in r) for r in rows), f"빈 행이 남았다: {rows}"
