# -*- coding: utf-8 -*-
"""3파일 병합 저장 — 결과 = 가운데 원본 + 결정. 실제 xlsx 를 쓰고 같은 로더로 되읽어 확인한다.

모두 합성 데이터다(공개 저장소 — 실데이터 금지).
"""
import os

import openpyxl
import pytest
from openpyxl.styles import Font, PatternFill

from excelmerge.diff3_engine import (SRC_LEFT, SRC_RIGHT, auto_merge_plan, compute_diff3,
                                     match_columns3)
from excelmerge.diff3_workers import write_merge3
from excelmerge.loaders import clear_values_cache, load_values_any
from excelmerge.merge3_build import build_merge3_patches, expand_row_removals

H = ["ID", "V", "W"]


def _xlsx(path, rows, styler=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for r in rows:
        ws.append(r)
    if styler:
        styler(ws)
    wb.save(str(path))
    return str(path)


def _vals(path, sheet="Data"):
    clear_values_cache()
    return load_values_any(path, None, sheet)


def _diff(l, b, r):
    data = [_vals(p) for p in (l, b, r)]
    cm = match_columns3(*data, 0)
    m, meta = compute_diff3(*data, 0, 0, cm)
    return m, meta, cm


def _three(tmp_path, left, base, right, **styles):
    return (_xlsx(tmp_path / "l.xlsx", left, styles.get("l")),
            _xlsx(tmp_path / "b.xlsx", base, styles.get("b")),
            _xlsx(tmp_path / "r.xlsx", right, styles.get("r")))


def _save(tmp_path, l, b, r, decisions, m, meta, cm, out=None):
    out = out or str(tmp_path / "result.xlsx")
    write_merge3(b, out, [("Data", m, meta, cm, decisions)], {SRC_LEFT: l, SRC_RIGHT: r})
    return out


def test_non_conflicting_changes_from_both_sides(tmp_path):
    base = [H, [1, "a", "p"], [2, "b", "q"]]
    left = [H, [1, "A", "p"], [2, "b", "q"]]
    right = [H, [1, "a", "p"], [2, "b", "Q"]]
    l, b, r = _three(tmp_path, left, base, right)
    m, meta, cm = _diff(l, b, r)
    before = open(b, "rb").read()
    out = _save(tmp_path, l, b, r, auto_merge_plan(m), m, meta, cm)
    assert _vals(out) == [H, ["1", "A", "p"], ["2", "b", "Q"]]
    assert open(b, "rb").read() == before, "결과 파일을 따로 줬는데 가운데 원본이 바뀌었다"


def test_rows_deleted_and_added_on_each_side(tmp_path):
    base = [H, [1, "a", ""], [2, "b", ""], [3, "c", ""]]
    left = [H, [1, "a", ""], [3, "c", ""], [4, "L4", ""]]      # 2 지움, 4 추가
    right = [H, [1, "a", ""], [2, "b", ""], [3, "c", ""], [5, "R5", ""]]   # 5 추가
    l, b, r = _three(tmp_path, left, base, right)
    m, meta, cm = _diff(l, b, r)
    out = _save(tmp_path, l, b, r, auto_merge_plan(m), m, meta, cm)
    assert _vals(out) == [H, ["1", "a", ""], ["3", "c", ""], ["4", "L4", ""], ["5", "R5", ""]]


def test_delete_versus_edit_resolved_either_way(tmp_path):
    base = [H, [1, "a", "p"], [2, "b", "q"]]
    left = [H, [1, "a", "p"]]                        # 2 지움
    right = [H, [1, "a", "p"], [2, "b", "Q"]]        # 2 고침
    l, b, r = _three(tmp_path, left, base, right)
    m, meta, cm = _diff(l, b, r)
    row2 = next(i for i, mm in enumerate(meta) if mm == (None, 2, 2))
    # 왼쪽(지움)을 고르면 칸 하나만 골라도 행 전체가 따라와 통째로 지워진다
    picked = expand_row_removals(m, meta, {(row2, 2)}, SRC_LEFT)
    out = _save(tmp_path, l, b, r, {k: SRC_LEFT for k in picked}, m, meta, cm,
                str(tmp_path / "res_left.xlsx"))
    assert _vals(out) == [H, ["1", "a", "p"]]
    # 오른쪽(고침)을 고르면 그 값
    out = _save(tmp_path, l, b, r, {(row2, 2): SRC_RIGHT}, m, meta, cm,
                str(tmp_path / "res_right.xlsx"))
    assert _vals(out) == [H, ["1", "a", "p"], ["2", "b", "Q"]]


def test_saving_again_rebuilds_from_the_original(tmp_path):
    """같은 결과 파일에 두 번 — 두 번째도 원본 좌표로 맞는 행에 쓴다(행을 지운 뒤에도)."""
    base = [H, [1, "a", ""], [2, "b", ""], [3, "c", ""]]
    left = [H, [1, "a", ""], [3, "C", ""]]           # 2 지움, 3 고침
    right = base
    l, b, r = _three(tmp_path, left, base, right)
    m, meta, cm = _diff(l, b, r)
    plan = auto_merge_plan(m)
    row2 = next(i for i, mm in enumerate(meta) if mm[1] == 2)
    first = {k: v for k, v in plan.items() if k[0] == row2}
    out = _save(tmp_path, l, b, r, first, m, meta, cm)
    assert _vals(out) == [H, ["1", "a", ""], ["3", "c", ""]]
    out = _save(tmp_path, l, b, r, plan, m, meta, cm, out)
    assert _vals(out) == [H, ["1", "a", ""], ["3", "C", ""]], "두 번째 저장이 옛 좌표로 썼다"


def test_result_can_be_the_middle_file_itself(tmp_path):
    base = [H, [1, "a", "p"]]
    left = [H, [1, "A", "p"]]
    l, b, r = _three(tmp_path, left, base, base)
    template = str(tmp_path / "template.xlsx")
    import shutil
    shutil.copyfile(b, template)
    m, meta, cm = _diff(l, template, r)
    write_merge3(template, b, [("Data", m, meta, cm, auto_merge_plan(m))],
                 {SRC_LEFT: l, SRC_RIGHT: r})
    assert _vals(b) == [H, ["1", "A", "p"]]


def test_styles_come_from_the_side_that_gave_the_value(tmp_path):
    base = [H, [1, "a", "p"]]
    left = [H, [1, "a", "p"], [2, "L2", ""]]
    right = [H, [1, "a", "P"]]

    def bold_new_row(ws):
        ws["B3"].font = Font(bold=True)

    def fill_w(ws):
        ws["C2"].fill = PatternFill("solid", fgColor="FFFF0000")
    l, b, r = _three(tmp_path, left, base, right, l=bold_new_row, r=fill_w)
    m, meta, cm = _diff(l, b, r)
    out = _save(tmp_path, l, b, r, auto_merge_plan(m), m, meta, cm)
    ws = openpyxl.load_workbook(out)["Data"]
    assert ws["B3"].value == "L2" and ws["B3"].font.bold, "왼쪽에서 온 새 행의 서식이 빠졌다"
    assert ws["C2"].value == "P" and ws["C2"].fill.fgColor.rgb == "FFFF0000", \
        "오른쪽에서 온 칸의 서식이 빠졌다"


def test_a_failed_save_leaves_the_result_untouched(tmp_path):
    base = [H, [1, "a", "p"]]
    left = [H, [1, "A", "p"]]
    l, b, r = _three(tmp_path, left, base, base)
    out = str(tmp_path / "result.xlsx")
    open(out, "wb").write(b"previous result")
    m, meta, cm = _diff(l, b, r)
    with pytest.raises(Exception):
        write_merge3(b, out, [("NoSuchSheet", m, meta, cm, auto_merge_plan(m))],
                     {SRC_LEFT: l, SRC_RIGHT: r})
    assert open(out, "rb").read() == b"previous result"
    assert [n for n in os.listdir(tmp_path) if ".em3_" in n or ".tmp_merge" in n] == []


def test_patch_builder_merges_inserts_from_both_sides_into_one_row():
    """양쪽이 같은 새 행을 다르게 만들었고 칸마다 다른 쪽을 고르면 한 행으로 붙는다."""
    m = [[("same", "ID", "ID", "ID"), ("same", "V", "V", "V"), ("same", "W", "W", "W")],
         [("both", "9", "", "9"), ("conflict", "L", "", "R"), ("conflict", "l", "", "r")]]
    meta = [(0, 0, 0), (1, None, 1)]
    cm = [(0, 0, 0), (1, 1, 1), (2, 2, 2)]
    p, ins, _ss = build_merge3_patches(m, meta, cm, {(1, 0): SRC_LEFT, (1, 1): SRC_RIGHT,
                                                     (1, 2): SRC_LEFT})
    assert p == {}
    assert [(c, v) for c, v, _s in ins[1]] == [(0, "9"), (1, "R"), (2, "l")]
