# -*- coding: utf-8 -*-
"""크로스포맷 bool 표기 일치 회귀 테스트.

xlsx 는 파이썬 str(bool) 을 써서 "True"/"False" 를, json/uasset 로더는 "true"/"false" 를
내보내고 있었다. 그 결과 **같은 논리값인데도 xlsx↔json 비교에서 bool 열 전체가 '변경'으로
잡히는** 거짓 차이가 났다. 세 포맷 모두 constants.bool_to_str 정본을 쓰는지 확인한다.
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl

from excelmerge import constants, loaders
from excelmerge.diff_engine import compute_diff


def test_canonical_bool_tokens():
    assert constants.bool_to_str(True) == "true"
    assert constants.bool_to_str(False) == "false"


def test_xlsx_bool_cells_use_canonical_form(tmp_path):
    """엑셀 TRUE/FALSE 셀이 소문자 정본으로 로드돼야 한다."""
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Flag"])
    ws.append([1, True])
    ws.append([2, False])
    wb.save(p)

    vals = loaders.load_values_any(str(p))
    assert vals[1] == ["1", "true"], vals[1]
    assert vals[2] == ["2", "false"], vals[2]


def test_json_bool_matches_xlsx_bool(tmp_path):
    """동일 데이터의 xlsx 와 json 이 bool 열에서 차이로 잡히지 않아야 한다."""
    xp = tmp_path / "a.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Flag"])
    ws.append([1, True])
    ws.append([2, False])
    wb.save(xp)

    jp = tmp_path / "b.json"
    jp.write_text(json.dumps([{"ID": 1, "Flag": True}, {"ID": 2, "Flag": False}]),
                  encoding="utf-8")

    a = loaders.load_values_any(str(xp))
    b = loaders.load_values_any(str(jp))
    dm, _meta = compute_diff(a, b, key_col=0)

    # 헤더 아래 모든 셀이 same 이어야 한다(bool 표기 차이로 인한 거짓 diff 없음)
    changed = [(r, c, cell) for r, row in enumerate(dm) for c, cell in enumerate(row)
               if cell[0] != "same"]
    assert not changed, f"거짓 차이 발생: {changed}"


def test_uasset_bool_uses_same_helper():
    """uasset 파서도 같은 정본 함수를 쓰는지(리터럴 재도입 방지)."""
    import inspect
    from excelmerge import uasset_parser
    src = inspect.getsource(uasset_parser)
    assert 'bool_to_str' in src
    assert '"true" if' not in src, "uasset 에 bool 리터럴이 되살아났다"
