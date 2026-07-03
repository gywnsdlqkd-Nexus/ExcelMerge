"""파일 로더 — xlsx/json 로딩과 확장자 디스패처 (excel_diff_merge.py에서 분리)."""
import os
import re
import json

import openpyxl
from openpyxl.utils import column_index_from_string

from .uasset_parser import load_uasset_as_matrix


_EXCEL_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb"}
_SUPPORTED_EXTS = _EXCEL_EXTS | {".json", ".uasset"}


def _cell_to_str(v) -> str:
    """openpyxl 캐시값을 문자열로 변환 (정수형 float은 정수로)."""
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


def load_sheet_with_formulas(path: str) -> tuple[list[list], list[list]]:
    """
    한 번의 호출로 (계산값 시트, 수식 시트)를 함께 반환한다.

    핵심 개선:
      - read_only=True 로 두 워크북 모두 로드 → 메모리/속도 모두 향상
        (기본 모드 대비 수~수십 배 빠름).
      - Excel COM 의존 완전 제거 — win32com 호출 없음.
      - 미캐시 수식 셀만 _eval_formula_with_row 로 같은 행 컨텍스트 계산
        (전체 시트를 셀 단위로 COM 호출하던 기존 폴백 제거).

    수식 시트는 NetmarbleCompare.md의 '수식 보존 병합' 기능을 위해
    저장 시 원본 수식(=...)을 그대로 기록하는 데 사용되므로 항상 함께 반환.
    """
    # 1차 — 캐시값
    wb_val = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws_val = wb_val.worksheets[0]
    values: list[list[str]] = [
        [_cell_to_str(v) for v in row]
        for row in ws_val.iter_rows(values_only=True)
    ]
    try:
        wb_val.close()
    except Exception:
        pass

    # 2차 — 수식 텍스트 (data_only=False)
    wb_fml = openpyxl.load_workbook(path, read_only=True, data_only=False)
    ws_fml = wb_fml.worksheets[0]
    formulas: list[list[str]] = []
    needs_eval: list[tuple[int, int, str]] = []   # (r, c, formula_text)
    for r_idx, row in enumerate(ws_fml.iter_rows(values_only=True)):
        out_row = []
        for c_idx, v in enumerate(row):
            if v is None:
                out_row.append("")
                continue
            s = str(v)
            out_row.append(s)
            # 같은 좌표의 캐시값이 비었고 수식 텍스트면 자체 계산 대상
            if (
                s.startswith("=")
                and r_idx < len(values)
                and c_idx < len(values[r_idx])
                and values[r_idx][c_idx] == ""
            ):
                needs_eval.append((r_idx, c_idx, s))
        formulas.append(out_row)
    try:
        wb_fml.close()
    except Exception:
        pass

    # 미캐시 수식만 자체 계산 — 캐시가 모두 채워진 일반 파일은 0회
    for r_idx, c_idx, fml in needs_eval:
        if r_idx < len(values) and c_idx < len(values[r_idx]):
            values[r_idx][c_idx] = _eval_formula_with_row(fml, values[r_idx])

    return values, formulas


def load_sheet(path: str) -> list[list]:
    """하위 호환용 — 계산값 시트만 반환."""
    values, _ = load_sheet_with_formulas(path)
    return values


def _eval_formula_with_row(formula: str, row_data: list) -> str:
    """수식을 같은 행 데이터를 컨텍스트로 계산. 실패 시 수식 문자열 그대로 반환."""
    try:
        import formulas as _formulas
        import numpy as _numpy
        # A22, B5 등 행 번호를 전부 1로 정규화 (같은 행 데이터로 매핑하기 위해)
        normalized = re.sub(r'([A-Za-z]+)\d+', lambda m: m.group(1).upper() + '1', formula)
        col_refs = re.findall(r'([A-Z]+)1', normalized)
        kwargs = {}
        for ref in set(col_refs):
            c_idx = column_index_from_string(ref) - 1
            if 0 <= c_idx < len(row_data):
                val = row_data[c_idx]
                try:
                    fv = float(val)
                    val = int(fv) if fv == int(fv) else fv
                except (ValueError, TypeError):
                    pass
                kwargs[f'{ref}1'] = _numpy.array([[val]])
        func = _formulas.Parser().ast(normalized)[1].compile()
        result = func(**kwargs)
        v = list(result.flat)[0] if hasattr(result, 'flat') else result
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v)
    except Exception:
        return formula


def load_sheet_formulas(path: str) -> list[list]:
    """하위 호환용 — load_sheet_with_formulas()의 두 번째 결과만 반환."""
    _, formulas = load_sheet_with_formulas(path)
    return formulas


def _json_value_to_str(v) -> str:
    """JSON 스칼라/구조 값을 셀 표시용 문자열로 변환.
    - dict/list 등 비스칼라는 compact JSON 으로 직렬화 — 셀 한 칸에 들어가도록.
    """
    if v is None:
        return ""
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (int, float)):
        return _cell_to_str(v)
    if isinstance(v, str):
        return v
    return json.dumps(v, ensure_ascii=False)


def _flatten_json(node, prefix: str = "") -> list[tuple[str, str]]:
    """객체/배열을 점-경로 평탄화 — `[ (경로, 값문자열), ... ]`. 폴백 표기용."""
    rows: list[tuple[str, str]] = []
    if isinstance(node, dict):
        if not node:
            rows.append((prefix or "(empty object)", "{}"))
            return rows
        for k, v in node.items():
            key = f"{prefix}.{k}" if prefix else str(k)
            if isinstance(v, (dict, list)):
                rows.extend(_flatten_json(v, key))
            else:
                rows.append((key, _json_value_to_str(v)))
    elif isinstance(node, list):
        if not node:
            rows.append((prefix or "(empty array)", "[]"))
            return rows
        for i, v in enumerate(node):
            key = f"{prefix}[{i}]" if prefix else f"[{i}]"
            if isinstance(v, (dict, list)):
                rows.extend(_flatten_json(v, key))
            else:
                rows.append((key, _json_value_to_str(v)))
    else:
        rows.append((prefix or "(value)", _json_value_to_str(node)))
    return rows


def load_json_as_matrix(path: str) -> list[list[str]]:
    """JSON 파일을 비교용 2D 매트릭스로 변환.
    - 최상위가 객체 배열 [ {...}, {...}, ... ] 이면: 첫 행=키 union(객체 등장 순서 보존),
      이후 각 객체를 한 행으로 — 표 형태로 행 단위 비교 가능.
    - 그 외(객체/스칼라/스칼라 배열 등)는 점-경로 평탄화로 [경로, 값] 2열 폴백.
    """
    with open(path, "r", encoding="utf-8-sig") as f:
        node = json.load(f)

    if isinstance(node, list) and node and all(isinstance(x, dict) for x in node):
        # 키 union — 첫 등장 순서 보존
        keys: list[str] = []
        seen: set[str] = set()
        for obj in node:
            for k in obj.keys():
                if k not in seen:
                    seen.add(k)
                    keys.append(k)
        matrix: list[list[str]] = [list(keys)]
        for obj in node:
            row = [_json_value_to_str(obj.get(k, "")) for k in keys]
            matrix.append(row)
        return matrix

    # 폴백: 평탄화 2열
    matrix = [["path", "value"]]
    for path_str, val in _flatten_json(node):
        matrix.append([path_str, val])
    return matrix


def load_sheet_with_formulas_any(path: str) -> tuple[list[list], list[list]]:
    """확장자에 따라 (values, formulas) 매트릭스를 반환하는 통합 디스패처.
    - xlsx 계열: 기존 load_sheet_with_formulas() 사용 (수식 별도 추출).
    - json/uasset: 수식 개념이 없으므로 동일 매트릭스를 두 번 반환.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in _EXCEL_EXTS:
        return load_sheet_with_formulas(path)
    if ext == ".json":
        m = load_json_as_matrix(path)
        return m, [list(row) for row in m]
    if ext == ".uasset":
        m = load_uasset_as_matrix(path)
        return m, [list(row) for row in m]
    # 알려지지 않은 확장자 — 기존 로직(엑셀)로 폴백
    return load_sheet_with_formulas(path)
