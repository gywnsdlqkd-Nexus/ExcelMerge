"""파일 로더 — xlsx/json 로딩과 확장자 디스패처 (excel_diff_merge.py에서 분리)."""
import io
import os
import re
import json
import zipfile

import openpyxl
from openpyxl.utils import column_index_from_string

from .uasset_parser import load_uasset_as_matrix


_EXCEL_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb"}
_SUPPORTED_EXTS = _EXCEL_EXTS | {".json", ".uasset"}

# 빈 <fill/> — openpyxl이 로드 시 거부(TypeError: expected ...Fill)하는 요소
_EMPTY_FILL_RE = re.compile(r"<fill\s*/>|<fill>\s*</fill>")


def _sanitize_xlsx_bytes(path: str) -> io.BytesIO:
    """styles.xml의 빈 <fill/>을 유효한 최소 형태로 치환한 zip 사본을 메모리로 반환.
    비-Excel 툴이 만든 파일에서 openpyxl이 fill 파싱에 실패하는 경우의 폴백.
    numFmt 등 나머지는 그대로라 날짜/숫자 파싱은 정상 유지된다."""
    buf = io.BytesIO()
    with zipfile.ZipFile(path) as zin, \
         zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                text = data.decode("utf-8", errors="replace")
                text = _EMPTY_FILL_RE.sub("<fill><patternFill/></fill>", text)
                data = text.encode("utf-8")
            zout.writestr(item, data)
    buf.seek(0)
    return buf


def _open_workbook(path: str, data_only: bool):
    """openpyxl 워크북 로드 — 빈 fill 등으로 스타일 파싱이 실패하면 styles.xml을
    정제한 사본으로 1회 재시도한다. 정상 파일은 예외 경로를 타지 않는다."""
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=data_only)
    except TypeError:
        return openpyxl.load_workbook(
            _sanitize_xlsx_bytes(path), read_only=True, data_only=data_only)


def _cell_to_str(v) -> str:
    """openpyxl 캐시값을 문자열로 변환 (정수형 float은 정수로)."""
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


def _load_values_pass(path: str, progress=None) -> list[list[str]]:
    """data_only=True 패스 — 캐시된 계산값 시트.
    progress(done_rows, total_rows_or_None)를 500행마다 호출한다."""
    wb = _open_workbook(path, data_only=True)
    try:
        ws = wb.worksheets[0]
        # read_only 모드의 max_row는 dimension 헤더에서 읽는다 — 없으면 None(불확정)
        total = ws.max_row
        values: list[list[str]] = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            values.append([_cell_to_str(v) for v in row])
            if progress is not None and (i + 1) % 500 == 0:
                progress(i + 1, total)
        if progress is not None:
            progress(len(values), total)
        return values
    finally:
        try:
            wb.close()
        except Exception:
            pass


def _load_formulas_pass(path: str) -> tuple[list[list[str]], list[tuple[int, int, str]]]:
    """data_only=False 패스 — 수식 텍스트 시트 + '='로 시작하는 셀 좌표 목록."""
    wb = _open_workbook(path, data_only=False)
    try:
        ws = wb.worksheets[0]
        formulas: list[list[str]] = []
        candidates: list[tuple[int, int, str]] = []   # (r, c, formula_text)
        for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
            out_row = []
            for c_idx, v in enumerate(row):
                if v is None:
                    out_row.append("")
                    continue
                s = str(v)
                out_row.append(s)
                if s.startswith("="):
                    candidates.append((r_idx, c_idx, s))
            formulas.append(out_row)
        return formulas, candidates
    finally:
        try:
            wb.close()
        except Exception:
            pass


def load_sheet_with_formulas(path: str, progress=None) -> tuple[list[list], list[list]]:
    """
    한 번의 호출로 (계산값 시트, 수식 시트)를 함께 반환한다.

    핵심 개선:
      - read_only=True 로 두 워크북 모두 로드 → 메모리/속도 모두 향상.
      - 미캐시 수식 셀만 _eval_formula_with_row 로 같은 행 컨텍스트 계산.
      - progress(done_rows, total_rows_or_None) 콜백으로 로딩 진행률 보고.

    두 패스(값/수식)를 스레드로 동시 실행하는 안은 실측 결과 이득이 없어
    (30k×30 파일 기준 0.95배 — openpyxl XML 파싱이 GIL에 묶여 있음) 순차 유지.
    openpyxl은 한 파스에서 캐시값과 수식을 동시에 줄 수 없어 2패스 자체는 필수.

    수식 시트는 NetmarbleCompare.md의 '수식 보존 병합' 기능을 위해
    저장 시 원본 수식(=...)을 그대로 기록하는 데 사용되므로 항상 함께 반환.
    """
    values = _load_values_pass(path, progress)
    formulas, candidates = _load_formulas_pass(path)

    # '빈 캐시값 + 수식 텍스트' 판정은 두 패스 조인 후 수행 — 순차 구현과 결과 동일.
    # 캐시가 모두 채워진 일반 파일은 0회.
    for r_idx, c_idx, fml in candidates:
        if (r_idx < len(values) and c_idx < len(values[r_idx])
                and values[r_idx][c_idx] == ""):
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


def load_sheet_with_formulas_any(path: str, progress=None) -> tuple[list[list], list[list]]:
    """확장자에 따라 (values, formulas) 매트릭스를 반환하는 통합 디스패처.
    - xlsx 계열: 기존 load_sheet_with_formulas() 사용 (수식 별도 추출).
    - json/uasset: 수식 개념이 없으므로 동일 매트릭스를 두 번 반환.
    progress 콜백은 행 수 파악이 가능한 xlsx 경로에서만 쓰인다.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in _EXCEL_EXTS:
        return load_sheet_with_formulas(path, progress)
    if ext == ".json":
        m = load_json_as_matrix(path)
        return m, [list(row) for row in m]
    if ext == ".uasset":
        m = load_uasset_as_matrix(path)
        return m, [list(row) for row in m]
    # 알려지지 않은 확장자 — 기존 로직(엑셀)로 폴백
    return load_sheet_with_formulas(path, progress)
