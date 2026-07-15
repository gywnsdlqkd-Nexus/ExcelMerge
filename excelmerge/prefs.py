"""사용자 로컬 설정(전역) 저장 — 현재는 '키 헤더 위치'(키 행 + 키 열)만 다룬다.

저장 위치는 updater 의 %APPDATA%/ExcelMerge/ 규약을 그대로 따른다(APPDATA 없으면 홈).
디폴트 앵커는 A1(key_row=0, key_col=0)이며, 사용자가 키 행/열을 한 번이라도 바꾸면
그 위치를 여기에 기록해 이후 모든 비교의 기본 앵커로 쓴다.

읽기/쓰기 모두 예외를 삼켜 UI를 절대 죽이지 않는다(설정은 부가 기능이므로 실패는 조용히 무시).
"""
import json
import os

from .logutil import log

_DEFAULT_KEY_ROW = 0
_DEFAULT_KEY_COL = 0
_LAST_SHEETS_MAX = 50   # 파일별 마지막 시트 기억 상한(오래된 항목부터 제거)


def _prefs_path() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "ExcelMerge", "prefs.json")


def _read_prefs() -> dict:
    """prefs.json 전체를 dict로 반환. 없음/깨짐이면 빈 dict."""
    try:
        p = _prefs_path()
        if os.path.isfile(p):
            with open(p, "r", encoding="utf-8-sig") as f:
                c = json.load(f)
            if isinstance(c, dict):
                return c
    except Exception:
        log.debug("prefs 읽기 실패(기본값 사용)", exc_info=True)
    return {}


def _write_prefs(data: dict) -> None:
    """prefs.json 전체를 덮어쓴다(호출부가 읽고-병합 후 넘긴다). 실패는 조용히 무시."""
    try:
        p = _prefs_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        pass


def load_key_prefs() -> tuple[int, int]:
    """저장된 (key_row, key_col) 반환. 파일 없음/깨짐/이상값이면 기본값 (0, 0).
    key_row >= 0, key_col >= -1(-1 = ROW 순서 모드) 만 유효로 인정한다."""
    c = _read_prefs()
    kr = c.get("key_row", _DEFAULT_KEY_ROW)
    kc = c.get("key_col", _DEFAULT_KEY_COL)
    if isinstance(kr, int) and isinstance(kc, int) and kr >= 0 and kc >= -1:
        return kr, kc
    return _DEFAULT_KEY_ROW, _DEFAULT_KEY_COL


def save_key_prefs(key_row: int, key_col: int) -> None:
    """(key_row, key_col)를 전역 설정으로 기록. 다른 설정(last_sheets 등)은 보존.
    실패는 조용히 무시."""
    c = _read_prefs()
    c["key_row"] = int(key_row)
    c["key_col"] = int(key_col)
    _write_prefs(c)


def load_last_sheet(path: str):
    """해당 파일에서 사용자가 마지막으로 선택한 시트 이름. 없으면 None."""
    if not path:
        return None
    m = _read_prefs().get("last_sheets")
    if isinstance(m, dict):
        v = m.get(os.path.abspath(path))
        if isinstance(v, str):
            return v
    return None


def save_last_sheet(path: str, name: str) -> None:
    """파일별 마지막 선택 시트를 기록(간단 LRU, 상한 초과 시 오래된 항목 제거).
    실패는 조용히 무시."""
    if not path or not isinstance(name, str):
        return
    c = _read_prefs()
    m = c.get("last_sheets")
    if not isinstance(m, dict):
        m = {}
    key = os.path.abspath(path)
    m.pop(key, None)   # 재삽입으로 최근 항목을 뒤로
    m[key] = name
    if len(m) > _LAST_SHEETS_MAX:
        for k in list(m.keys())[: len(m) - _LAST_SHEETS_MAX]:
            del m[k]
    c["last_sheets"] = m
    _write_prefs(c)
