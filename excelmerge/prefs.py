"""사용자 로컬 설정(전역) 저장 — 현재는 '키 헤더 위치'(키 행 + 키 열)만 다룬다.

저장 위치는 updater 의 %APPDATA%/ExcelMerge/ 규약을 그대로 따른다(APPDATA 없으면 홈).
디폴트 앵커는 A1(key_row=0, key_col=0)이며, 사용자가 키 행/열을 한 번이라도 바꾸면
그 위치를 여기에 기록해 이후 모든 비교의 기본 앵커로 쓴다.

읽기/쓰기 모두 예외를 삼켜 UI를 절대 죽이지 않는다(설정은 부가 기능이므로 실패는 조용히 무시).
"""
import json
import os
import shutil
import time
from contextlib import contextmanager

try:
    import msvcrt
except ImportError:          # 윈도우 밖(개발 환경) — 프로세스 간 잠금 없이
    msvcrt = None

from .logutil import log
from .constants import appdata_path

_DEFAULT_KEY_ROW = 0
_DEFAULT_KEY_COL = 0
_LAST_SHEETS_MAX = 50   # 파일별 마지막 시트 기억 상한(오래된 항목부터 제거)


def _prefs_path() -> str:
    return appdata_path("prefs.json")


def _load_prefs_file(p: str):
    """prefs.json 을 읽는다 → dict, 파일이 없으면 {}, 끝내 못 읽으면 None.

    잠깐씩 다시 시도한다. 다른 프로세스가 os.replace 로 파일을 바꾸는 그 순간엔 윈도우가
    열기 자체를 거부한다(실측: 쓰는 도중 4,860번 읽기 중 41번). 그걸 '설정 없음'으로 보면
    그 비교는 기본 키(A1)로 돌고, 고쳐 쓰는 쪽은 빈 설정으로 덮어쓴다.

    내용이 깨진 경우(JSON 오류)는 **한 번만** 다시 본다. 이 앱은 원자적으로 쓰므로 깨진 내용은
    대개 정말 깨진 것이다(옛 버전 앱이 쓰는 중인 드문 경우만 한 번 더 보면 된다). 매번 여러 번
    기다리면 깨진 파일이 고쳐질 때까지 **실행할 때마다** 1초 넘게 멈췄다(2차 검토).
    """
    parse_retries = 1
    for _ in range(6):
        if not os.path.isfile(p):
            return {}
        try:
            with open(p, "r", encoding="utf-8-sig") as f:
                text = f.read()
        except OSError:
            time.sleep(0.02)                     # 교체 순간의 열기 거부 — 곧 풀린다
            continue
        try:
            c = json.loads(text)
        except ValueError:
            if parse_retries <= 0:
                return None
            parse_retries -= 1
            time.sleep(0.02)
            continue
        return c if isinstance(c, dict) else None
    return None


def _read_prefs() -> dict:
    """prefs.json 전체를 dict로 반환. 없음/깨짐이면 빈 dict."""
    try:
        c = _load_prefs_file(_prefs_path())
        if c is not None:
            return c
        log.debug("prefs 읽기 실패(기본값 사용)")
    except Exception:
        log.debug("prefs 읽기 실패(기본값 사용)", exc_info=True)
    return {}


def _write_prefs(data: dict) -> None:
    """prefs.json 전체를 **원자적으로** 바꾼다 — 임시 파일에 다 쓴 뒤 os.replace.

    예전엔 open(p, "w") 로 파일을 먼저 비우고 썼다. 그 사이 다른 ExcelMerge 프로세스가
    읽으면 빈/반쪽 파일을 받아 {} 로 보고, 자기 변경만 담아 다시 썼다 — 파일별 키 위치·
    제외 열·시트 기억이 통째로 사라졌다(실측: 동시에 뜬 6개 중 2/10회). P4V 에서 여러 파일을
    한꺼번에 Diff 하면 프로세스가 거의 동시에 뜬다. 실패는 조용히 무시(설정은 부가 기능).
    """
    tmp = None
    try:
        p = _prefs_path()
        if os.path.exists(p) and not os.access(p, os.W_OK):
            log.debug("prefs.json 이 읽기 전용 — 저장하지 않음")   # 기다려도 안 풀린다
            return
        os.makedirs(os.path.dirname(p), exist_ok=True)
        tmp = f"{p}.{os.getpid()}.tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f)
            f.flush()
            os.fsync(f.fileno())
        # 다른 프로세스가 읽느라 열어 둔 순간엔 윈도우가 교체를 거절한다 — 잠깐 다시 시도
        for _ in range(10):
            try:
                os.replace(tmp, p)
                tmp = None
                return
            except PermissionError:
                time.sleep(0.02)
        log.debug("prefs 교체 실패(다른 프로세스가 붙잡고 있음) — 이번 저장은 건너뜀")
    except Exception:
        log.debug("prefs 쓰기 실패", exc_info=True)
    finally:
        if tmp is not None:
            try:
                os.remove(tmp)
            except OSError:
                pass


@contextmanager
def _prefs_lock(timeout: float = 1.0):
    """여러 프로세스의 '읽고-고쳐-쓰기'를 한 줄로 세운다(prefs.json.lock 1바이트 잠금).

    원자적 교체만으로는 **변경 유실**이 남는다 — 둘이 같은 옛 내용을 읽고 각자 고쳐 쓰면
    나중 것이 앞의 것을 지운다. 잠금을 못 얻으면(1초) 그냥 진행한다 — UI 를 멈추지 않는다.
    """
    fh = None
    locked = False
    try:
        p = _prefs_path() + ".lock"
        os.makedirs(os.path.dirname(p), exist_ok=True)
        fh = open(p, "a+b")
        if msvcrt is not None:
            end = time.monotonic() + timeout
            while True:
                try:
                    fh.seek(0)
                    msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
                    locked = True
                    break
                except OSError:
                    if time.monotonic() >= end:
                        log.debug("prefs 잠금 시간 초과 — 잠금 없이 진행")
                        break
                    time.sleep(0.01)
    except Exception:
        log.debug("prefs 잠금 실패 — 잠금 없이 진행", exc_info=True)
    try:
        yield
    finally:
        if fh is not None:
            try:
                if locked:
                    fh.seek(0)
                    msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
            except OSError:
                pass
            fh.close()


def _read_prefs_for_update() -> dict:
    """고쳐 쓰기 전에 읽는다 — 못 읽는 파일을 빈 설정으로 덮어써 날리지 않는다.

    잠깐 다시 읽어 본다(_load_prefs_file — 교체 중이거나 옛 버전 앱이 쓰는 중일 수 있다).
    그래도 못 읽으면 정말 깨진 것이다 — 원본을 prefs.json.corrupt-<시각> 으로 남기고 새로 시작한다. 쓰기를 아예
    포기하면 그 뒤로 설정을 영영 저장하지 못한다.
    """
    p = _prefs_path()
    c = _load_prefs_file(p)
    if c is not None:
        return c
    try:
        backup = f"{p}.corrupt-{time.strftime('%Y%m%d_%H%M%S')}"
        shutil.copy2(p, backup)
        log.warning("prefs.json 을 읽을 수 없어 새로 씁니다 — 원본은 %s", backup)
    except OSError:
        log.debug("깨진 prefs 백업 실패", exc_info=True)
    return {}


def _update_prefs(mutate) -> None:
    """잠금 안에서 읽고 → mutate(설정 dict) → 쓴다. mutate 가 False 를 돌려주면 쓰지 않는다.
    어떤 실패도 밖으로 내보내지 않는다."""
    try:
        with _prefs_lock():
            c = _read_prefs_for_update()
            if mutate(c) is False:
                return
            _write_prefs(c)
    except Exception:
        log.debug("prefs 갱신 실패", exc_info=True)


def _lru_put(c: dict, field: str, key: str, value) -> None:
    """파일별 기억(간단 LRU): key 를 맨 뒤로 다시 넣고 상한을 넘는 오래된 항목을 뺀다."""
    m = c.get(field)
    if not isinstance(m, dict):
        m = {}
    m.pop(key, None)                      # 재삽입으로 최근 항목을 뒤로
    m[key] = value
    if len(m) > _LAST_SHEETS_MAX:
        for k in list(m.keys())[: len(m) - _LAST_SHEETS_MAX]:
            del m[k]
    c[field] = m


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
    def mutate(c):
        c["key_row"] = int(key_row)
        c["key_col"] = int(key_col)

    _update_prefs(mutate)


def load_last_key(path: str):
    """그 파일에서 사용자가 마지막으로 고른 (key_row, key_col). 없으면 None.

    키는 오랫동안 **전역 하나**였다. 그래서 한 파일에서 B열을 키로 잡으면 그 뒤로 여는
    모든 파일이 B열로 비교됐고, B가 키가 아닌 파일에서는 중복 키로 수천 행이 조용히
    비교에서 빠졌다(실측: Data_SkillModule_CS.xlsx 14,659행). 파일별로 기억해 두면
    한 번 제대로 잡은 파일은 다음부터 틀리지 않는다. 기억이 없으면 전역 기본을 쓴다.
    """
    if not path:
        return None
    m = _read_prefs().get("last_keys")
    if isinstance(m, dict):
        v = m.get(os.path.abspath(path))
        # [row, col] (옛) 또는 [row, col, 헤더이름] (지금) — 둘 다 읽는다.
        if (isinstance(v, list) and len(v) in (2, 3)
                and all(isinstance(x, int) for x in v[:2])
                and v[0] >= 0 and v[1] >= -1):
            return (v[0], v[1])
    return None


def load_last_key_name(path: str):
    """그 파일에서 마지막으로 고른 **키 열의 헤더 이름**. 없으면 None.

    번호만 기억하면 열 하나가 끼는 순간 옆 열을 키로 잡는다. 이름이 있으면 자리가
    바뀌어도 따라간다 — 번호는 이름을 못 찾았을 때의 대비책으로만 쓴다.
    """
    if not path:
        return None
    m = _read_prefs().get("last_keys")
    if isinstance(m, dict):
        v = m.get(os.path.abspath(path))
        if isinstance(v, list) and len(v) >= 3 and isinstance(v[2], str) and v[2]:
            return v[2]
    return None


def save_last_key(path: str, key_row: int, key_col: int, key_name: str = "") -> None:
    """파일별 키 위치를 기록(간단 LRU). 실패는 조용히 무시.

    형식: [key_row, key_col] (옛) 또는 [key_row, key_col, 헤더이름] (지금).
    옛 형식도 그대로 읽히므로 기존 설정이 깨지지 않는다.
    """
    if not path:
        return
    entry = [int(key_row), int(key_col), str(key_name or "")]
    _update_prefs(lambda c: _lru_put(c, "last_keys", os.path.abspath(path), entry))


def _sheet_key_id(path: str, sheet) -> str:
    return os.path.abspath(path) + "|" + (sheet or "")


def load_sheet_key(path: str, sheet):
    """3파일 비교가 **그 파일의 그 시트**에서 고른 (key_row, key_col, 헤더이름). 없으면 None.

    파일 단위 기억(load_last_key)은 시트를 가리지 않아, 한 시트에서 고른 키 이름(ItemID)이 다른 시트에
    외래 키 열로 있으면 그 열이 키가 됐다 — 3파일은 P4V 첫 방문에 자동 병합을 걸므로 같은 ID 두 행이
    저장됐다(검토 실측). 그래서 3파일은 시트마다 따로 기억한다(2파일 기억과 섞지 않는다).
    """
    if not path:
        return None
    m = _read_prefs().get("sheet_keys")
    if isinstance(m, dict):
        v = m.get(_sheet_key_id(path, sheet))
        if (isinstance(v, list) and len(v) == 3 and all(isinstance(x, int) for x in v[:2])
                and v[0] >= 0 and v[1] >= -1 and isinstance(v[2], str)):
            return (v[0], v[1], v[2])
    return None


def save_sheet_key(path: str, sheet, key_row: int, key_col: int, key_name: str = "") -> None:
    """3파일 시트별 키 기억(간단 LRU). 실패는 조용히 무시."""
    if not path:
        return
    entry = [int(key_row), int(key_col), str(key_name or "")]
    _update_prefs(lambda c: _lru_put(c, "sheet_keys", _sheet_key_id(path, sheet), entry))


def load_last_excluded(path: str):
    """그 파일에서 사용자가 마지막으로 지정한 '검사 제외 열' 목록. 없으면 None.

    제외는 비교할 때마다 초기화됐다. 그런데 어떤 테이블에서 무엇을 안 볼지는 **그 테이블의
    성질**이지 그때그때 기분이 아니다 — 주석 열(#Desc 등)이나 현지화 열은 늘 같은 것을
    뺀다. 그래서 파일을 열 때마다 헤더를 우클릭해 같은 열을 다시 골라야 했다.
    키 위치를 파일별로 기억하게 만든 것과 같은 이유다(load_last_key 참조).

    **빈 목록과 '기억 없음'은 다르다.** 사용자가 제외를 전부 해제한 것도 선택이므로
    `[]` 를 그대로 돌려준다 — 호출부는 `is None` 으로 구분해야 한다.

    돌려주는 모양(save_last_excluded 참조):
      · 지금 형식 — 칸마다 "이름"(그 이름의 첫째) 또는 (이름, 몇 번째) 또는 ("", 화면 열 번호).
        이름만 있는 옛 설정도 이 형식의 일부로 그대로 읽힌다(= 그 이름의 첫째).
      · 아주 옛 형식 — 숫자만 있는 목록(화면 열 번호). 정렬해 돌려준다.
    알아볼 수 없는 칸이 하나라도 있으면 None(기억 없음)이다.
    """
    if not path:
        return None
    m = _read_prefs().get("last_excluded")
    if isinstance(m, dict):
        v = m.get(os.path.abspath(path))
        if isinstance(v, list):
            if v and all(_is_col_no(x) for x in v):
                return sorted(set(v))                # 아주 옛 형식 — 화면 열 번호
            out = []
            for x in v:
                if x == "":
                    continue                         # 빈 이름 — 찾을 수 없다
                ident = _excluded_ident(x)
                if ident is None:
                    return None
                out.append(x if isinstance(x, str) else ident)
            return out
    return None


def _is_col_no(x) -> bool:
    return isinstance(x, int) and not isinstance(x, bool) and x >= 0


def _excluded_ident(x):
    """제외 기억 한 칸 → (이름, n). 알아볼 수 없으면 None.

      "이름"     → (이름, 0)   그 이름의 첫째 열. 이름만 남기던 옛 설정도 이렇게 읽는다.
      [이름, k]  → (이름, k)   같은 이름의 k번째 열(0부터) — '#' 주석 열이 여럿일 때.
      ["", c]    → ("", c)     이름 없는 열 — 화면 c열(자리로 기억한다).
    """
    if isinstance(x, str):
        return (x, 0)
    if (isinstance(x, (list, tuple)) and len(x) == 2
            and isinstance(x[0], str) and _is_col_no(x[1])):
        return (x[0], x[1])
    return None


def save_last_excluded(path: str, cols) -> None:
    """파일별 제외 열을 기록(간단 LRU). 실패는 조용히 무시.

    cols 의 칸마다 (이름, 그 이름의 몇 번째) — 열 매칭(diff_engine.column_tags)과 같은
    식별자다. 이름만 기억하면 '#' 처럼 같은 이름이 여럿일 때 엉뚱한 '#' 를 가리킨다.
    이름 없는 열(빈 헤더)은 ("", 화면 열 번호) 로 자리를 기억한다. "이름" 은 (이름, 0).

    저장 모양: 그 이름의 첫째는 예전처럼 **이름 문자열만** 남긴다(이 버전 이전도 읽는다),
    둘째부터와 이름 없는 열은 [이름, k] · ["", c]. 숫자만 주면 아주 옛 형식(화면 열 번호)으로
    저장한다(읽기 호환 확인용 경로). 알아볼 수 없는 칸이 섞이면 아무것도 저장하지 않는다.
    """
    if not path:
        return
    cols = list(cols)
    if cols and all(isinstance(c, int) and not isinstance(c, bool) for c in cols):
        vals = sorted({c for c in cols if c >= 0})
    else:
        ids = set()
        for c in cols:
            if c == "":
                continue                             # 빈 이름 — 찾을 수 없다
            ident = _excluded_ident(c)
            if ident is None:
                return
            ids.add(ident)
        vals = [n if (n and k == 0) else [n, k] for n, k in sorted(ids)]
    _update_prefs(lambda c: _lru_put(c, "last_excluded", os.path.abspath(path), vals))


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
    _update_prefs(lambda c: _lru_put(c, "last_sheets", os.path.abspath(path), name))


# ── A/B 경로칸 히스토리 (드롭다운) ─────────────────────────────────────────────
# 쪽(a/b)마다 최근에 연 경로를 최신순으로 둔다. 항목은 [경로, 종류('f' 파일 / 'd' 폴더)].
# 파일 비교와 폴더 비교가 같은 목록을 쓰고, 보여 줄 때 각 화면이 필요한 종류만 고른다.
# 상한은 **종류마다** 20개 — 한 상한을 나눠 쓰면 파일을 많이 연 날 폴더 기록이 밀려나
# 폴더 비교의 ▼ 가 비어 버렸다(검토에서 확인).
_PATH_HISTORY_MAX = 20
_KINDS = ("f", "d")


def _transient_roots() -> list:
    """곧 지워지는 경로의 뿌리 — 시스템 임시 폴더.

    P4V 는 비교할 리비전을 %TEMP%/p4v/... 에 내려받아 넘긴다. 그런 파일은 금방 사라지므로
    히스토리에 남기면 목록만 어지럽힌다. (테스트가 임시 폴더에 파일을 만들므로 바꿔 끼울 수
    있게 함수로 둔다.)
    """
    import tempfile
    return [tempfile.gettempdir()]


def _norm_variants(path: str) -> set:
    """대소문자·짧은 이름(GYWNSD~1)·상대 경로 차이를 없앤 비교용 형태들."""
    out = {os.path.normcase(os.path.abspath(path))}
    try:
        out.add(os.path.normcase(os.path.realpath(path)))
    except (OSError, ValueError):
        pass
    return out


def _is_transient_path(path: str) -> bool:
    for root in _transient_roots():
        for r in _norm_variants(root):
            r = r.rstrip("\\/")
            for p in _norm_variants(path):
                if p == r or p.startswith(r + os.sep):
                    return True
    return False


def _history_key(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def _history_entries(c: dict, side: str) -> list:
    """저장된 그 쪽 항목을 [(경로, 종류)] 로. 깨진 항목은 버린다. 손으로 고친 문자열
    항목도 받아 준다(종류는 지금 폴더인지로 정한다)."""
    h = c.get("path_history")
    v = h.get(side) if isinstance(h, dict) else None
    out = []
    for e in v if isinstance(v, list) else []:
        if isinstance(e, str) and e:
            out.append((e, "d" if os.path.isdir(e) else "f"))
        elif (isinstance(e, (list, tuple)) and len(e) == 2 and isinstance(e[0], str)
              and e[0] and e[1] in _KINDS):
            out.append((e[0], e[1]))
    return out


def _trim(entries: list) -> list:
    """종류마다 최신 _PATH_HISTORY_MAX 개만 — 순서는 그대로."""
    seen = {k: 0 for k in _KINDS}
    out = []
    for p, k in entries:
        if seen[k] < _PATH_HISTORY_MAX:
            seen[k] += 1
            out.append((p, k))
    return out


def load_path_history(side: str, kinds=_KINDS) -> list:
    """쪽(a/b)의 경로 히스토리 — 최신순, kinds('f'/'d') 만. 없거나 깨졌으면 빈 목록."""
    return [p for p, k in _trim(_history_entries(_read_prefs(), side)) if k in kinds]


def add_path_history(side: str, path: str) -> None:
    """경로를 그 쪽 히스토리 맨 앞에 둔다(같은 경로는 하나만). 임시 폴더 경로는 남기지 않는다.
    이미 맨 앞이면 쓰지 않는다(켤 때마다 쓰는 일을 줄인다). 실패는 조용히 무시."""
    if side not in ("a", "b", "m") or not isinstance(path, str) or not path.strip():
        return
    path = os.path.abspath(path.strip().strip('"'))
    if _is_transient_path(path):
        return
    kind = "d" if os.path.isdir(path) else "f"
    key = _history_key(path)

    def mutate(c):
        cur = _history_entries(c, side)
        if cur and cur[0] == (path, kind):
            return False
        new = _trim([(path, kind)] + [(p, k) for p, k in cur if _history_key(p) != key])
        h = c.get("path_history")
        if not isinstance(h, dict):
            h = {}
        h[side] = [[p, k] for p, k in new]
        c["path_history"] = h

    _update_prefs(mutate)


def clear_path_history(side: str, kinds=_KINDS) -> None:
    """그 쪽 히스토리에서 kinds 종류만 지운다 — 폴더 화면의 '목록 지우기'가 보이지도 않던
    파일 기록까지 지우지 않게."""
    def mutate(c):
        cur = _history_entries(c, side)
        keep = [[p, k] for p, k in cur if k not in kinds]
        if len(keep) == len(cur):
            return False
        h = c.get("path_history")
        if not isinstance(h, dict):
            h = {}
        h[side] = keep
        c["path_history"] = h

    _update_prefs(mutate)
