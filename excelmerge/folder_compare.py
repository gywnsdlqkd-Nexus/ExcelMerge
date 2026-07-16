"""폴더 단위 비교 백엔드 — 순수 로직(UI 비의존).

두 폴더를 재귀 순회해 지원 확장자(SUPPORTED_EXTS) 파일을 상대경로로 매칭하고,
각 파일 쌍을 same/modified/only_a/only_b 로 분류한다.

동일/변경 판정은 **바이트 해시**를 먼저 쓴다(크기 우선 비교 후 같으면 청크 단위 SHA1).
바이트가 같으면 즉시 SAME(가장 빠른 경로). 바이트가 다르면 지원 형식(xlsx/json/uasset)에
한해 **내용 비교**로 재확인한다 — xlsx 재저장(메타데이터·수식 캐시·스타일만 바뀜)처럼
논리적으로 동일한데 바이트만 다른 파일을 SAME으로 바로잡기 위함(v165~).
내용 비교는 바이트가 다른 소수 파일에만 수행되므로 대부분 파일은 빠른 바이트 경로를 탄다.
"""
import os
from concurrent.futures import ThreadPoolExecutor

from .loaders import (
    SUPPORTED_EXTS, EXCEL_EXTS, load_values_any, list_sheet_names,
)

# 상태 상수 — 뷰/테마가 색 매핑에 사용.
SAME = "same"
MODIFIED = "modified"
ONLY_A = "only_a"
ONLY_B = "only_b"

_CHUNK = 1 << 20   # 1 MiB — 바이트 비교 청크
# 매칭 파일쌍 동일성 판정 병렬도. SHA·calamine 모두 GIL을 풀어 실제 병렬 이득.
_CMP_WORKERS = min(8, (os.cpu_count() or 2))

# 파일쌍 동일성 판정 캐시 — (경로,mtime,size)×2 키. 파일이 바뀌면 키가 달라져 자동 무효화되므로
# 새로고침·병합 직후 재스캔에서 안 바뀐 파일은 재해시/재로드를 건너뛴다. dict get/set는 GIL 하
# 원자적이라 스레드 풀에서 락 없이 안전(동일 키 중복 계산은 같은 값이라 무해).
_pair_equal_cache: dict = {}
_CACHE_CAP = 20000   # 무한 증식 방지 상한 — 초과 시 전체 비움


class FolderEntry:
    """폴더 비교 결과의 한 파일 항목.

    rel_path : 폴더 기준 상대경로(표시·매칭용, 항상 '/' 구분자).
    path_a   : A 폴더 내 절대경로(없으면 "").
    path_b   : B 폴더 내 절대경로(없으면 "").
    status   : SAME | MODIFIED | ONLY_A | ONLY_B.
    """
    __slots__ = ("rel_path", "path_a", "path_b", "status")

    def __init__(self, rel_path: str, path_a: str, path_b: str, status: str):
        self.rel_path = rel_path
        self.path_a = path_a
        self.path_b = path_b
        self.status = status

    @property
    def name(self) -> str:
        return self.rel_path.rsplit("/", 1)[-1]

    def __repr__(self):
        return f"FolderEntry({self.rel_path!r}, {self.status})"


def _is_supported(name: str) -> bool:
    return os.path.splitext(name)[1].lower() in SUPPORTED_EXTS


def scan_folder(root: str) -> dict:
    """root 이하를 재귀 순회하며 지원 확장자 파일을 수집한다.

    반환: {rel_key: (rel_display, abs_path)}
      rel_key     — 매칭용 정규화 키(소문자, '/' 구분자). 대소문자 무시 매칭.
      rel_display — 표시용 상대경로(원본 대소문자 유지, '/' 구분자).
    """
    result: dict = {}
    if not root or not os.path.isdir(root):
        return result
    root = os.path.abspath(root)
    for dirpath, _dirnames, filenames in os.walk(root):
        for fn in filenames:
            if not _is_supported(fn):
                continue
            abs_path = os.path.join(dirpath, fn)
            rel = os.path.relpath(abs_path, root).replace(os.sep, "/")
            result[rel.lower()] = (rel, abs_path)
    return result


def files_equal(path_a: str, path_b: str) -> bool:
    """두 파일이 바이트 단위로 동일한지. 크기가 다르면 즉시 False.
    청크 단위 스트리밍 비교로 **첫 불일치에서 즉시 중단**한다(구현: 해시 대신 직접 비교 —
    다른 파일을 끝까지 읽지 않아 빠름). 오류(권한/삭제 등)는 보수적으로 False(=변경 취급)."""
    try:
        if os.path.getsize(path_a) != os.path.getsize(path_b):
            return False
        with open(path_a, "rb") as fa, open(path_b, "rb") as fb:
            while True:
                ca = fa.read(_CHUNK)
                cb = fb.read(_CHUNK)
                if ca != cb:
                    return False
                if not ca:      # 둘 다 EOF (크기 동일 보장) → 완전 일치
                    return True
    except OSError:
        return False


def content_equal(path_a: str, path_b: str) -> bool:
    """지원 형식(xlsx/json/uasset)의 '논리적 내용'이 같은지 — 바이트가 다를 때만 쓰는 보정.
    - xlsx 계열: 시트 이름 목록이 같고, 각 시트의 계산값 매트릭스가 모두 같으면 True.
    - json/uasset: 비교용 매트릭스가 같으면 True.
    로드 실패/미지원 확장자는 False(=바이트 판정 유지 → modified)."""
    ext = os.path.splitext(path_a)[1].lower()
    try:
        if ext in EXCEL_EXTS:
            sheets_a = list_sheet_names(path_a)
            sheets_b = list_sheet_names(path_b)
            if sheets_a != sheets_b:
                return False
            for sh in (sheets_a or [None]):   # 시트명 취득 실패 시 첫 시트만
                if load_values_any(path_a, None, sh) != load_values_any(path_b, None, sh):
                    return False
            return True
        if ext in (".json", ".uasset", ".csv", ".tsv"):
            return load_values_any(path_a) == load_values_any(path_b)
    except Exception:
        return False
    return False


def _pair_same(path_a: str, path_b: str) -> bool:
    """파일쌍이 (논리적으로) 동일한지 — (경로,mtime,size) 캐시 경유.
    바이트가 같으면 SAME, 다르면 내용 비교로 재확인(재저장 오탐 보정). stat 실패 시 캐시 우회."""
    try:
        sa, sb = os.stat(path_a), os.stat(path_b)
        key = (path_a, sa.st_mtime, sa.st_size, path_b, sb.st_mtime, sb.st_size)
    except OSError:
        key = None
    if key is not None:
        cached = _pair_equal_cache.get(key)
        if cached is not None:
            return cached
    same = files_equal(path_a, path_b) or content_equal(path_a, path_b)
    if key is not None:
        if len(_pair_equal_cache) > _CACHE_CAP:
            _pair_equal_cache.clear()
        _pair_equal_cache[key] = same
    return same


def compare_folders(root_a: str, root_b: str, progress=None) -> list:
    """A/B 폴더를 비교해 FolderEntry 리스트를 rel_path 정렬 순으로 반환한다.

    한쪽 root가 비어 있으면(폴더 미지정) 반대쪽 파일만 only_* 로 나열한다.

    양쪽에 있는 파일쌍의 동일성 판정(바이트 비교 + 내용 비교)은 스레드 풀로 병렬화하며
    (경로,mtime,size) 캐시로 반복 스캔의 재계산을 피한다. 결과 순서는 최종 조립에서 rel_path
    정렬로 결정론적으로 유지된다.
    progress(done, total) 콜백이 주어지면 파일쌍 판정 진행 상황을 보고한다(소비 스레드 1곳에서만 호출)."""
    a_map = scan_folder(root_a)
    b_map = scan_folder(root_b)

    keys = sorted(set(a_map) | set(b_map))
    matched = [k for k in keys if k in a_map and k in b_map]

    def _classify(key):
        pa, pb = a_map[key][1], b_map[key][1]
        return key, SAME if _pair_same(pa, pb) else MODIFIED

    status_map: dict = {}
    if matched:
        total = len(matched)
        done = 0
        with ThreadPoolExecutor(max_workers=_CMP_WORKERS) as ex:
            for key, status in ex.map(_classify, matched):
                status_map[key] = status
                done += 1
                if progress is not None:
                    progress(done, total)

    entries: list = []
    for key in keys:
        a = a_map.get(key)
        b = b_map.get(key)
        if a and b:
            entries.append(FolderEntry(a[0], a[1], b[1], status_map[key]))
        elif a:
            entries.append(FolderEntry(a[0], a[1], "", ONLY_A))
        else:
            entries.append(FolderEntry(b[0], "", b[1], ONLY_B))
    return entries


def summarize(entries: list) -> dict:
    """상태별 개수 요약 — 상태바/툴팁 표시용."""
    counts = {SAME: 0, MODIFIED: 0, ONLY_A: 0, ONLY_B: 0}
    for e in entries:
        counts[e.status] = counts.get(e.status, 0) + 1
    return counts
