"""백그라운드 QThread 워커 (excel_diff_merge.py에서 분리)."""
import os
import shutil
import threading
import time

from PyQt5.QtCore import QThread, pyqtSignal

from .loaders import load_values_any, load_formula_flags_any
from .diff_engine import compute_diff, count_changed, count_dropped_key_rows
from .folder_compare import compare_folders
from .xlsx_writer import _promote_empty_cols_to_delete, _write_patches_to_file
from .merge_build import build_side_patches
from .constants import DIR_A2B, DIR_B2A

_PROGRESS_INTERVAL = 0.1   # 진행률 emit 최소 간격 (초) — 상태바 스팸 방지


def _fmt_progress(done: int, total) -> str:
    if total:
        return f"{done:,}/{total:,}행"
    return f"{done:,}행"


class PreviewWorker(QThread):
    """단일 파일의 '값'만 로드해 미리보기 데이터를 돌려준다(빠른 경로)."""
    done = pyqtSignal(str, list)   # side('a'|'b'), data(값)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)         # 로딩 진행률 상태바 메시지
    progress_n = pyqtSignal(int, int)  # (done_rows, total_rows) — 진행률 바용(total<=0이면 불확정)

    def __init__(self, side: str, path: str, sheet_name=None):
        super().__init__()
        self.side = side
        self.path = path
        self.sheet_name = sheet_name

    def run(self):
        try:
            last_emit = [0.0]

            def _on_progress(done_rows, total_rows):
                now = time.monotonic()
                if now - last_emit[0] < _PROGRESS_INTERVAL:
                    return
                last_emit[0] = now
                label = "A" if self.side == "a" else "B"
                self.progress.emit(
                    f"{label} 파일 로딩 중... {_fmt_progress(done_rows, total_rows)}")
                self.progress_n.emit(int(done_rows or 0), int(total_rows or 0))

            data = load_values_any(self.path, _on_progress, self.sheet_name)
            self.done.emit(self.side, data)
        except Exception as e:
            self.error.emit(str(e))


class LoadWorker(QThread):
    """A/B 두 파일의 '값'만 병렬 로드(calamine는 GIL을 풀어 실제 병렬)."""
    done = pyqtSignal(list, list)   # a_data, b_data
    error = pyqtSignal(str)
    progress = pyqtSignal(str)
    progress_n = pyqtSignal(int, int)  # A/B 합산 (done, total) — 진행률 바용

    def __init__(self, path_a, path_b, sheet_name=None):
        super().__init__()
        self.path_a = path_a
        self.path_b = path_b
        self.sheet_name = sheet_name

    def run(self):
        try:
            from concurrent.futures import ThreadPoolExecutor

            counts = {"a": (0, None), "b": (0, None)}
            lock = threading.Lock()
            last_emit = [0.0]

            def _make_progress(side):
                def _on_progress(done_rows, total_rows):
                    with lock:
                        counts[side] = (done_rows, total_rows)
                        now = time.monotonic()
                        if now - last_emit[0] < _PROGRESS_INTERVAL:
                            return
                        last_emit[0] = now
                        parts = [
                            f"{s.upper()} {_fmt_progress(d, t)}"
                            for s, (d, t) in counts.items() if d
                        ]
                        done_sum = sum(d for d, _ in counts.values())
                        totals = [t for _, t in counts.values()]
                        # 양쪽 total 을 모두 알 때만 확정, 하나라도 미상이면 0(불확정).
                        total_sum = sum(totals) if all(t for t in totals) else 0
                    self.progress.emit("로딩 중... " + " · ".join(parts))
                    self.progress_n.emit(int(done_sum), int(total_sum))
                return _on_progress

            def _load(path, side):
                if not path:
                    return []
                return load_values_any(path, _make_progress(side), self.sheet_name)

            self.progress.emit("A/B 파일 병렬 로딩 중...")
            with ThreadPoolExecutor(max_workers=2) as ex:
                fa = ex.submit(_load, self.path_a, "a")
                fb = ex.submit(_load, self.path_b, "b")
                a = fa.result()
                b = fb.result()

            self.progress.emit("로딩 완료 — 비교 계산 중...")
            self.done.emit(a, b)
        except Exception as e:
            self.error.emit(str(e))


class DiffWorker(QThread):
    """A/B 값 매트릭스의 비교(compute_diff)와 무거운 O(R×C) 스캔(변경 셀 수·드롭 행 수)을
    백그라운드에서 수행한다 — 과거엔 _on_loaded/_recompute_diff가 이를 UI 스레드에서 동기
    실행해 대형 데이터에서 로드 직후 프리즈가 생겼다.
    token/mode를 그대로 echo해 UI가 낡은 결과를 버리고(빠른 연속 키 변경 등) 상황별 후처리를
    하도록 한다. matrix/row_meta는 동일 프로세스라 시그널로 참조 전달(복사 없음)."""
    # (token, matrix, row_meta, changed_count, dropped_count, mode)
    done = pyqtSignal(int, object, object, int, int, str)
    error = pyqtSignal(str)

    def __init__(self, a_data, b_data, key_col, key_row, excluded_cols,
                 token: int, mode: str, want_dropped: bool = True):
        super().__init__()
        self.a_data = a_data
        self.b_data = b_data
        self.key_col = key_col
        self.key_row = key_row
        self.excluded_cols = excluded_cols
        self.token = token
        self.mode = mode
        self.want_dropped = want_dropped

    def run(self):
        try:
            matrix, row_meta = compute_diff(
                self.a_data, self.b_data, self.key_col, self.key_row)
            changed = count_changed(matrix, self.excluded_cols)
            dropped = (count_dropped_key_rows(
                self.a_data, self.b_data, self.key_col, self.key_row)
                if self.want_dropped else 0)
            self.done.emit(self.token, matrix, row_meta, changed, dropped, self.mode)
        except Exception as e:
            self.error.emit(str(e))


class FormulaFlagWorker(QThread):
    """A/B 각 파일의 '수식 셀 좌표 집합'을 백그라운드로 로드한다.
    값은 calamine 캐시값으로 이미 표시되고, 이 워커는 그리드에서 수식 결과 셀을
    파랑 폰트로 구분하기 위한 좌표만 시트 XML을 직접 스캔해 로딩한다(대형 시트도 수십 ms)."""
    done = pyqtSignal(object, object)   # a_flags, b_flags — {(row0, col0)} 집합
    error = pyqtSignal(str)

    def __init__(self, path_a, path_b, sheet_name=None):
        super().__init__()
        self.path_a = path_a
        self.path_b = path_b
        self.sheet_name = sheet_name

    def run(self):
        try:
            af = load_formula_flags_any(self.path_a, self.sheet_name) if self.path_a else set()
            bf = load_formula_flags_any(self.path_b, self.sheet_name) if self.path_b else set()
            self.done.emit(af, bf)
        except Exception as e:
            self.error.emit(str(e))


class SheetDiffWorker(QThread):
    """A/B 파일의 각 시트가 서로 동일한지(값 매트릭스 비교) 백그라운드로 판정한다.
    폴더 diff처럼 '완전히 동일하지 않으면 변경'으로 보며(키 매칭 무관), 한쪽에만 있는 시트도 변경.
    변경점이 있는 시트 이름 집합을 done으로 방출 — 하단 시트 탭 노란색 표시에 사용."""
    done = pyqtSignal(object)   # set[str] — 변경점 있는 시트 이름
    error = pyqtSignal(str)

    def __init__(self, path_a, path_b, sheet_names):
        super().__init__()
        self.path_a = path_a
        self.path_b = path_b
        self.sheet_names = list(sheet_names)

    def run(self):
        try:
            changed = set()
            for name in self.sheet_names:
                try:
                    va = load_values_any(self.path_a, None, name) if self.path_a else []
                    vb = load_values_any(self.path_b, None, name) if self.path_b else []
                    if va != vb:                 # 한쪽에만 있으면 한쪽이 [] → 변경으로 잡힘
                        changed.add(name)
                except Exception:
                    changed.add(name)            # 판정 실패는 보수적으로 변경 표시
            self.done.emit(changed)
        except Exception as e:
            self.error.emit(str(e))


class FolderScanWorker(QThread):
    """폴더 비교(compare_folders)를 백그라운드에서 수행 — UI 프리즈 방지.
    파일쌍 판정 진행 상황을 progress로 보고한다(취소는 미지원)."""
    done = pyqtSignal(object)        # list[FolderEntry]
    progress = pyqtSignal(int, int)  # (done, total)
    error = pyqtSignal(str)

    def __init__(self, root_a, root_b):
        super().__init__()
        self.root_a = root_a
        self.root_b = root_b

    def run(self):
        try:
            last = [0.0]

            def _cb(done, total):
                now = time.monotonic()
                # 완료 시점(done==total)은 항상 보내고, 그 외엔 0.1s 스로틀.
                if done >= total or now - last[0] >= _PROGRESS_INTERVAL:
                    last[0] = now
                    self.progress.emit(done, total)

            self.done.emit(compare_folders(self.root_a, self.root_b, progress=_cb))
        except Exception as e:
            self.error.emit(str(e))


def copy_pairs(pairs, progress=None):
    """폴더 병합의 물리 복사 — 순수 함수(UI/Qt 비의존, 단위테스트 가능).

    pairs: [(src, dst, rel_path), ...]  — 호출부(UI 스레드)에서 미리 해결한 복사 목록.
    각 항목마다 대상 디렉터리를 만들고, 기존 대상이 있으면 .bak 백업 후 shutil.copy2 로 덮어쓴다.
    반환: (done, fails, merged_rels)
      done        : 성공 복사 수
      fails        : ["<rel>: <오류>", ...]
      merged_rels : 성공한 rel_path 목록(연파랑 '병합' 표시용)
    src 가 빈 값이면 스킵. progress(done_so_far, total) 콜백(선택).
    """
    total = len(pairs)
    done = 0
    fails = []
    merged_rels = []
    for i, (src, dst, rel) in enumerate(pairs):
        if not src:
            continue
        try:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            # 덮어쓰기 전 대상 원본을 .bak로 백업(복구 수단).
            if os.path.exists(dst):
                try:
                    shutil.copy2(dst, dst + ".bak")
                except OSError:
                    pass
            shutil.copy2(src, dst)
            done += 1
            merged_rels.append(rel)
        except OSError as ex:
            fails.append(f"{rel}: {ex}")
        if progress is not None:
            progress(i + 1, total)
    return done, fails, merged_rels


class FolderMergeWorker(QThread):
    """폴더 병합(파일 복사)을 백그라운드에서 수행 — 동기 복사로 인한 UI 프리즈 제거.
    복사 목록은 UI 스레드에서 미리 해결해 넘긴다(스레드 밖으로 FolderEntry/UI 접근 금지)."""
    done = pyqtSignal(int, object, object)  # (done_count, fails, merged_rels)
    progress = pyqtSignal(int, int)         # (done, total)
    error = pyqtSignal(str)

    def __init__(self, pairs):
        super().__init__()
        self.pairs = pairs   # [(src, dst, rel_path), ...]

    def run(self):
        try:
            last = [0.0]

            def _cb(done, total):
                now = time.monotonic()
                if done >= total or now - last[0] >= _PROGRESS_INTERVAL:
                    last[0] = now
                    self.progress.emit(done, total)

            done, fails, merged_rels = copy_pairs(self.pairs, progress=_cb)
            self.done.emit(done, fails, merged_rels)
        except Exception as e:
            self.error.emit(str(e))


class StagedMergeWorker(QThread):
    """staged(병합 준비) 셀을 파일에 기록."""
    done = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(self, path_a, path_b, diff_matrix, row_meta, staged: dict,
                 sheet_name=None, src_path=None, src_sheet_name=None):
        # staged:   {(display_r, c): 'a_to_b'|'b_to_a'}
        # row_meta: [(orig_a_row, orig_b_row), ...]
        super().__init__()
        self.path_a = path_a
        self.path_b = path_b
        self.diff_matrix = diff_matrix
        self.row_meta = row_meta
        self.staged = staged
        self.sheet_name = sheet_name   # 저장 대상 시트 (None이면 activeTab 폴백)
        # 서식 병합용 — 저장 대상의 반대편(소스) 파일 경로/시트.
        # 한 번의 저장은 한 방향만 활성(path_a 또는 path_b 하나만 채워짐)이라 단일 소스로 충분.
        self.src_path = src_path
        self.src_sheet_name = src_sheet_name

    def run(self):
        try:
            total = 0
            # a2b는 B 파일에, b2a는 A 파일에 쓴다. 방향별 패치 구성은 순수 모듈에 위임.
            for direction, path in ((DIR_A2B, self.path_b), (DIR_B2A, self.path_a)):
                patches, insert_rows, style_src = build_side_patches(
                    direction, self.diff_matrix, self.row_meta, self.staged)
                delete_rows: set[int] = set()   # 1-based 행 번호(빈행 승격 결과)
                if (patches or insert_rows) and path:
                    patches, delete_rows, del_cols = _promote_empty_cols_to_delete(
                        patches, delete_rows, path, self.sheet_name)
                    # 삭제 승격으로 사라진 patch 키의 서식 매핑 제거
                    style_src = {k: v for k, v in style_src.items() if k in patches}
                    if patches or insert_rows or delete_rows or del_cols:
                        _write_patches_to_file(
                            path, patches, list(insert_rows.values()),
                            delete_rows, del_cols, self.sheet_name,
                            self.src_path, self.src_sheet_name, style_src)
                total += len(patches) + len(insert_rows) + len(delete_rows)
            self.done.emit(total)
        except Exception as e:
            self.error.emit(str(e))
