"""3파일 비교 백그라운드 워커 — 읽기·비교·시트 변경 판정·저장.

2파일 워커(workers.py)와 같은 틀이다: 결과는 시그널로 참조 전달, 비교는 token/mode 를 그대로
돌려줘 화면이 낡은 결과를 버린다.
"""
import os
import shutil
import threading
import time

from PyQt5.QtCore import QThread, pyqtSignal

from .diff3_engine import compute_diff3, conflict_masks, count_dropped3
from .diff_engine import count_changed_masked, keep_mask, row_change_masks
from .loaders import load_values_any
from .merge3_build import build_merge3_patches
from .xlsx_writer import _is_file_locked, _promote_empty_cols_to_delete, _write_patches_to_file

_PROGRESS_INTERVAL = 0.1
_LABELS = ("왼쪽", "가운데", "오른쪽")


class Load3Worker(QThread):
    """세 파일의 값을 병렬로 읽는다(calamine 은 GIL 을 풀어 실제로 병렬)."""
    done = pyqtSignal(list, list, list)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, paths, sheet_name=None):
        super().__init__()
        self.paths = tuple(paths)          # (왼쪽, 가운데 원본 사본, 오른쪽)
        self.sheet_name = sheet_name

    def run(self):
        try:
            from concurrent.futures import ThreadPoolExecutor
            counts = [0, 0, 0]
            lock = threading.Lock()
            last = [0.0]

            def loader(i):
                def on_progress(done_rows, _total):
                    with lock:
                        counts[i] = done_rows or 0
                        now = time.monotonic()
                        if now - last[0] < _PROGRESS_INTERVAL:
                            return
                        last[0] = now
                        parts = [f"{_LABELS[k]} {counts[k]:,}행" for k in range(3) if counts[k]]
                    self.progress.emit("로딩 중... " + " · ".join(parts))

                def run_one():
                    p = self.paths[i]
                    return load_values_any(p, on_progress, self.sheet_name) if p else []
                return run_one

            self.progress.emit("세 파일 읽는 중...")
            with ThreadPoolExecutor(max_workers=3) as ex:
                futs = [ex.submit(loader(i)) for i in range(3)]
                data = [f.result() for f in futs]
            self.done.emit(*data)
        except Exception as e:
            self.error.emit(str(e))


class Diff3Worker(QThread):
    # (token, matrix, row_meta3, col_meta3, row_masks, conflict_masks, changed, dropped, mode,
    #  ambiguous — 내용으로 짝지은 화면 행 집합)
    done = pyqtSignal(int, object, object, object, object, object, int, int, str, object)
    error = pyqtSignal(str)

    def __init__(self, left, base, right, key_col, key_row, excluded_cols, token, mode,
                 col_meta3=None):
        super().__init__()
        self.data = (left, base, right)
        self.key_col = key_col
        self.key_row = key_row
        self.excluded_cols = set(excluded_cols)
        self.token = token
        self.mode = mode
        self.col_meta3 = col_meta3

    def run(self):
        try:
            ambiguous = set()
            matrix, row_meta = compute_diff3(*self.data, self.key_col, self.key_row,
                                             self.col_meta3, ambiguous)
            masks = row_change_masks(matrix)
            cmasks = conflict_masks(matrix)
            cols = len(matrix[0]) if matrix else 0
            changed = count_changed_masked(masks, keep_mask(cols, self.excluded_cols))
            dropped = count_dropped3(*self.data, self.col_meta3, self.key_col, self.key_row)
            self.done.emit(self.token, matrix, row_meta, self.col_meta3, masks, cmasks,
                           changed, dropped, self.mode, ambiguous)
        except Exception as e:
            self.error.emit(str(e))


class Sheet3DiffWorker(QThread):
    """시트마다 '왼쪽이나 오른쪽이 가운데와 다른가'(값 매트릭스 비교, 키 무관). 하단 탭 색칠과
    '열어 보지 않은 시트에 변경이 있다' 안내에 쓴다."""
    done = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(self, paths, sheet_names):
        super().__init__()
        self.paths = tuple(paths)
        self.sheet_names = list(sheet_names)

    def run(self):
        try:
            changed = set()
            for name in self.sheet_names:
                try:
                    vl, vb, vr = (load_values_any(p, None, name) if p else []
                                  for p in self.paths)
                    if vl != vb or vr != vb:
                        changed.add(name)
                except Exception:
                    changed.add(name)        # 판정 실패는 보수적으로 변경
            self.done.emit(changed)
        except Exception as e:
            self.error.emit(str(e))


def write_merge3(template: str, out_path: str, sheets: list, sources: dict) -> int:
    """결과 = 가운데 원본(template) + 모든 시트의 결정. 쓴 칸·행 수를 돌려준다.

    sheets  : [(시트 이름, matrix, row_meta3, col_meta3, decisions)] — 결정이 없는 시트도 된다.
    sources : {SRC_LEFT: 경로, SRC_RIGHT: 경로} — 서식을 가져올 파일.

    **언제나 원본에서 새로 만든다.** 결정 좌표는 원본 기준이라, 이미 고친 결과 파일 위에 다시
    쓰면 지운 행·붙인 행만큼 좌표가 어긋난다(2파일에서 실측된 '두 번째 저장이 엉뚱한 행에'
    사고). 그래서 원본을 옆 작업 파일로 떠서 시트마다 차례로 고치고, 끝까지 성공했을 때만
    결과 파일과 바꾼다 — 중간에 실패하면 결과 파일은 그대로다.
    """
    target = out_path
    if _is_file_locked(target):
        raise PermissionError(f"파일을 저장할 수 없습니다.\n{target}")
    # 저장할 때마다 다른 이름 — 값 캐시 열쇠가 (경로, 수정 시각, 시트)라 같은 이름을 다시 쓰면
    # 수정 시각이 겹칠 때 지난번 작업 파일의 값을 '저장 전 값'으로 읽을 수 있다.
    stage = f"{target}.em3_{os.getpid()}_{time.monotonic_ns()}"
    total = 0
    try:
        shutil.copyfile(template, stage)
        for name, matrix, row_meta, col_meta, decisions in sheets:
            if not decisions:
                continue
            patches, inserts, style_src = build_merge3_patches(
                matrix, row_meta, col_meta, decisions)
            insert_list = list(inserts.values())
            patches, delete_rows, del_cols = _promote_empty_cols_to_delete(
                patches, set(), stage, name, insert_rows=insert_list)
            style_src = {k: v for k, v in style_src.items() if k in patches}
            if patches or insert_list or delete_rows or del_cols:
                src = {tag: (p, name) for tag, p in sources.items() if p}
                _write_patches_to_file(stage, patches, insert_list, delete_rows, del_cols,
                                       name, src, None, style_src)
            total += len(patches) + len(insert_list) + len(delete_rows)
        os.replace(stage, target)
    except Exception:
        if os.path.exists(stage):
            os.remove(stage)
        raise
    return total


class Merge3Worker(QThread):
    done = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(self, template, out_path, sheets, sources):
        super().__init__()
        self.args = (template, out_path, sheets, sources)

    def run(self):
        try:
            self.done.emit(write_merge3(*self.args))
        except Exception as e:
            self.error.emit(str(e))
