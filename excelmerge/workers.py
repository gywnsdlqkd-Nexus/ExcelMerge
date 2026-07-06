"""백그라운드 QThread 워커 (excel_diff_merge.py에서 분리)."""
import threading
import time

from PyQt5.QtCore import QThread, pyqtSignal

from .loaders import load_sheet_with_formulas_any
from .xlsx_writer import (
    _cell_ref, _promote_empty_cols_to_delete, _write_patches_to_file,
)

_PROGRESS_INTERVAL = 0.1   # 진행률 emit 최소 간격 (초) — 상태바 스팸 방지


def _fmt_progress(done: int, total) -> str:
    if total:
        return f"{done:,}/{total:,}행"
    return f"{done:,}행"


class PreviewWorker(QThread):
    """단일 파일을 로드해 미리보기 데이터를 돌려준다."""
    done = pyqtSignal(str, list, list)   # side('a'|'b'), data(값), formula_data(수식)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)           # 로딩 진행률 상태바 메시지

    def __init__(self, side: str, path: str):
        super().__init__()
        self.side = side
        self.path = path

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

            # 통합 함수 — 같은 파일 1회 파싱으로 값/수식 모두 획득
            # xlsx/json/uasset 모두 디스패처 한 곳에서 처리
            data, formulas = load_sheet_with_formulas_any(self.path, _on_progress)
            self.done.emit(self.side, data, formulas)
        except Exception as e:
            self.error.emit(str(e))


class LoadWorker(QThread):
    done = pyqtSignal(list, list, list, list)   # a_data, b_data, a_formulas, b_formulas
    error = pyqtSignal(str)
    progress = pyqtSignal(str)                  # 단계별 상태바 메시지

    def __init__(self, path_a, path_b):
        super().__init__()
        self.path_a = path_a
        self.path_b = path_b

    def run(self):
        try:
            from concurrent.futures import ThreadPoolExecutor

            # A/B 각 로더 스레드가 진행률을 갱신하면 스로틀을 거쳐 합산 메시지로 emit.
            # (pyqtSignal emit은 스레드 세이프 — queued connection으로 전달된다)
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
                    self.progress.emit("로딩 중... " + " · ".join(parts))
                return _on_progress

            def _load(path, side):
                if not path:
                    return [], []
                return load_sheet_with_formulas_any(path, _make_progress(side))

            self.progress.emit("A/B 파일 병렬 로딩 중...")
            # A/B 두 파일을 스레드 2개로 동시 파싱. 각 xlsx는 내부에서 값/수식
            # 2패스를 다시 병렬화하므로 비교 경로는 최대 4개 파스 스레드가 돈다.
            with ThreadPoolExecutor(max_workers=2) as ex:
                fa = ex.submit(_load, self.path_a, "a")
                fb = ex.submit(_load, self.path_b, "b")
                a, af = fa.result()
                b, bf = fb.result()

            self.progress.emit("로딩 완료 — 비교 계산 중...")
            self.done.emit(a, b, af, bf)
        except Exception as e:
            self.error.emit(str(e))


class StagedMergeWorker(QThread):
    """staged(병합 준비) 셀을 파일에 기록."""
    done = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(self, path_a, path_b, diff_matrix, row_meta, staged: dict,
                 formula_data_a: list = None, formula_data_b: list = None):
        # staged:   {(display_r, c): 'a_to_b'|'b_to_a'}
        # row_meta: [(orig_a_row, orig_b_row), ...]
        super().__init__()
        self.path_a = path_a
        self.path_b = path_b
        self.diff_matrix = diff_matrix
        self.row_meta = row_meta
        self.staged = staged
        self.formula_data_a = formula_data_a or []
        self.formula_data_b = formula_data_b or []

    def _formula_val(self, formula_data: list, row_idx: int, col_idx: int, fallback: str) -> str:
        try:
            v = formula_data[row_idx][col_idx]
            return v if v != "" else fallback
        except (IndexError, TypeError):
            return fallback

    def _meta(self, display_r: int) -> tuple:
        if display_r < len(self.row_meta):
            return self.row_meta[display_r]
        return (None, None)

    def run(self):
        try:
            a2b = {k for k, v in self.staged.items() if v == "a_to_b"}
            b2a = {k for k, v in self.staged.items() if v == "b_to_a"}

            # ── B 파일에 쓸 내용 ──────────────────────────────────────────────
            patches_b: dict[str, str] = {}
            insert_rows_b: dict[int, list[tuple[int, str]]] = {}
            delete_rows_b: set[int] = set()   # 1-based 행 번호

            # a2b 병합: display_r 단위로 행 전체가 a2b 대상인지 판별
            a2b_display_rows = {r for (r, c) in a2b}
            for r in a2b_display_rows:
                row = self.diff_matrix[r]
                a_orig, b_orig = self._meta(r)
                # 해당 display 행의 모든 셀이 a2b 대상인지 확인
                row_cells_in_a2b = [(r, c) for c in range(len(row)) if (r, c) in a2b]
                if not row_cells_in_a2b:
                    continue

                if b_orig is not None:
                    # 스테이징된 셀을 patches에 추가
                    # 수식 데이터 > diff_matrix a_val 순으로 우선 사용
                    for (_, c) in row_cells_in_a2b:
                        _, a_val, _ = self.diff_matrix[r][c]
                        val = self._formula_val(self.formula_data_a, a_orig, c, a_val)
                        patches_b[_cell_ref(b_orig, c)] = val
                else:
                    # B에 행 없음(삭제됨 행) → 새 행 삽입
                    for (_, c) in row_cells_in_a2b:
                        _, a_val, _ = self.diff_matrix[r][c]
                        val = self._formula_val(self.formula_data_a, a_orig, c, a_val)
                        insert_rows_b.setdefault(r, []).append((c, val))

            # 패치 후 빈 열 감지 및 삭제 준비 (쓸 내용이 있을 때만 파일 접근)
            if (patches_b or insert_rows_b or delete_rows_b) and self.path_b:
                patches_b, delete_rows_b, del_cols_b = _promote_empty_cols_to_delete(
                    patches_b, delete_rows_b, self.path_b
                )
                if (patches_b or insert_rows_b or delete_rows_b or del_cols_b):
                    _write_patches_to_file(
                        self.path_b, patches_b,
                        list(insert_rows_b.values()),
                        delete_rows_b,
                        del_cols_b,
                    )

            # ── A 파일에 쓸 내용 ──────────────────────────────────────────────
            patches_a: dict[str, str] = {}
            insert_rows_a: dict[int, list[tuple[int, str]]] = {}
            delete_rows_a: set[int] = set()

            b2a_display_rows = {r for (r, c) in b2a}
            for r in b2a_display_rows:
                row = self.diff_matrix[r]
                a_orig, b_orig = self._meta(r)
                row_cells_in_b2a = [(r, c) for c in range(len(row)) if (r, c) in b2a]
                if not row_cells_in_b2a:
                    continue

                if a_orig is not None:
                    # 스테이징된 셀을 patches에 추가
                    # 수식 데이터 > diff_matrix b_val 순으로 우선 사용
                    for (_, c) in row_cells_in_b2a:
                        _, _, b_val = self.diff_matrix[r][c]
                        val = self._formula_val(self.formula_data_b, b_orig, c, b_val)
                        patches_a[_cell_ref(a_orig, c)] = val
                else:
                    # A에 행 없음(추가됨 행) → 새 행 삽입
                    for (_, c) in row_cells_in_b2a:
                        _, _, b_val = self.diff_matrix[r][c]
                        val = self._formula_val(self.formula_data_b, b_orig, c, b_val)
                        insert_rows_a.setdefault(r, []).append((c, val))

            if (patches_a or insert_rows_a or delete_rows_a) and self.path_a:
                patches_a, delete_rows_a, del_cols_a = _promote_empty_cols_to_delete(
                    patches_a, delete_rows_a, self.path_a
                )
                if (patches_a or insert_rows_a or delete_rows_a or del_cols_a):
                    _write_patches_to_file(
                        self.path_a, patches_a,
                        list(insert_rows_a.values()),
                        delete_rows_a,
                        del_cols_a,
                    )

            total = (len(patches_b) + len(insert_rows_b) + len(delete_rows_b)
                     + len(patches_a) + len(insert_rows_a) + len(delete_rows_a))
            self.done.emit(total)
        except Exception as e:
            self.error.emit(str(e))
