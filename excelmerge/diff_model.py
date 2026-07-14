"""DiffTableModel — diff_matrix/staged/merged 상태를 그대로 참조해 O(1)로
표시 텍스트·색상을 파생하는 QAbstractTableModel.

기존 QTableWidget populate()의 셀 아이템 생성 로직(텍스트/배경 우선순위)을
role 계산으로 옮긴 것으로, populate는 begin/endResetModel 한 번으로 끝난다.
상태 객체(dict/set)는 MainWindow가 소유·변형하며, 변형 후 notify_* 계열을
호출해 최소 범위 dataChanged만 발생시킨다.
"""
import difflib

from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex
from PyQt5.QtGui import QColor
from openpyxl.utils import get_column_letter

from .theme import (
    DIFF_COLORS, EXCLUDED_CELL_BG, CELL_FORMULA_FG, ui_font,
    key_header_icon, exclude_header_icon,
)

EXTRA_ROWS = 20   # 데이터 끝에 추가할 빈 행 수
EXTRA_COLS = 5    # 데이터 끝에 추가할 빈 열 수

_ALIGN = int(Qt.AlignVCenter | Qt.AlignLeft)   # PyQt5: int 캐스팅 필수


def _char_diff_ranges(own: str, other: str) -> list:
    """own 문자열에서 other와 다른 [start, end) 구간 목록.
    SequenceMatcher의 replace/delete 구간(own측 i1:i2)만 수집 —
    insert(other에만 있는 부분)는 own에 표시할 구간이 없으므로 반대쪽에서 강조된다."""
    ranges = []
    sm = difflib.SequenceMatcher(None, own, other, autojunk=False)
    for tag, i1, i2, _j1, _j2 in sm.get_opcodes():
        if tag in ("replace", "delete") and i1 < i2:
            ranges.append((i1, i2))
    return ranges

# 헤더 색 — 기존 _refresh_key_col_header와 동일 값
_HDR_KEY_BG = QColor(255, 213, 0)
_HDR_EXCL_BG = QColor(220, 220, 220)
_HDR_NORMAL_BG = QColor(232, 234, 240)
_HDR_BLACK = QColor(0, 0, 0)
_HDR_GRAY = QColor(140, 140, 140)


class DiffTableModel(QAbstractTableModel):
    def __init__(self, side: str, parent=None):
        super().__init__(parent)
        self.side = side                     # 'a' | 'b'
        self._mode = "empty"                 # 'empty' | 'diff' | 'preview'
        self._diff_matrix: list = []
        self._row_meta: list = []
        self._staged: dict = {}
        self._merged: set = set()
        self._excluded: set = set()
        self._preview: list = []
        # 이 side의 '수식 셀' 원본 좌표 집합 {(row0, col0)}. 지연 로딩되며
        # 비면(로드 전/json·uasset) 모든 셀을 비수식으로 취급한다.
        self._formula_flags: set = set()
        self._key_col: int = 0
        self._key_row: int = 0
        self._data_rows: int = 0
        self._data_cols: int = 0
        # (r,c) -> _char_diff_ranges 결과 메모이즈. 순수 SequenceMatcher(own,other) 결과만
        # 담고 own/other는 리셋 없이는 안 바뀌므로 안전. 리셋 경로에서만 무효화한다.
        # (staged/merged 게이팅은 diff_char_ranges가 조회 전에 재평가 → 스테일 없음)
        self._char_range_cache: dict = {}

    # ── 모드 전환 (전부 모델 리셋) ────────────────────────────────────────────
    def set_diff_data(self, diff_matrix: list, row_meta: list,
                      staged: dict, merged: set, excluded_cols: set):
        self.beginResetModel()
        self._char_range_cache.clear()
        self._mode = "diff"
        self._diff_matrix = diff_matrix
        self._row_meta = row_meta or []
        self._staged = staged
        self._merged = merged
        self._excluded = excluded_cols
        self._preview = []
        self._formula_flags = []   # 새 비교 — 수식 플래그는 지연 로딩으로 다시 채운다
        self._data_rows = len(diff_matrix)
        self._data_cols = len(diff_matrix[0]) if diff_matrix else 0
        self.endResetModel()

    def set_preview_data(self, data: list):
        self.beginResetModel()
        self._char_range_cache.clear()
        self._mode = "preview"
        self._preview = data
        self._diff_matrix = []
        self._row_meta = []
        self._formula_flags = []
        self._data_rows = len(data)
        self._data_cols = max((len(r) for r in data), default=0)
        self.endResetModel()

    def set_formula_flags(self, flags: list):
        """수식 여부 매트릭스를 지연 세팅(백그라운드 로딩 완료 시).
        is_formula가 live로 참조하므로 리셋 없이 교체만 하고, 그리드 전체를 다시 칠한다."""
        self._formula_flags = flags or []
        if self._data_rows and self._data_cols:
            # 역할 지정 없이(전체) 방출 — 델리게이트가 is_formula로 파랑/기본을 다시 칠한다.
            self.dataChanged.emit(
                self.index(0, 0),
                self.index(self._data_rows - 1, self._data_cols - 1))

    def clear(self):
        self.beginResetModel()
        self._char_range_cache.clear()
        self._mode = "empty"
        self._diff_matrix = []
        self._row_meta = []
        self._preview = []
        self._formula_flags = []
        self._data_rows = 0
        self._data_cols = 0
        self.endResetModel()

    # ── 부분 갱신 — 상태는 이미 MainWindow가 변형한 뒤 호출된다 ───────────────
    def notify_cells(self, cells):
        """변경된 (r, c) 집합을 행별 col-span으로 병합해 dataChanged 방출."""
        if not cells:
            return
        by_row: dict[int, list[int]] = {}
        for (r, c) in cells:
            if 0 <= r < self._data_rows and 0 <= c < self._data_cols:
                by_row.setdefault(r, []).append(c)
        roles = [Qt.DisplayRole, Qt.BackgroundRole]
        if len(by_row) > 200:
            # 대량 변경은 단일 bounding rect가 더 저렴
            rows = sorted(by_row)
            cols = [c for cs in by_row.values() for c in cs]
            self.dataChanged.emit(
                self.index(rows[0], min(cols)), self.index(rows[-1], max(cols)), roles)
            return
        for r, cs in by_row.items():
            self.dataChanged.emit(
                self.index(r, min(cs)), self.index(r, max(cs)), roles)

    def notify_columns(self, cols):
        """열 전체 갱신 (제외 토글 등) + 가로 헤더 갱신."""
        last_row = max(0, self._data_rows - 1)
        roles = [Qt.DisplayRole, Qt.BackgroundRole]
        for c in cols:
            if 0 <= c < self._data_cols:
                self.dataChanged.emit(self.index(0, c), self.index(last_row, c), roles)
            if 0 <= c < self.columnCount():
                self.headerDataChanged.emit(Qt.Horizontal, c, c)

    def set_key_col(self, col: int):
        old = self._key_col
        self._key_col = col
        for c in (old, col):
            if 0 <= c < self.columnCount():
                self.headerDataChanged.emit(Qt.Horizontal, c, c)

    def set_key_row(self, row: int):
        old = self._key_row
        self._key_row = row
        for r in (old, row):
            if 0 <= r < self.rowCount():
                self.headerDataChanged.emit(Qt.Vertical, r, r)

    def set_excluded_cols(self, cols: set):
        changed = self._excluded ^ set(cols)
        self._excluded = set(cols)
        self.notify_columns(sorted(changed))

    # ── 조회 API — 기존 item()/색상 스니핑 호출부의 대체 ──────────────────────
    @property
    def data_rows(self) -> int:
        return self._data_rows

    @property
    def data_cols(self) -> int:
        return self._data_cols

    def is_data_cell(self, r: int, c: int) -> bool:
        return 0 <= r < self._data_rows and 0 <= c < self._data_cols

    def orig_row(self, r: int) -> int | None:
        """display 행 r 의 원본 파일 행 번호(0-based). A 우선, 없으면 B, 둘 다 없거나
        범위 밖이면 None. 키 행 지정 시 display→원본 매핑에 쓴다."""
        if 0 <= r < len(self._row_meta):
            a_idx, b_idx = self._row_meta[r]
            return a_idx if a_idx is not None else b_idx
        return None

    def col_has_values(self, c: int) -> bool:
        """열에 값이 하나라도 있는지 (A/B 어느 쪽이든).
        엑셀 유령 셀(서식만 있고 값 없음) 열과 여분 열은 False —
        헤더 단위 Ctrl+Shift 점프의 빈/값 판정에 쓴다."""
        if self._mode == "diff":
            for row in self._diff_matrix:
                if c < len(row) and (row[c][1] != "" or row[c][2] != ""):
                    return True
            return False
        if self._mode == "preview":
            for row in self._preview:
                if c < len(row) and row[c] != "":
                    return True
        return False

    def row_has_values(self, r: int) -> bool:
        """행에 값이 하나라도 있는지 (A/B 어느 쪽이든)."""
        if not (0 <= r < self._data_rows):
            return False
        if self._mode == "diff":
            return any(a != "" or b != "" for (_, a, b) in self._diff_matrix[r])
        if self._mode == "preview":
            return any(v != "" for v in self._preview[r])
        return False

    def display_text(self, r: int, c: int) -> str:
        if not self.is_data_cell(r, c):
            return ""
        if self._mode == "preview":
            row = self._preview[r]
            return row[c] if c < len(row) else ""
        _, a_val, b_val = self._diff_matrix[r][c]
        direction = self._staged.get((r, c))
        if direction == "a_to_b":
            return a_val
        if direction == "b_to_a":
            return b_val
        return a_val if self.side == "a" else b_val

    def cell_kind(self, r: int, c: int) -> str:
        """'changed' | 'staged' | 'merged' | 'same' — 배경색 스니핑 대체.
        우선순위는 populate 색상 로직과 동일: 제외 > merged > staged > status."""
        if self._mode != "diff" or not self.is_data_cell(r, c):
            return "same"
        if c in self._excluded:
            return "same"
        if (r, c) in self._merged:
            return "merged"
        if (r, c) in self._staged:
            return "staged"
        status = self._diff_matrix[r][c][0]
        return "changed" if status in ("added", "modified") else "same"

    def _formula_orig_row(self, r: int):
        """이 side에서 (표시) 행 r에 대응하는 원본(수식 플래그 매트릭스) 행 인덱스.
        diff 모드는 row_meta[side]로, preview는 그대로. 대응 없음이면 None."""
        if self._mode == "diff":
            side_idx = 0 if self.side == "a" else 1
            try:
                return self._row_meta[r][side_idx]
            except (IndexError, TypeError):
                return r
        return r   # preview — 표시 행 == 원본 행

    def is_formula(self, r: int, c: int) -> bool:
        """이 side의 (표시) 셀 값이 원본에서 수식으로 계산된 값인지.
        _formula_flags는 원본 좌표 {(row0, col0)} 집합. 표시되는 값이 이 side의 값일 때만
        의미가 있으므로, 병합 준비/완료로 반대 side 값을 보여주는 셀은 제외한다."""
        if not self._formula_flags or not self.is_data_cell(r, c):
            return False
        if self._mode == "diff" and (
                (r, c) in self._merged or (r, c) in self._staged):
            return False
        orig = self._formula_orig_row(r)
        if orig is None:
            return False
        return (orig, c) in self._formula_flags

    def staged_coords(self):
        """병합 준비(staged) 셀 좌표 {(r, c)} 뷰 — 보통 소수.
        대량 선택에서 'staged 셀 포함 여부'를 전 셀 순회 없이 O(#staged)로 판정하는 데 쓴다."""
        return self._staged.keys()

    def diff_char_ranges(self, r: int, c: int) -> list:
        """이 side 표시 값에서 반대 side와 다른 [start,end) 문자 구간.
        '값이 서로 다른' modified 셀에서만 산출 — same/added(한쪽만)/staged/merged/
        제외열/preview/여분 셀은 []( 강조 없음). 그리드 델리게이트·셀값란 공통 소스."""
        if self._mode != "diff" or not self.is_data_cell(r, c):
            return []
        if c in self._excluded or (r, c) in self._merged or (r, c) in self._staged:
            return []
        status, a_val, b_val = self._diff_matrix[r][c]
        if status != "modified" or a_val == b_val:
            return []
        cached = self._char_range_cache.get((r, c))
        if cached is None:   # []도 유효한 결과 → is None 센티넬로 미계산과 구분
            own, other = (a_val, b_val) if self.side == "a" else (b_val, a_val)
            cached = _char_diff_ranges(own, other)
            self._char_range_cache[(r, c)] = cached
        return cached

    def is_added_placeholder(self, r: int, c: int) -> bool:
        """이 side가 '신규(added)' 셀의 빈 쪽인지 — 반대 side에만 값이 있는 칸.
        그리드 델리게이트가 이 칸에 대각선 해치를 그려 매칭을 표시한다."""
        if self._mode != "diff" or not self.is_data_cell(r, c):
            return False
        if c in self._excluded or (r, c) in self._merged or (r, c) in self._staged:
            return False
        status, a_val, b_val = self._diff_matrix[r][c]
        if status != "added":
            return False
        own = a_val if self.side == "a" else b_val
        return own == ""

    # ── Qt 오버라이드 ─────────────────────────────────────────────────────────
    def rowCount(self, parent=QModelIndex()):
        if parent.isValid() or self._mode == "empty":
            return 0
        return self._data_rows + EXTRA_ROWS

    def columnCount(self, parent=QModelIndex()):
        if parent.isValid() or self._mode == "empty":
            return 0
        return self._data_cols + EXTRA_COLS

    def flags(self, index):
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r, c = index.row(), index.column()
        if role in (Qt.DisplayRole, Qt.EditRole):
            text = self.display_text(r, c)
            return text if text != "" else None
        if role == Qt.BackgroundRole:
            if self._mode != "diff" or not self.is_data_cell(r, c):
                return None   # 기본(흰색) — 기존 아이템 없는 여분 셀과 동일
            if c in self._excluded:
                return EXCLUDED_CELL_BG   # 제외 열은 회색 배경으로 구분
            if (r, c) in self._merged:
                return DIFF_COLORS["merged"]
            if (r, c) in self._staged:
                return DIFF_COLORS["staged"]
            status, a_val, b_val = self._diff_matrix[r][c]
            if status == "added":
                # 신규(연초록)는 값이 실제로 있는 쪽 패널에만 표시.
                # A에만 값 → A패널만 초록, B에만 값 → B패널만 초록. 빈 쪽은 흰색.
                own_val = a_val if self.side == "a" else b_val
                return DIFF_COLORS["added"] if own_val != "" else DIFF_COLORS["same"]
            return DIFF_COLORS[status]
        if role == Qt.ForegroundRole:
            # 수식으로 계산된 값 셀은 파랑 폰트로 구분(기본 델리게이트가 비선택 셀에 적용).
            # 변경 문자 구간의 빨강은 델리게이트(diff_char_ranges 경로)가 별도로 덮어쓴다.
            return CELL_FORMULA_FG if self.is_formula(r, c) else None
        if role == Qt.TextAlignmentRole:
            return _ALIGN
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal:
            col_letter = get_column_letter(section + 1)
            # 미리보기/빈 상태: 기존과 동일하게 색·아이콘 없는 순수 라벨
            if self._mode != "diff":
                return col_letter if role == Qt.DisplayRole else None
            if role == Qt.DisplayRole:
                # 아이콘(DecorationRole)으로 표시하므로 텍스트는 열 문자만.
                return col_letter
            if role == Qt.DecorationRole:
                if section == self._key_col:
                    return key_header_icon()
                if section in self._excluded:
                    return exclude_header_icon()
                return None
            if role == Qt.BackgroundRole:
                if section == self._key_col:
                    return _HDR_KEY_BG
                if section in self._excluded:
                    return _HDR_EXCL_BG
                return _HDR_NORMAL_BG
            if role == Qt.ForegroundRole:
                return _HDR_GRAY if (section in self._excluded
                                     and section != self._key_col) else _HDR_BLACK
            if role == Qt.FontRole:
                return ui_font(9, bold=(section == self._key_col))
            return None
        # 세로 헤더 — 키 행(헤더 행)이면 아이콘/노랑 배경/볼드(가로 키 열과 대칭)
        if (self._mode == "diff" and self._row_meta
                and section < self._data_rows and section == self._key_row):
            if role == Qt.DecorationRole:
                return key_header_icon()
            if role == Qt.BackgroundRole:
                return _HDR_KEY_BG
            if role == Qt.FontRole:
                return ui_font(9, bold=True)
        if role != Qt.DisplayRole:
            return None
        if self._mode == "diff" and self._row_meta:
            if section < self._data_rows:
                if section < len(self._row_meta):
                    side_idx = 0 if self.side == "a" else 1
                    orig = self._row_meta[section][side_idx]
                    return str(orig + 1) if orig is not None else "-"
                return "-"
            return str(section + 1)
        return str(section + 1)
