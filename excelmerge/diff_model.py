"""DiffTableModel — diff_matrix/staged/merged 상태를 그대로 참조해 O(1)로
표시 텍스트·색상을 파생하는 QAbstractTableModel.

기존 QTableWidget populate()의 셀 아이템 생성 로직(텍스트/배경 우선순위)을
role 계산으로 옮긴 것으로, populate는 begin/endResetModel 한 번으로 끝난다.
상태 객체(dict/set)는 MainWindow가 소유·변형하며, 변형 후 notify_* 계열을
호출해 최소 범위 dataChanged만 발생시킨다.
"""
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex
from PyQt5.QtGui import QColor
from openpyxl.utils import get_column_letter

from .theme import DIFF_COLORS, ui_font

EXTRA_ROWS = 20   # 데이터 끝에 추가할 빈 행 수
EXTRA_COLS = 5    # 데이터 끝에 추가할 빈 열 수

_ALIGN = int(Qt.AlignVCenter | Qt.AlignLeft)   # PyQt5: int 캐스팅 필수

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
        self._key_col: int = 0
        self._data_rows: int = 0
        self._data_cols: int = 0

    # ── 모드 전환 (전부 모델 리셋) ────────────────────────────────────────────
    def set_diff_data(self, diff_matrix: list, row_meta: list,
                      staged: dict, merged: set, excluded_cols: set):
        self.beginResetModel()
        self._mode = "diff"
        self._diff_matrix = diff_matrix
        self._row_meta = row_meta or []
        self._staged = staged
        self._merged = merged
        self._excluded = excluded_cols
        self._preview = []
        self._data_rows = len(diff_matrix)
        self._data_cols = len(diff_matrix[0]) if diff_matrix else 0
        self.endResetModel()

    def set_preview_data(self, data: list):
        self.beginResetModel()
        self._mode = "preview"
        self._preview = data
        self._diff_matrix = []
        self._row_meta = []
        self._data_rows = len(data)
        self._data_cols = max((len(r) for r in data), default=0)
        self.endResetModel()

    def clear(self):
        self.beginResetModel()
        self._mode = "empty"
        self._diff_matrix = []
        self._row_meta = []
        self._preview = []
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

    def last_nonempty_col(self) -> int:
        """값이 하나라도 있는 마지막 열 인덱스 (없으면 0).
        엑셀 유령 셀(서식만 있고 값 없음) 때문에 매트릭스 폭이 실제 데이터보다
        넓을 수 있어, Ctrl+Shift 헤더 확장은 이 값을 경계로 쓴다."""
        if self._mode == "diff":
            for c in range(self._data_cols - 1, -1, -1):
                for row in self._diff_matrix:
                    if c < len(row) and (row[c][1] != "" or row[c][2] != ""):
                        return c
        elif self._mode == "preview":
            for c in range(self._data_cols - 1, -1, -1):
                for row in self._preview:
                    if c < len(row) and row[c] != "":
                        return c
        return 0

    def last_nonempty_row(self) -> int:
        """값이 하나라도 있는 마지막 행 인덱스 (없으면 0)."""
        if self._mode == "diff":
            for r in range(self._data_rows - 1, -1, -1):
                if any(a != "" or b != "" for (_, a, b) in self._diff_matrix[r]):
                    return r
        elif self._mode == "preview":
            for r in range(self._data_rows - 1, -1, -1):
                if any(v != "" for v in self._preview[r]):
                    return r
        return 0

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
                return DIFF_COLORS["same"]
            if (r, c) in self._merged:
                return DIFF_COLORS["merged"]
            if (r, c) in self._staged:
                return DIFF_COLORS["staged"]
            return DIFF_COLORS[self._diff_matrix[r][c][0]]
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
                if section == self._key_col:
                    return f"🔑 {col_letter}"
                if section in self._excluded:
                    return f"⊘ {col_letter}"
                return col_letter
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
        # 세로 헤더
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
