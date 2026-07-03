"""테이블/스크롤/입력 위젯 (excel_diff_merge.py에서 분리)."""
import os

from PyQt5.QtWidgets import (
    QApplication, QTableView, QAbstractItemView, QLineEdit, QPlainTextEdit,
    QHeaderView, QMenu, QStyle, QScrollBar, QStyleOptionSlider,
)
from PyQt5.QtCore import (
    Qt, pyqtSignal, QItemSelection, QItemSelectionRange, QItemSelectionModel,
)
from PyQt5.QtGui import QPainter
from openpyxl.utils import get_column_letter

from .diff_model import DiffTableModel
from .diff_model import EXTRA_ROWS as EXTRA_ROWS, EXTRA_COLS as EXTRA_COLS  # 재노출
from .loaders import _SUPPORTED_EXTS
from .theme import MENU_QSS, MINIMAP_MARKER_COLOR, ui_font


# 자동 컬럼 너비 상한 — 150px
# 데이터가 긴 셀 때문에 열이 화면을 가리지 않도록 제한.
# 사용자가 헤더 드래그로 직접 넓힌 열은 _user_col_widths 에 기록되어 이 상한 무시.
# 새로고침 시에는 _run_refresh()가 _user_col_widths를 비우므로 모든 열이 디폴트로 복귀.
MAX_AUTO_COL_WIDTH_PX = 150


class MinimapScrollBar(QScrollBar):
    """수직 또는 수평 스크롤바 위에 변경된 셀(행/열)의 위치를 색상 마커로 오버레이.
    paintEvent에서 super 호출 후, orientation에 맞춰 트랙(groove) 영역에
    비율 위치(0.0~1.0)별로 가는 막대를 그린다."""
    _MARKER_COLOR = MINIMAP_MARKER_COLOR

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self._ratios: list = []   # 0.0~1.0 사이 변경 위치 목록

    def set_change_ratios(self, ratios):
        # 변경된 경우에만 repaint (불필요한 페인트 방지)
        if list(ratios) != self._ratios:
            self._ratios = list(ratios)
            self.update()

    def paintEvent(self, e):
        super().paintEvent(e)
        if not self._ratios:
            return
        # QStyle을 통해 정확한 trough(groove) 영역을 얻는다
        opt = QStyleOptionSlider()
        self.initStyleOption(opt)
        groove = self.style().subControlRect(
            QStyle.CC_ScrollBar, opt, QStyle.SC_ScrollBarGroove, self)
        if groove.width() <= 0 or groove.height() <= 0:
            return
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, False)
        painter.setPen(Qt.NoPen)
        painter.setBrush(self._MARKER_COLOR)
        if self.orientation() == Qt.Vertical:
            track_top = groove.top()
            track_h = groove.height()
            x = groove.left() + 2
            w = max(1, groove.width() - 4)
            denom = max(0, track_h - 2)
            for ratio in self._ratios:
                y = track_top + int(ratio * denom)
                painter.fillRect(x, y, w, 2, self._MARKER_COLOR)
        else:
            track_left = groove.left()
            track_w = groove.width()
            y = groove.top() + 2
            h = max(1, groove.height() - 4)
            denom = max(0, track_w - 2)
            for ratio in self._ratios:
                x = track_left + int(ratio * denom)
                painter.fillRect(x, y, 2, h, self._MARKER_COLOR)
        painter.end()


class ExcelTableView(QTableView):
    stage_requested   = pyqtSignal(str)   # direction: 'a_to_b' | 'b_to_a'
    unstage_requested = pyqtSignal()
    key_col_changed   = pyqtSignal(int)   # 키 열 변경 요청
    columns_exclude_set = pyqtSignal(list, bool)   # (cols, exclude) — True: 제외 추가, False: 제외 해제
    column_resized    = pyqtSignal(int, int)   # (col, new_width) — 사용자 조작에 의한 변경만
    row_resized       = pyqtSignal(int, int)   # (row, new_height) — 사용자 조작에 의한 변경만
    edit_focus_requested = pyqtSignal()   # F2 — 패널 cell_edit 으로 포커스 이동 요청
    delete_cell_requested = pyqtSignal(int, int)   # (row, col) — Delete 키로 셀 값 비우기 요청

    def __init__(self, side: str, parent=None):
        super().__init__(parent)
        self.side = side
        self._model = DiffTableModel(side, self)
        self.setModel(self._model)   # selectionModel은 여기서 1회 생성 — 이후 교체 없음
        self._populating = False
        self._key_col: int = 0
        self._excluded_cols: set[int] = set()   # 변경 검사 제외 열 (display 인덱스)
        # 사용자가 직접 조정한 열/행 크기 — 세션 동안만 유지 (재로드/저장/새로고침 후 복원)
        self._user_col_widths: dict[int, int] = {}
        self._user_row_heights: dict[int, int] = {}
        # 외부(다른 패널)에서 크기를 강제 적용 중일 때 sectionResized 재방출 방지
        self._applying_sizes: bool = False
        # 헤더 다중 선택의 anchor (Shift+방향 확장의 고정점)
        self._header_anchor_col: int | None = None
        self._header_anchor_row: int | None = None
        self.setFont(ui_font(9))
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.verticalHeader().setDefaultSectionSize(22)
        self.setAlternatingRowColors(False)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_context_menu)
        self.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.horizontalHeader().customContextMenuRequested.connect(
            self._show_header_context_menu)
        self.verticalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.verticalHeader().customContextMenuRequested.connect(
            self._show_row_header_context_menu)
        # 헤더 크기 변경 추적 — 사용자 조작 시에만 저장/시그널 발행
        self.horizontalHeader().sectionResized.connect(self._on_section_h_resized)
        self.verticalHeader().sectionResized.connect(self._on_section_v_resized)
        # 헤더 클릭 시 anchor 갱신 (Shift 없는 클릭 → 새 anchor / Shift 클릭 → 기존 유지)
        self.horizontalHeader().sectionPressed.connect(self._on_h_section_pressed)
        self.verticalHeader().sectionPressed.connect(self._on_v_section_pressed)

    # ── QTableWidget 호환 헬퍼 ───────────────────────────────────────────────
    def rowCount(self) -> int:
        return self._model.rowCount()

    def columnCount(self) -> int:
        return self._model.columnCount()

    def _current_cell(self) -> tuple:
        """QTableWidget.currentRow()/currentColumn() 대응 — 무효 시 (-1, -1)."""
        idx = self.currentIndex()
        return (idx.row(), idx.column()) if idx.isValid() else (-1, -1)

    def _set_current_cell(self, r: int, c: int):
        """QTableWidget.setCurrentCell() 대응.
        setCurrentIndex()는 호출 시점의 키보드 수정자에 따라 선택을 확장/클리어하는
        기존 setCurrentCell과 완전히 같은 경로(selectionCommand)를 탄다."""
        if 0 <= r < self.rowCount() and 0 <= c < self.columnCount():
            self.setCurrentIndex(self._model.index(r, c))

    # ── 사용자 헤더 크기 추적 ────────────────────────────────────────────────
    def _on_section_h_resized(self, logical_index: int, _old: int, new_size: int):
        # populate 중이거나 다른 패널에서 강제 적용 중인 변경은 무시
        if self._populating or self._applying_sizes:
            return
        self._user_col_widths[logical_index] = new_size
        self.column_resized.emit(logical_index, new_size)

    def _on_section_v_resized(self, logical_index: int, _old: int, new_size: int):
        if self._populating or self._applying_sizes:
            return
        self._user_row_heights[logical_index] = new_size
        self.row_resized.emit(logical_index, new_size)

    def _on_h_section_pressed(self, logical_index: int):
        """열 헤더 클릭: Shift/Ctrl 없으면 anchor 갱신, 동반이면 유지."""
        mods = QApplication.keyboardModifiers()
        if not (mods & (Qt.ShiftModifier | Qt.ControlModifier)):
            self._header_anchor_col = logical_index
        elif self._header_anchor_col is None:
            # Shift+클릭인데 anchor가 없으면 현재 클릭 지점을 anchor로
            self._header_anchor_col = logical_index
        # 행 anchor는 무관 — 열 헤더 클릭은 행 헤더 모드를 종료시킴
        self._header_anchor_row = None

    def _on_v_section_pressed(self, logical_index: int):
        mods = QApplication.keyboardModifiers()
        if not (mods & (Qt.ShiftModifier | Qt.ControlModifier)):
            self._header_anchor_row = logical_index
        elif self._header_anchor_row is None:
            self._header_anchor_row = logical_index
        self._header_anchor_col = None

    def apply_column_width(self, col: int, width: int):
        """반대 패널에서의 열 너비 변경을 동기 적용 (시그널 재방출 안 함)."""
        self._user_col_widths[col] = width
        if 0 <= col < self.columnCount() and self.columnWidth(col) != width:
            self._applying_sizes = True
            try:
                self.setColumnWidth(col, width)
            finally:
                self._applying_sizes = False

    def apply_row_height(self, row: int, height: int):
        """반대 패널에서의 행 높이 변경을 동기 적용."""
        self._user_row_heights[row] = height
        if 0 <= row < self.rowCount() and self.rowHeight(row) != height:
            self._applying_sizes = True
            try:
                self.setRowHeight(row, height)
            finally:
                self._applying_sizes = False

    def _apply_user_sizes(self):
        """저장된 사용자 크기를 현재 테이블에 다시 적용 (populate 후 호출)."""
        # 0-크기 값은 hidden 행/열에 대한 sectionResized 시그널이 남긴 오염값일 수
        # 있으므로 무시한다. UI상 0으로 만드는 사용자 조작은 없다.
        self._applying_sizes = True
        try:
            for col, w in self._user_col_widths.items():
                if 0 <= col < self.columnCount() and w > 0:
                    self.setColumnWidth(col, w)
            for row, h in self._user_row_heights.items():
                if 0 <= row < self.rowCount() and h > 0:
                    self.setRowHeight(row, h)
        finally:
            self._applying_sizes = False

    def _auto_size_columns(self, max_samples: int = 200):
        """resizeColumnsToContents() 대체 — 셀 텍스트를 샘플링해 QFontMetrics로 측정.
        모든 자동 폭이 MAX_AUTO_COL_WIDTH_PX로 클립되므로 상한 도달 시 조기 종료해
        전체 셀 스캔을 피한다. sectionResized가 사용자 변경으로 기록되지 않도록
        _applying_sizes 플래그로 차단."""
        m = self._model
        rows = m.data_rows
        cols = self.columnCount()
        if cols <= 0:
            return
        fm = self.fontMetrics()
        adv = fm.horizontalAdvance
        pad = 14   # 셀 좌우 여백+그리드 근사
        if rows <= max_samples:
            sample_rows = range(rows)
        else:
            stride = max(1, rows // (max_samples - 100))
            sample_rows = list(range(100)) + list(range(100, rows, stride))
        hdr_fm = self.horizontalHeader().fontMetrics()
        self._applying_sizes = True
        try:
            for c in range(cols):
                w = hdr_fm.horizontalAdvance(get_column_letter(c + 1)) + 24
                for r in sample_rows:
                    text = m.display_text(r, c)
                    if not text:
                        continue
                    if "\n" in text:
                        tw = max(adv(line) for line in text.split("\n"))
                    else:
                        tw = adv(text)
                    if tw + pad > w:
                        w = tw + pad
                        if w >= MAX_AUTO_COL_WIDTH_PX:
                            break
                self.setColumnWidth(c, min(w, MAX_AUTO_COL_WIDTH_PX))
        finally:
            self._applying_sizes = False

    def set_key_col(self, col: int):
        self._key_col = col
        self._model.set_key_col(col)

    def set_excluded_cols(self, cols: set):
        """외부(MainWindow)에서 제외 열 집합을 갱신하고 헤더/셀을 다시 칠한다."""
        self._excluded_cols = set(cols)
        self._model.set_excluded_cols(self._excluded_cols)

    def _col_stage_kinds(self, col: int) -> tuple:
        """지정 열에 변경/스테이징 셀이 있는지 (has_changed, has_staged) 반환."""
        m = self._model
        has_changed = has_staged = False
        for r in range(m.data_rows):
            kind = m.cell_kind(r, col)
            if kind == "changed":
                has_changed = True
            elif kind == "staged":
                has_staged = True
            if has_changed and has_staged:
                break
        return has_changed, has_staged

    def _row_stage_kinds(self, row: int) -> tuple:
        """지정 행에 변경/스테이징 셀이 있는지 (has_changed, has_staged) 반환."""
        m = self._model
        has_changed = has_staged = False
        for c in range(m.data_cols):
            kind = m.cell_kind(row, c)
            if kind == "changed":
                has_changed = True
            elif kind == "staged":
                has_staged = True
            if has_changed and has_staged:
                break
        return has_changed, has_staged

    def _select_col(self, col: int):
        """해당 열 전체 셀을 선택 상태로 설정."""
        sm = self.selectionModel()
        rows = self.rowCount()
        cols = self.columnCount()
        if sm is None or rows == 0 or cols == 0 or not (0 <= col < cols):
            return
        model = self.model()
        sel = QItemSelection(model.index(0, col), model.index(rows - 1, col))
        sm.select(sel, QItemSelectionModel.ClearAndSelect)

    def _select_row(self, row: int):
        """해당 행 전체 셀을 선택 상태로 설정."""
        sm = self.selectionModel()
        rows = self.rowCount()
        cols = self.columnCount()
        if sm is None or rows == 0 or cols == 0 or not (0 <= row < rows):
            return
        model = self.model()
        sel = QItemSelection(model.index(row, 0), model.index(row, cols - 1))
        sm.select(sel, QItemSelectionModel.ClearAndSelect)

    def _selected_header_cols(self, anchor_col: int) -> list[int]:
        """우클릭 시 대상 열 집합 결정.
        - 우클릭한 열이 현재 헤더 다중 선택에 포함되어 있으면 그 선택 전체.
        - 아니면 우클릭한 단일 열만.
        """
        sel_model = self.selectionModel()
        cols: set[int] = set()
        if sel_model is not None:
            # selectedColumns()는 한 칼럼당 한 인덱스만 반환 — 헤더 클릭으로 전체열 선택 시 채워짐.
            for idx in sel_model.selectedColumns():
                cols.add(idx.column())
            # 셀 선택 모드에서 헤더를 Shift/Ctrl-클릭한 경우 selectedIndexes()도 보충.
            if not cols:
                for idx in sel_model.selectedIndexes():
                    cols.add(idx.column())
        if anchor_col in cols and len(cols) > 1:
            return sorted(cols)
        return [anchor_col]

    def _show_header_context_menu(self, pos):
        col = self.horizontalHeader().logicalIndexAt(pos)
        if col < 0:
            return

        target_cols = self._selected_header_cols(col)
        multi = len(target_cols) > 1
        col_letter = get_column_letter(col + 1)
        if multi:
            cols_label = ", ".join(get_column_letter(c + 1) for c in target_cols)
        else:
            cols_label = col_letter

        is_excluded = col in self._excluded_cols
        # 다중 선택 시 제외 토글은 "혼합 상태"를 다룬다: 하나라도 비제외면 일괄 제외, 전부 제외면 일괄 해제.
        if multi:
            any_not_excluded = any(c not in self._excluded_cols for c in target_cols)
            multi_action_exclude = any_not_excluded   # True → 제외 추가, False → 제외 해제
        else:
            multi_action_exclude = not is_excluded

        col_changed, col_staged = self._col_stage_kinds(col)
        # 제외된 열은 stage/unstage 액션을 표시하지 않는다 — 변경이 'same'으로 노출되므로 의미 없음.
        # 다중 선택일 땐 stage/key 액션은 단순화를 위해 노출하지 않는다 (제외 토글만 일괄 처리).
        has_changed = col_changed and not is_excluded and not multi
        has_staged  = col_staged and not is_excluded and not multi

        menu = QMenu(self)
        menu.setStyleSheet(MENU_QSS)

        # ── 병합 준비 항목 ──
        act_a2b = act_b2a = act_unstage = None
        if has_changed:
            act_a2b = menu.addAction(f"선택 열: A → B  병합 준비  [{col_letter}열]")
            act_b2a = menu.addAction(f"선택 열: B → A  병합 준비  [{col_letter}열]")
        if has_staged:
            if has_changed:
                menu.addSeparator()
            act_unstage = menu.addAction(f"선택 열: 병합 준비 취소  [{col_letter}열]")

        if has_changed or has_staged:
            menu.addSeparator()

        # ── 키 열 항목 (단일 선택일 때만) ──
        act_key_clear = act_reset = act_key_set = None
        if not multi:
            if col == self._key_col:
                act_key_cur = menu.addAction(f"[키 열]  {col_letter}열 — 현재 키 열")
                act_key_cur.setEnabled(False)
                act_key_clear = menu.addAction("🔓  키 열 해제 (ROW 순서 기반 비교)")
                act_reset = menu.addAction("↩  A열(기본값)으로 초기화") if col != 0 else None
            else:
                act_key_set = menu.addAction(f"키 열로 설정  [{col_letter}열]")

        # ── 변경 검사 제외 토글 ──
        menu.addSeparator()
        if multi_action_exclude:
            label = f"⊘  변경 검사에서 제외  [{cols_label}열]"
        else:
            label = f"↺  검사 제외 해제  [{cols_label}열]"
        act_excl = menu.addAction(label)

        act = menu.exec_(self.horizontalHeader().mapToGlobal(pos))
        if act is None:
            return

        if act_a2b is not None and act == act_a2b:
            self._select_col(col)
            self.stage_requested.emit("a_to_b")
        elif act_b2a is not None and act == act_b2a:
            self._select_col(col)
            self.stage_requested.emit("b_to_a")
        elif act_unstage is not None and act == act_unstage:
            self._select_col(col)
            self.unstage_requested.emit()
        elif act_key_clear is not None and act == act_key_clear:
            self.key_col_changed.emit(-1)
        elif act_reset is not None and act == act_reset:
            self.key_col_changed.emit(0)
        elif act_key_set is not None and act == act_key_set:
            self.key_col_changed.emit(col)
        elif act == act_excl:
            self.columns_exclude_set.emit(target_cols, multi_action_exclude)

    def _show_row_header_context_menu(self, pos):
        row = self.verticalHeader().logicalIndexAt(pos)
        if row < 0:
            return

        has_changed, has_staged = self._row_stage_kinds(row)

        if not has_changed and not has_staged:
            return

        menu = QMenu(self)
        menu.setStyleSheet(MENU_QSS)

        act_a2b = act_b2a = act_unstage = None
        if has_changed:
            act_a2b = menu.addAction("선택 행: A → B  병합 준비")
            act_b2a = menu.addAction("선택 행: B → A  병합 준비")
        if has_staged:
            if has_changed:
                menu.addSeparator()
            act_unstage = menu.addAction("선택 행: 병합 준비 취소")

        act = menu.exec_(self.verticalHeader().mapToGlobal(pos))
        if act is None:
            return

        if act_a2b is not None and act == act_a2b:
            self._select_row(row)
            self.stage_requested.emit("a_to_b")
        elif act_b2a is not None and act == act_b2a:
            self._select_row(row)
            self.stage_requested.emit("b_to_a")
        elif act_unstage is not None and act == act_unstage:
            self._select_row(row)
            self.unstage_requested.emit()

    def _show_context_menu(self, pos):
        sm = self.selectionModel()
        selected = sm.selectedIndexes() if sm is not None else []
        if not selected:
            return

        m = self._model
        kinds = {m.cell_kind(idx.row(), idx.column()) for idx in selected}
        has_changed = "changed" in kinds
        has_staged  = "staged" in kinds

        if not has_changed and not has_staged:
            return

        menu = QMenu(self)
        menu.setStyleSheet(MENU_QSS)

        act_a2b     = menu.addAction("선택 셀: A -> B  병합 준비") if has_changed else None
        act_b2a     = menu.addAction("선택 셀: B -> A  병합 준비") if has_changed else None
        act_unstage = None
        if has_staged:
            if has_changed:
                menu.addSeparator()
            act_unstage = menu.addAction("선택 셀: 병합 준비 취소")

        act = menu.exec_(self.viewport().mapToGlobal(pos))
        if act is None:
            return
        if act_a2b is not None and act == act_a2b:
            self.stage_requested.emit("a_to_b")
        elif act_b2a is not None and act == act_b2a:
            self.stage_requested.emit("b_to_a")
        elif act_unstage is not None and act == act_unstage:
            self.unstage_requested.emit()

    # ── 엑셀식 키보드 네비/선택/병합 단축키 ──────────────────────────────────
    def _is_empty_cell(self, r: int, c: int) -> bool:
        if r < 0 or r >= self.rowCount() or c < 0 or c >= self.columnCount():
            return True
        return self._model.display_text(r, c) == ""

    def _jump_target(self, r: int, c: int, dr: int, dc: int) -> tuple:
        """엑셀의 Ctrl+방향키 시맨틱으로 점프 대상 (row, col) 반환."""
        max_r = self.rowCount() - 1
        max_c = self.columnCount() - 1
        if max_r < 0 or max_c < 0:
            return (max(0, r), max(0, c))
        nr, nc = r + dr, c + dc
        # 범위 밖이면 그대로
        if nr < 0 or nr > max_r or nc < 0 or nc > max_c:
            return (max(0, min(r, max_r)), max(0, min(c, max_c)))
        cur_empty = self._is_empty_cell(r, c)
        next_empty = self._is_empty_cell(nr, nc)
        if cur_empty:
            # 다음 비어있지 않은 셀까지
            while 0 <= nr <= max_r and 0 <= nc <= max_c and self._is_empty_cell(nr, nc):
                nr += dr; nc += dc
            if nr < 0 or nr > max_r or nc < 0 or nc > max_c:
                # 못 찾으면 끝까지
                return (max(0, min(nr - dr, max_r)), max(0, min(nc - dc, max_c)))
            return (nr, nc)
        if next_empty:
            # 빈 구간 건너 다음 비어있지 않은 셀까지
            while 0 <= nr <= max_r and 0 <= nc <= max_c and self._is_empty_cell(nr, nc):
                nr += dr; nc += dc
            if nr < 0 or nr > max_r or nc < 0 or nc > max_c:
                return (max(0, min(nr - dr, max_r)), max(0, min(nc - dc, max_c)))
            return (nr, nc)
        # 연속 데이터의 마지막 비어있지 않은 셀까지
        while 0 <= nr + dr <= max_r and 0 <= nc + dc <= max_c \
                and not self._is_empty_cell(nr + dr, nc + dc):
            nr += dr; nc += dc
        return (nr, nc)

    def _select_range(self, r1: int, c1: int, r2: int, c2: int):
        """(r1,c1)~(r2,c2) 직사각형의 셀들을 모두 선택 상태로 설정 (기존 선택은 클리어).
        기존 구현은 아이템이 있는 셀만 선택됐으므로 데이터 영역으로 클램프한다."""
        sm = self.selectionModel()
        m = self._model
        if sm is None:
            return
        rs, re_ = sorted((r1, r2))
        cs, ce_ = sorted((c1, c2))
        rs, cs = max(0, rs), max(0, cs)
        re_ = min(re_, m.data_rows - 1)
        ce_ = min(ce_, m.data_cols - 1)
        if rs > re_ or cs > ce_:
            self.clearSelection()
            return
        sel = QItemSelection(m.index(rs, cs), m.index(re_, ce_))
        sm.select(sel, QItemSelectionModel.ClearAndSelect)

    def _has_changed_selection(self) -> bool:
        m = self._model
        return any(m.cell_kind(idx.row(), idx.column()) == "changed"
                   for idx in self.selectionModel().selectedIndexes())

    def _has_staged_selection(self) -> bool:
        m = self._model
        return any(m.cell_kind(idx.row(), idx.column()) == "staged"
                   for idx in self.selectionModel().selectedIndexes())

    # ── 헤더 다중 선택 지원 ──────────────────────────────────────────────────
    def _full_columns_selected(self) -> list[int]:
        """selectionModel().selectedColumns()는 한 칼럼이 모든 행에 걸쳐 선택된
        경우만 반환 → 비어있지 않으면 '열 헤더 선택' 상태."""
        sm = self.selectionModel()
        if sm is None:
            return []
        return sorted({idx.column() for idx in sm.selectedColumns()})

    def _full_rows_selected(self) -> list[int]:
        sm = self.selectionModel()
        if sm is None:
            return []
        return sorted({idx.row() for idx in sm.selectedRows()})

    def _select_column_range(self, c1: int, c2: int):
        """[c1..c2] 모든 열의 모든 행 셀을 선택."""
        sm = self.selectionModel()
        rows = self.rowCount()
        cols = self.columnCount()
        if sm is None or rows == 0 or cols == 0:
            return
        cs, ce_ = sorted((c1, c2))
        cs = max(0, cs); ce_ = min(cols - 1, ce_)
        if cs > ce_:
            return
        model = self.model()
        sel = QItemSelection(model.index(0, cs), model.index(rows - 1, ce_))
        sm.select(sel, QItemSelectionModel.ClearAndSelect)

    def _select_row_range(self, r1: int, r2: int):
        sm = self.selectionModel()
        rows = self.rowCount()
        cols = self.columnCount()
        if sm is None or rows == 0 or cols == 0:
            return
        rs, re_ = sorted((r1, r2))
        rs = max(0, rs); re_ = min(rows - 1, re_)
        if rs > re_:
            return
        model = self.model()
        sel = QItemSelection(model.index(rs, 0), model.index(re_, cols - 1))
        sm.select(sel, QItemSelectionModel.ClearAndSelect)

    def keyPressEvent(self, event):
        if self._populating:
            return super().keyPressEvent(event)
        key = event.key()
        mods = event.modifiers()
        ctrl = bool(mods & Qt.ControlModifier)
        shift = bool(mods & Qt.ShiftModifier)
        alt = bool(mods & Qt.AltModifier)

        cur_r, cur_c = self._current_cell()

        # ── 헤더 다중 선택 확장 (Shift / Ctrl+Shift + 방향키) ──
        # 열 전체가 선택된 상태에서 Shift+←/→ 는 열 단위 확장,
        # 행 전체가 선택된 상태에서 Shift+↑/↓ 는 행 단위 확장.
        if shift and not alt and key in (Qt.Key_Left, Qt.Key_Right, Qt.Key_Up, Qt.Key_Down):
            full_cols = self._full_columns_selected()
            full_rows = self._full_rows_selected()
            is_col_mode = bool(full_cols) and key in (Qt.Key_Left, Qt.Key_Right)
            is_row_mode = bool(full_rows) and key in (Qt.Key_Up, Qt.Key_Down)

            if is_col_mode and self.columnCount() > 0 and self.rowCount() > 0:
                # anchor 초기화: 단일 열만 선택돼있고 anchor 없음 → 그 열을 anchor로
                if self._header_anchor_col is None:
                    if len(full_cols) == 1:
                        self._header_anchor_col = full_cols[0]
                    else:
                        # 다중 열 이미 선택 — currentIndex와 가장 먼 끝을 anchor로
                        cur_col_idx = cur_c if cur_c >= 0 else full_cols[-1]
                        self._header_anchor_col = (
                            full_cols[0] if cur_col_idx == full_cols[-1] else full_cols[-1]
                        )
                # 현재 확장 끝점 = currentIndex 또는 anchor 반대편 끝
                if cur_c >= 0 and cur_c in full_cols:
                    cur_end = cur_c
                else:
                    cur_end = full_cols[-1] if self._header_anchor_col == full_cols[0] else full_cols[0]
                if ctrl:
                    target = 0 if key == Qt.Key_Left else self.columnCount() - 1
                else:
                    delta = -1 if key == Qt.Key_Left else 1
                    target = max(0, min(self.columnCount() - 1, cur_end + delta))
                self._select_column_range(self._header_anchor_col, target)
                self._set_current_cell(max(0, cur_r if cur_r >= 0 else 0), target)
                event.accept(); return

            if is_row_mode and self.columnCount() > 0 and self.rowCount() > 0:
                if self._header_anchor_row is None:
                    if len(full_rows) == 1:
                        self._header_anchor_row = full_rows[0]
                    else:
                        cur_row_idx = cur_r if cur_r >= 0 else full_rows[-1]
                        self._header_anchor_row = (
                            full_rows[0] if cur_row_idx == full_rows[-1] else full_rows[-1]
                        )
                if cur_r >= 0 and cur_r in full_rows:
                    cur_end = cur_r
                else:
                    cur_end = full_rows[-1] if self._header_anchor_row == full_rows[0] else full_rows[0]
                if ctrl:
                    target = 0 if key == Qt.Key_Up else self.rowCount() - 1
                else:
                    delta = -1 if key == Qt.Key_Up else 1
                    target = max(0, min(self.rowCount() - 1, cur_end + delta))
                self._select_row_range(self._header_anchor_row, target)
                self._set_current_cell(target, max(0, cur_c if cur_c >= 0 else 0))
                event.accept(); return

        # 헤더 anchor 라이프사이클: 헤더 모드 분기에 들어가지 않은 일반 키는 anchor 무효화
        # (단순 Shift 아닌 키, 혹은 헤더가 아닌 일반 셀 선택 상태일 때)
        if key in (Qt.Key_Left, Qt.Key_Right, Qt.Key_Up, Qt.Key_Down):
            if not shift:
                self._header_anchor_col = None
                self._header_anchor_row = None

        # ── 병합 단축키 (Alt+Left / Alt+Right / Alt+Backspace) ──
        if alt and not ctrl and not shift:
            if key == Qt.Key_Right:
                if self._has_changed_selection():
                    self.stage_requested.emit("a_to_b")
                event.accept(); return
            if key == Qt.Key_Left:
                if self._has_changed_selection():
                    self.stage_requested.emit("b_to_a")
                event.accept(); return
            if key in (Qt.Key_Backspace, Qt.Key_Delete):
                if self._has_staged_selection():
                    self.unstage_requested.emit()
                event.accept(); return

        # ── F2: 셀 편집란 포커스 요청 ──
        if key == Qt.Key_F2 and not ctrl and not shift and not alt:
            self.edit_focus_requested.emit()
            event.accept(); return

        # ── Delete: 단일 셀 값 비우기 ──
        if key == Qt.Key_Delete and not ctrl and not shift and not alt:
            sm = self.selectionModel()
            sel = sm.selectedIndexes() if sm is not None else []
            if len(sel) == 1:
                idx = sel[0]
                self.delete_cell_requested.emit(idx.row(), idx.column())
                event.accept(); return
            # 다중 선택 일괄 삭제는 사고 위험 — 무시 (기본 동작도 막음)
            event.accept(); return

        # ── Enter/Return: 엑셀처럼 아래 칸으로 이동 ──
        if key in (Qt.Key_Return, Qt.Key_Enter) and not ctrl and not alt:
            if cur_r >= 0 and cur_c >= 0 and cur_r + 1 < self.rowCount():
                self._set_current_cell(cur_r + 1, cur_c)
            event.accept(); return

        # ── Shift+Space: 행 전체, Ctrl+Space: 열 전체 ──
        if key == Qt.Key_Space and shift and not ctrl and not alt and cur_r >= 0:
            self._select_row(cur_r)
            event.accept(); return
        if key == Qt.Key_Space and ctrl and not shift and not alt and cur_c >= 0:
            self._select_col(cur_c)
            event.accept(); return

        # ── Ctrl(+Shift)+방향키: 데이터 경계 점프 (Excel 시맨틱) ──
        if ctrl and not alt and key in (Qt.Key_Left, Qt.Key_Right, Qt.Key_Up, Qt.Key_Down):
            if cur_r < 0 or cur_c < 0:
                return super().keyPressEvent(event)
            dr = -1 if key == Qt.Key_Up else (1 if key == Qt.Key_Down else 0)
            dc = -1 if key == Qt.Key_Left else (1 if key == Qt.Key_Right else 0)
            tr, tc = self._jump_target(cur_r, cur_c, dr, dc)
            if shift:
                # anchor = 현재 selection의 처음 시작점 추정 (currentIndex 기준)
                anchor = self.currentIndex()
                ar = anchor.row() if anchor.isValid() else cur_r
                ac = anchor.column() if anchor.isValid() else cur_c
                self._select_range(ar, ac, tr, tc)
                self._set_current_cell(tr, tc)
            else:
                self._set_current_cell(tr, tc)
            event.accept(); return

        # ── Ctrl+Home / Ctrl+End ──
        if ctrl and not alt and key == Qt.Key_Home:
            tr, tc = 0, 0
            if shift and cur_r >= 0 and cur_c >= 0:
                self._select_range(cur_r, cur_c, tr, tc)
            self._set_current_cell(tr, tc)
            event.accept(); return
        if ctrl and not alt and key == Qt.Key_End:
            tr, tc = max(0, self.rowCount() - 1), max(0, self.columnCount() - 1)
            if shift and cur_r >= 0 and cur_c >= 0:
                self._select_range(cur_r, cur_c, tr, tc)
            self._set_current_cell(tr, tc)
            event.accept(); return

        super().keyPressEvent(event)

    def populate(self, diff_matrix: list[list], which: str,
                 merged_set: set = None, staged: dict = None,
                 row_meta: list = None, excluded_cols: set = None):
        if not diff_matrix:
            self._safe_clear()
            return

        self._excluded_cols = set(excluded_cols) if excluded_cols else set()
        self._populating = True
        self._header_anchor_col = None
        self._header_anchor_row = None
        prev_updates = self.updatesEnabled()
        self.setUpdatesEnabled(False)
        try:
            # 모델 리셋 한 번 — 셀 아이템 생성 없음(O(1)). 리셋이 selectionModel
            # 시그널(선택 해제)을 발화시키므로 _populating 플래그 유지가 필수.
            self._model.set_diff_data(diff_matrix, row_meta, staged or {},
                                      merged_set or set(), self._excluded_cols)
            # 1) 샘플 기반 자동 너비(상한 클립 포함)
            # → 2) 사용자가 직접 조정한 열/행만 그 위에 덮어쓰기 (상한 무시).
            # 새로고침(_run_refresh)은 _user_col_widths/_user_row_heights를 미리
            # 비우므로 그 경로에서는 2)가 건너뛰어져 디폴트로 복귀한다.
            self._auto_size_columns()
            if self._user_col_widths or self._user_row_heights:
                self._apply_user_sizes()
        finally:
            self.setUpdatesEnabled(prev_updates)
            self._populating = False

    def populate_preview(self, data: list[list]):
        if not data:
            self._safe_clear()
            return
        self._populating = True
        self._header_anchor_col = None
        self._header_anchor_row = None
        prev_updates = self.updatesEnabled()
        self.setUpdatesEnabled(False)
        try:
            self._model.set_preview_data(data)
            self._auto_size_columns()
            if self._user_col_widths or self._user_row_heights:
                self._apply_user_sizes()
        finally:
            self.setUpdatesEnabled(prev_updates)
            self._populating = False

    def _safe_clear(self):
        self._populating = True
        try:
            self._model.clear()
        finally:
            self._populating = False
        self._header_anchor_col = None
        self._header_anchor_row = None

    def get_selected_cells(self) -> set:
        sm = self.selectionModel()
        if sm is None:
            return set()
        return {(idx.row(), idx.column()) for idx in sm.selectedIndexes()}

    def mirror_selection(self, cells: set):
        # 셀 집합을 row별로 묶고 연속 column 구간을 QItemSelectionRange로 만들어
        # 한 번의 select() 호출로 일괄 적용 — 헤더 클릭처럼 N=수천 셀일 때 결정적.
        self._populating = True
        prev_updates = self.updatesEnabled()
        self.setUpdatesEnabled(False)
        try:
            sm = self.selectionModel()
            if sm is None:
                return
            row_max = self.rowCount() - 1
            col_max = self.columnCount() - 1
            if row_max < 0 or col_max < 0:
                sm.clearSelection()
                return
            by_row: dict[int, list[int]] = {}
            for (r, c) in cells:
                if 0 <= r <= row_max and 0 <= c <= col_max:
                    by_row.setdefault(r, []).append(c)
            sel = QItemSelection()
            model = self.model()
            for r, cs in by_row.items():
                cs.sort()
                start = prev = cs[0]
                for c in cs[1:]:
                    if c == prev + 1:
                        prev = c
                        continue
                    sel.append(QItemSelectionRange(model.index(r, start), model.index(r, prev)))
                    start = prev = c
                sel.append(QItemSelectionRange(model.index(r, start), model.index(r, prev)))
            sm.select(sel, QItemSelectionModel.ClearAndSelect)
        finally:
            self.setUpdatesEnabled(prev_updates)
            self._populating = False




def _extract_supported_path(mime_data) -> str:
    if mime_data.hasUrls():
        for url in mime_data.urls():
            path = url.toLocalFile()
            if os.path.splitext(path)[1].lower() in _SUPPORTED_EXTS:
                return path
    return ""


class DropLineEdit(QLineEdit):
    file_dropped = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if _extract_supported_path(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if _extract_supported_path(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        path = _extract_supported_path(event.mimeData())
        if path:
            self.file_dropped.emit(path)
            event.acceptProposedAction()


class CellEditWidget(QPlainTextEdit):
    """셀 값 편집란 — Enter: 적용, Alt+Enter: 줄바꿈 입력.
    항상 2줄 고정 높이. 3줄 이상은 세로 스크롤.
    """
    apply_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        line_h = self.fontMetrics().lineSpacing()
        # 2줄이 잘리지 않게 여유 패딩 포함
        self.setFixedHeight(line_h * 2 + 12)
        self.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._inserting_newline = False
        self._block_auto_scroll = False   # ensureCursorVisible 차단 플래그

    def ensureCursorVisible(self):
        # Alt+Enter 삽입 중에는 Qt 자동 스크롤을 차단
        if self._block_auto_scroll:
            return
        super().ensureCursorVisible()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter) and not (event.modifiers() & Qt.AltModifier):
            self.apply_requested.emit()
        elif event.key() in (Qt.Key_Return, Qt.Key_Enter) and (event.modifiers() & Qt.AltModifier):
            vsb = self.verticalScrollBar()
            cur_scroll = vsb.value()
            line_h = self.fontMetrics().lineSpacing()
            # 문서 전체에서 텍스트가 있는 줄이 2줄 이상일 때 스크롤 대상
            doc = self.document()
            non_empty_lines = sum(
                1 for i in range(doc.blockCount())
                if doc.findBlockByNumber(i).text().strip()
            )
            has_two_or_more_lines = non_empty_lines >= 2
            # 이미 최하단에 도달한 경우 스크롤 생략
            at_bottom = cur_scroll >= vsb.maximum()
            should_scroll = has_two_or_more_lines and not at_bottom
            # 자동 스크롤 차단 후 삽입, 직접 스크롤 값 설정
            self._block_auto_scroll = True
            self.textCursor().insertText("\n")
            self._block_auto_scroll = False
            if should_scroll:
                vsb.setValue(cur_scroll + line_h)
        else:
            super().keyPressEvent(event)
            # 일반 타이핑 시 스크롤 위치 유지 (현재 블록이 2번째 이내면 맨 위 고정)
            self._clamp_scroll_if_not_last()



    def _clamp_scroll_if_not_last(self):
        """커서가 마지막 블록이 아니면 스크롤을 상단으로 고정."""
        cursor = self.textCursor()
        doc = self.document()
        if cursor.blockNumber() < doc.blockCount() - 1:
            self.verticalScrollBar().setValue(0)

    def text(self):
        return self.toPlainText()

    def setText(self, val: str):
        self.setPlainText(val if val is not None else "")
        # 텍스트 설정 후 항상 맨 위부터 표시
        self.verticalScrollBar().setValue(0)

    def clear(self):
        super().clear()

