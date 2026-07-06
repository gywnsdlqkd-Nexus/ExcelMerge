"""A/B 파일 패널 (excel_diff_merge.py에서 분리)."""
import os

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QFrame, QMenu, QShortcut, QStyle,
)
from PyQt5.QtCore import Qt, pyqtSignal, QSize
from PyQt5.QtGui import QKeySequence

from .theme import DROP_HIGHLIGHT_QSS, MENU_QSS, ui_font
from .widgets import (
    CellEditWidget, DropLineEdit, ExcelTableView, _extract_supported_path,
)


class FilePanel(QWidget):
    file_loaded = pyqtSignal(str)
    cell_value_edited = pyqtSignal(int, int, str)   # row, col, new_value

    def __init__(self, label: str, side: str, parent=None):
        super().__init__(parent)
        self.side = side
        self.setAcceptDrops(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        header = QHBoxLayout()
        title = QLabel(label)
        title.setFont(ui_font(10, bold=True))
        header.addWidget(title)
        self._drop_hint = QLabel("  엑셀/JSON/uasset 파일을 여기에 끌어다 놓으세요")
        self._drop_hint.setStyleSheet("color: #888; font-size: 9pt;")
        self._drop_hint.setFont(ui_font(9))
        header.addWidget(self._drop_hint)
        header.addStretch()
        layout.addLayout(header)

        file_row = QHBoxLayout()
        self.path_edit = DropLineEdit()
        self.path_edit.setPlaceholderText("엑셀/JSON/uasset 파일을 드래그하거나 찾아보기 버튼을 클릭하세요...")
        self.path_edit.setReadOnly(True)
        self.path_edit.setFocusPolicy(Qt.NoFocus)
        self.path_edit.setContextMenuPolicy(Qt.CustomContextMenu)
        self.path_edit.customContextMenuRequested.connect(self._on_path_context_menu)
        self.path_edit.file_dropped.connect(self._on_file_dropped)
        browse_btn = QPushButton()
        browse_btn.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
        browse_btn.setFixedSize(32, 32)
        browse_btn.setIconSize(QSize(18, 18))
        browse_btn.setToolTip("찾아보기")
        browse_btn.clicked.connect(self._browse)

        self.save_btn = QPushButton()
        self.save_btn.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.save_btn.setFixedSize(32, 32)
        self.save_btn.setIconSize(QSize(18, 18))
        self.save_btn.setToolTip("파일 저장")
        self.save_btn.setEnabled(False)
        self.save_btn.setObjectName("save_btn")

        file_row.addWidget(self.path_edit)
        file_row.addWidget(browse_btn)
        file_row.addWidget(self.save_btn)
        layout.addLayout(file_row)

        # 셀 값 편집 행 — 라벨은 표시하지 않고 입력란만 노출
        edit_row = QHBoxLayout()
        self.cell_edit = CellEditWidget()
        self.cell_edit.setPlaceholderText("셀 선택 후 F2로 편집 (Enter 적용 / Alt+Enter 줄바꿈)")
        self.cell_edit.setFont(ui_font(9))
        self.cell_edit.setEnabled(False)
        self.cell_edit.apply_requested.connect(self._apply_cell_edit)
        edit_row.addWidget(self.cell_edit)
        layout.addLayout(edit_row)
        self._selected_cell: tuple | None = None   # (row, col) 현재 선택 셀
        self._formula_data: list[list] = []
        self._row_meta: list = []   # [(orig_a_row, orig_b_row), ...]
        self._staged_display: dict[tuple, str] = {}   # (r,c) → 병합 준비 셀의 셀값란 표시 문자열
        self._edited_values: dict[tuple, str] = {}   # (r,c) → 직접 편집된 값 (셀값란 표시 우선)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)

        self.table = ExcelTableView(side)
        # QTableView에는 itemSelectionChanged가 없다 — selectionModel은 ctor에서
        # 1회 생성 후 교체되지 않으므로 여기서 connect해도 안전하다.
        self.table.selectionModel().selectionChanged.connect(
            lambda *_: self._on_table_selection_changed())
        self.table.edit_focus_requested.connect(self._focus_cell_edit)
        self.table.delete_cell_requested.connect(self._on_delete_cell_requested)
        layout.addWidget(self.table)

        copy_sc = QShortcut(QKeySequence("Ctrl+C"), self)
        copy_sc.activated.connect(self._on_copy_shortcut)

    def _get_formula(self, display_r: int, c: int) -> str:
        """표시 행(display_r) → 원본 파일 행 인덱스로 변환 후 수식 문자열만 반환.
        수식이 아닌 일반 값 셀은 빈 문자열 반환 (편집값 우선 표시를 위해)."""
        side_idx = 0 if self.side == "a" else 1
        try:
            orig = self._row_meta[display_r][side_idx]
        except (IndexError, TypeError):
            orig = display_r
        if orig is None:
            return ""
        try:
            v = self._formula_data[orig][c]
            return v if v.startswith("=") else ""
        except (IndexError, TypeError):
            return ""

    def _on_table_selection_changed(self):
        if self.table._populating:
            return
        # 다른 셀로 이동 시 편집 중인 값 자동 적용
        # cell_edit 값과 비교할 때 수식 표시 중일 수 있으므로 원래 표시값도 함께 확인
        if self._selected_cell is not None:
            current_text = self.cell_edit.text()
            pr, pc = self._selected_cell
            original_text = self.table.model().display_text(pr, pc)
            original_formula = self._get_formula(pr, pc)
            staged_override = self._staged_display.get((pr, pc))
            if (current_text != original_text
                    and current_text != original_formula
                    and (staged_override is None or current_text != staged_override)):
                self._apply_cell_edit()

        self._refresh_cell_edit_from_selection()

    def _refresh_cell_edit_from_selection(self):
        """현재 선택 셀에 맞게 cell_edit 값/활성 상태를 갱신."""
        model = self.table.model()
        cell = self.table._single_selected_cell()   # 대량 선택에서도 O(range 수)
        if cell is not None:
            r, c = cell
            self._selected_cell = (r, c)
            display = model.display_text(r, c)
            if model.cell_kind(r, c) in ("staged", "merged"):
                override = self._staged_display.get((r, c))
                self.cell_edit.setText(override if override is not None else display)
            else:
                # 직접 편집된 값이 있으면 최우선 표시
                edited_val = self._edited_values.get((r, c))
                if edited_val is not None:
                    self.cell_edit.setText(edited_val)
                else:
                    formula = self._get_formula(r, c)
                    self.cell_edit.setText(formula if formula else display)
            self.cell_edit.setEnabled(True)
        else:
            self._selected_cell = None
            self.cell_edit.clear()
            self.cell_edit.setEnabled(False)

    def _apply_cell_edit(self):
        if self._selected_cell is None:
            return
        r, c = self._selected_cell
        new_val = self.cell_edit.text()
        self.cell_value_edited.emit(r, c, new_val)
        self.cell_edit.clearFocus()

    def _sync_cell_edit(self):
        """mirror_selection 후 cell_edit 값을 현재 선택 셀에 맞게 갱신 (포커스 이동 없음)."""
        self._refresh_cell_edit_from_selection()

    def dragEnterEvent(self, event):
        if _extract_supported_path(event.mimeData()):
            event.acceptProposedAction()
            self._set_drop_highlight(True)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if _extract_supported_path(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, _event):
        self._set_drop_highlight(False)

    def dropEvent(self, event):
        path = _extract_supported_path(event.mimeData())
        self._set_drop_highlight(False)
        if path:
            self._on_file_dropped(path)
            event.acceptProposedAction()

    def _set_drop_highlight(self, active: bool):
        if active:
            self.setStyleSheet(DROP_HIGHLIGHT_QSS)
        else:
            self.setStyleSheet("")

    def _on_file_dropped(self, path: str):
        self.path_edit.setText(path)
        self.file_loaded.emit(path)

    def _copy_path(self):
        path = self.path_edit.text().strip()
        if path:
            QApplication.clipboard().setText(path)

    def _on_copy_shortcut(self):
        """Ctrl+C — 테이블에 포커스 시 선택 영역 TSV 복사, 그 외엔 경로 복사."""
        focused = QApplication.focusWidget()
        is_table_focus = False
        w = focused
        while w is not None:
            if w is self.table:
                is_table_focus = True
                break
            w = w.parentWidget()
        if is_table_focus:
            self._copy_selection_as_tsv()
        else:
            self._copy_path()

    def _copy_selection_as_tsv(self):
        """선택 셀들을 bounding box 기준 TSV로 클립보드에 복사."""
        model = self.table.model()
        sm = self.table.selectionModel()
        indexes = sm.selectedIndexes() if sm is not None else []
        if not indexes:
            return
        rows = [idx.row() for idx in indexes]
        cols = [idx.column() for idx in indexes]
        r1, r2 = min(rows), max(rows)
        c1, c2 = min(cols), max(cols)
        sel_set = {(idx.row(), idx.column()) for idx in indexes}
        lines = []
        for r in range(r1, r2 + 1):
            cells = []
            for c in range(c1, c2 + 1):
                cells.append(model.display_text(r, c) if (r, c) in sel_set else "")
            lines.append("\t".join(cells))
        QApplication.clipboard().setText("\r\n".join(lines))

    def _focus_cell_edit(self):
        if self.cell_edit.isEnabled():
            self.cell_edit.setFocus()
            # 전체 선택해서 바로 덮어쓰기 가능하게
            cursor = self.cell_edit.textCursor()
            cursor.select(cursor.Document)
            self.cell_edit.setTextCursor(cursor)

    def _on_delete_cell_requested(self, r: int, c: int):
        """Delete 키로 단일 셀 값 비우기 — 기존 편집 흐름(cell_value_edited) 재사용."""
        self.cell_value_edited.emit(r, c, "")

    def _on_path_context_menu(self, pos):
        menu = QMenu(self)
        menu.setStyleSheet(MENU_QSS)
        act = menu.addAction("경로 복사  (Ctrl+C)")
        act.setEnabled(bool(self.path_edit.text().strip()))
        if menu.exec_(self.path_edit.mapToGlobal(pos)) == act:
            self._copy_path()

    def _browse(self):
        current = self.path_edit.text().strip()
        init_dir = os.path.dirname(current) if current and os.path.exists(current) else os.path.expanduser("~")
        path, _ = QFileDialog.getOpenFileName(
            self, "비교할 파일 선택", init_dir,
            "Supported (*.xlsx *.xls *.xlsm *.json *.uasset);;"
            "Excel (*.xlsx *.xls *.xlsm);;"
            "JSON (*.json);;"
            "Unreal Asset (*.uasset);;"
            "All Files (*)",
        )
        if path:
            self._on_file_dropped(path)

    def get_path(self) -> str:
        return self.path_edit.text().strip()

    def set_path(self, path: str):
        self.path_edit.setText(path)

    def populate(self, diff_matrix: list[list], merged_set: set = None,
                 staged: dict = None, row_meta: list = None,
                 excluded_cols: set = None):
        self.table.populate(diff_matrix, self.side, merged_set, staged, row_meta,
                            excluded_cols)

    def preview(self, data: list[list]):
        self.table.populate_preview(data)

