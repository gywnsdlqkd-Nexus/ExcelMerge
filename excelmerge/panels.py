"""A/B 파일 패널 (excel_diff_merge.py에서 분리)."""
import os

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QShortcut, QStyle, QSplitter,
)
from PyQt5.QtCore import Qt, pyqtSignal, QSize
from PyQt5.QtGui import QKeySequence

from .theme import CELL_DIFF_HL, DROP_HIGHLIGHT_QSS, ui_font
from .widgets import (
    CellEditWidget, DropLineEdit, ExcelTableView,
    _extract_supported_path, _extract_folder_path,
)


class FilePanel(QWidget):
    file_loaded = pyqtSignal(str)
    folder_loaded = pyqtSignal(str)   # 폴더가 드롭/선택됨 → 상위(탭)가 폴더 모드로 전환

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
        self._drop_hint = QLabel("  엑셀/JSON/uasset 파일 또는 폴더를 여기에 끌어다 놓으세요")
        self._drop_hint.setStyleSheet("color: #888; font-size: 9pt;")
        self._drop_hint.setFont(ui_font(9))
        header.addWidget(self._drop_hint)
        header.addStretch()
        layout.addLayout(header)

        file_row = QHBoxLayout()
        self.path_edit = DropLineEdit()
        self.path_edit.setPlaceholderText("파일/폴더 경로 입력(Enter) 또는 드래그·찾아보기...")
        # 편집 가능 — 파일/폴더 경로를 직접 입력·붙여넣기(Enter로 로드). 표준 편집
        # 컨텍스트 메뉴(붙여넣기)를 쓰도록 CustomContextMenu는 지정하지 않는다.
        self.path_edit.file_dropped.connect(self._on_file_dropped)
        self.path_edit.returnPressed.connect(self._on_path_entered)
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

        # 셀 값 표시란 — 선택 셀의 값 확인용 (읽기전용, 직접 수정 불가)
        self.cell_edit = CellEditWidget()
        self.cell_edit.setPlaceholderText("셀을 선택하면 값이 여기에 표시됩니다")
        self.cell_edit.setFont(ui_font(9))
        self.cell_edit.setEnabled(False)
        self._selected_cell: tuple | None = None   # (row, col) 현재 선택 셀
        self._row_meta: list = []   # [(orig_a_row, orig_b_row), ...]
        self._staged_display: dict[tuple, str] = {}   # (r,c) → 병합 준비 셀의 셀값란 표시 문자열

        self.table = ExcelTableView(side)
        # QTableView에는 itemSelectionChanged가 없다 — selectionModel은 ctor에서
        # 1회 생성 후 교체되지 않으므로 여기서 connect해도 안전하다.
        self.table.selectionModel().selectionChanged.connect(
            lambda *_: self._on_table_selection_changed())

        # 셀값란/테이블을 수직 스플리터로 — 핸들을 드래그해 셀값란 높이를 조절.
        self.v_split = QSplitter(Qt.Vertical)
        self.v_split.setChildrenCollapsible(False)
        self.v_split.addWidget(self.cell_edit)
        self.v_split.addWidget(self.table)
        self.v_split.setStretchFactor(0, 0)
        self.v_split.setStretchFactor(1, 1)
        self.v_split.setSizes([self.cell_edit.sizeHint().height(), 100000])
        layout.addWidget(self.v_split)

        copy_sc = QShortcut(QKeySequence("Ctrl+C"), self)
        copy_sc.activated.connect(self._on_copy_shortcut)

    def _on_table_selection_changed(self):
        if self.table._populating:
            return
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
                # 값 표시 + A/B 차이 문자 강조 (modified 셀에서만 range 존재)
                ranges = model.diff_char_ranges(r, c)
                if ranges:
                    self.cell_edit.set_highlighted(display, ranges, CELL_DIFF_HL)
                else:
                    self.cell_edit.setText(display)
            self.cell_edit.setEnabled(True)
        else:
            self._selected_cell = None
            self.cell_edit.clear()
            self.cell_edit.setEnabled(False)

    def _sync_cell_edit(self):
        """mirror_selection 후 cell_edit 값을 현재 선택 셀에 맞게 갱신 (포커스 이동 없음)."""
        self._refresh_cell_edit_from_selection()

    @staticmethod
    def _accepts(mime) -> bool:
        return bool(_extract_supported_path(mime) or _extract_folder_path(mime))

    def dragEnterEvent(self, event):
        if self._accepts(event.mimeData()):
            event.acceptProposedAction()
            self._set_drop_highlight(True)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self._accepts(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, _event):
        self._set_drop_highlight(False)

    def dropEvent(self, event):
        self._set_drop_highlight(False)
        mime = event.mimeData()
        fpath = _extract_supported_path(mime)     # 파일 우선
        if fpath:
            self._on_file_dropped(fpath)
            event.acceptProposedAction()
            return
        dpath = _extract_folder_path(mime)
        if dpath:
            # 폴더 → 상위 탭이 폴더 비교로 전환하도록 신호만 방출
            self.folder_loaded.emit(dpath)
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

    # 폴더 선택용 센티넬 파일명 — 네이티브 파일 대화상자에서 이 이름으로 '열기'하면
    # 해당 폴더를 폴더 비교로 처리한다(Windows엔 파일+폴더 동시 선택 대화상자가 없어 우회).
    _FOLDER_SENTINEL = "폴더 선택."

    def _on_path_entered(self):
        """경로칸에 직접 입력한 경로를 파일/폴더로 판별해 로드 (Enter)."""
        p = self.path_edit.text().strip().strip('"')
        if not p:
            return
        if os.path.isdir(p):
            self.folder_loaded.emit(p)
        elif os.path.isfile(p):
            self._on_file_dropped(p)

    def _browse(self):
        """찾아보기 — 네이티브 윈도우 파일 대화상자.
        파일을 고르면 셀 비교. 폴더는 그 폴더로 들어가 파일명 '폴더 선택.' 그대로
        '열기'하면 그 폴더가 폴더 비교로 처리된다(파일 존재 검사 없는 AnyFile 모드)."""
        current = self.path_edit.text().strip()
        init_dir = (os.path.dirname(current) if current and os.path.exists(current)
                    else os.path.expanduser("~"))
        dlg = QFileDialog(self, "비교할 파일 선택 — 폴더는 해당 폴더로 이동 후 '열기'")
        dlg.setAcceptMode(QFileDialog.AcceptOpen)
        dlg.setFileMode(QFileDialog.AnyFile)   # 비존재 이름 허용(센티넬 반환용)
        dlg.setNameFilter(
            "Excel/JSON/uasset (*.xlsx *.xls *.xlsm *.json *.uasset);;모든 파일 (*)")
        dlg.setDirectory(init_dir)
        dlg.selectFile(self._FOLDER_SENTINEL)   # 파일명 칸 기본값 = 폴더 선택 센티넬
        if not dlg.exec_():
            return
        sel = dlg.selectedFiles()
        path = sel[0] if sel else ""
        if not path:
            return
        # 센티넬 문자열 매칭에 의존하지 않는다 — Windows가 파일명 끝 마침표를 제거해
        # '폴더 선택.' → '폴더 선택'이 되므로. 경로 존재 여부로 견고하게 판정:
        if os.path.isfile(path):          # 실제 파일 → 셀 비교
            self._on_file_dropped(path)
        elif os.path.isdir(path):         # 실제 폴더 반환 → 폴더 비교
            self.folder_loaded.emit(path)
        else:                             # 비존재 이름(센티넬 등) → 그 상위 폴더 선택으로 간주
            folder = os.path.dirname(path)
            if os.path.isdir(folder):
                self.folder_loaded.emit(folder)

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

