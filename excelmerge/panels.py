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
    AX_COL, AX_ROW, CellEditWidget, DropLineEdit, ExcelTableView,
    _extract_supported_path, _extract_folder_path,
)


def _tsv_cell(text: str) -> str:
    """엑셀이 읽는 TSV 한 칸 — 탭·줄바꿈·따옴표가 들어 있으면 큰따옴표로 감싼다.

    감싸지 않으면 값 안의 탭이 칸을, 줄바꿈이 행을 갈라 붙여넣기가 통째로 어긋난다
    (이 도구가 다루는 텍스트 테이블에는 줄바꿈이 든 셀이 흔하다).
    """
    if any(ch in text for ch in ("\t", "\r", "\n", '"')):
        return '"' + text.replace('"', '""') + '"'
    return text


class FilePanel(QWidget):
    file_loaded = pyqtSignal(str)
    status_message = pyqtSignal(str)   # 복사 결과 등 → 상위(DiffView) 상태바
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
        self._drop_hint = QLabel("  엑셀/CSV/JSON/uasset 파일 또는 폴더를 여기에 끌어다 놓으세요")
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
        browse_btn.setFocusPolicy(Qt.NoFocus)   # 눌러도 키보드 포커스는 표에 남긴다
        browse_btn.clicked.connect(self._browse)

        self.save_btn = QPushButton()
        self.save_btn.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.save_btn.setFixedSize(32, 32)
        self.save_btn.setIconSize(QSize(18, 18))
        self.save_btn.setToolTip("파일 저장")
        self.save_btn.setEnabled(False)
        self.save_btn.setObjectName("save_btn")
        self.save_btn.setFocusPolicy(Qt.NoFocus)

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

        # 컨텍스트를 반드시 위젯 범위로 — 기본값(WindowShortcut)이면 A/B 두 패널이
        # 같은 창에 똑같은 Ctrl+C 를 등록한 꼴이라 Qt 가 '모호함'으로 보고
        # activated 를 아예 쏘지 않는다(= 복사가 통째로 죽는다. 실측으로 확인).
        copy_sc = QShortcut(QKeySequence("Ctrl+C"), self)
        copy_sc.setContext(Qt.WidgetWithChildrenShortcut)
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
        """Ctrl+C — 표에 포커스면 선택 영역을 TSV로, 그 밖에선 경로를 복사.

        틀 고정 오버레이(_FrozenView)는 본체 표의 자식이라 부모를 타고 올라가면
        여기서 '표 포커스'로 잡힌다 — 밴드 안에서 눌러도 같은 선택이 복사된다.

        경로칸·셀값란은 여기까지 오지 않는다: 편집 위젯은 자기가 처리하는 키에
        ShortcutOverride 를 돌려줘 앱 단축키를 끄고 **자기 기본 복사**(고른 글자만)를
        한다. 그게 윈도우 표준 동작이라 그대로 둔다.
        """
        focused = QApplication.focusWidget()
        w = focused
        while w is not None:
            if w is self.table:
                self._copy_selection_as_tsv()
                return
            if w is self:
                break
            w = w.parentWidget()
        self._copy_path()

    def _copy_selection_as_tsv(self):
        """선택 영역을 엑셀이 그대로 붙여넣을 수 있는 TSV 로 클립보드에 복사한다.

        - **화면에서 볼 수 없는 행/열은 뺀다.** '변경점만 보기'로 숨은 행이 그렇다
          (Ctrl+A 는 숨은 행까지 잡는다 — 엑셀도 필터가 걸리면 보이는 행만 복사한다).
          틀 고정 밴드는 본체에서 숨김이지만 오버레이에 보이므로 포함한다(_navigable).
        - 선택이 닿은 행·열만 추려 붙인다. 사각형 선택이면 그대로고, 떨어진 두 열을
          고른 경우엔 그 둘이 나란히 나온다(엑셀은 이 경우 복사를 거부한다).
        - 사각형 안에서 고르지 않은 칸은 빈칸으로 둔다.
        """
        tbl = self.table
        model = tbl.model()
        sm = tbl.selectionModel()
        indexes = sm.selectedIndexes() if sm is not None else []
        if not indexes:
            self.status_message.emit("복사할 선택 영역이 없습니다.")
            return
        picked = {(idx.row(), idx.column()) for idx in indexes}
        rows = sorted({r for r, _c in picked if tbl._navigable(AX_ROW, r)})
        cols = sorted({c for _r, c in picked if tbl._navigable(AX_COL, c)})
        if not rows or not cols:
            self.status_message.emit("선택한 행/열이 화면에 없어 복사하지 않았습니다.")
            return
        lines = ["\t".join(_tsv_cell(model.display_text(r, c)) if (r, c) in picked else ""
                           for c in cols)
                 for r in rows]
        QApplication.clipboard().setText("\r\n".join(lines))
        self.status_message.emit(f"{len(rows)}행 × {len(cols)}열 복사됨")

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
            "Excel/CSV/JSON/uasset "
            "(*.xlsx *.xls *.xlsm *.csv *.tsv *.json *.uasset);;모든 파일 (*)")
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

