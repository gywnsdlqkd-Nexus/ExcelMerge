"""폴더 비교 뷰 — Beyond Compare 스타일 좌/우 파일 목록 탭 위젯.

좌(A)·우(B) 두 개의 트리를 **동일한 노드 구조**로 빌드해 행을 정렬한다.
한쪽에만 있는 파일은 반대쪽 같은 행에 빈 placeholder로 채워 항상 좌우가 맞는다.
파일 상태(동일/변경/A만/B만)는 배경색으로 구분하고, 파일을 더블클릭(또는 Enter)하면
open_file_pair 시그널로 그 파일 쌍의 diff 탭 열기를 요청한다.

스크롤·확장/축소·선택은 A↔B 트리가 연동된다(구조가 동일하므로 인덱스로 미러).
"""
import os

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit,
    QSplitter, QTreeWidget, QTreeWidgetItem, QFileDialog, QAbstractItemView,
    QFrame, QStyle, QShortcut, QStyledItemDelegate, QStyleOptionViewItem,
    QMenu, QMessageBox,
)
from PyQt5.QtCore import Qt, pyqtSignal, QSize, QItemSelectionModel
from PyQt5.QtGui import QBrush, QKeySequence

from .folder_compare import (
    summarize, SAME, MODIFIED, ONLY_A, ONLY_B,
)
from .workers import FolderScanWorker, FolderMergeWorker
from .constants import DIR_A2B, DIR_B2A
from .theme import (
    APP_QSS, ui_font, FOLDER_STATUS_COLORS, DIFF_COLORS, MENU_QSS,
    DROP_HIGHLIGHT_QSS, force_active_highlight,
)
from .widgets import _extract_folder_path, draw_diagonal_hatch
from .compare_toolbar import build_find_box, add_legend

_ROLE_PAIR = Qt.UserRole         # 이 항목의 (item_a, item_b) 쌍 인덱스 → self._pairs
_ROLE_ENTRY = Qt.UserRole + 1    # 파일이면 self._entries 인덱스, 폴더/빈칸이면 -1
_ROLE_PLACEHOLDER = Qt.UserRole + 2   # 반대쪽에만 있는 파일의 빈 칸(대각선 해치) 표시


class _HatchDelegate(QStyledItemDelegate):
    """placeholder(반대쪽에만 있는 파일의 빈 칸) 항목에 대각선 해치를 그린다."""

    def paint(self, painter, option, index):
        if index.data(_ROLE_PLACEHOLDER):
            opt = QStyleOptionViewItem(option)
            self.initStyleOption(opt, index)
            widget = opt.widget
            style = widget.style() if widget is not None else QApplication.style()
            opt.text = ""
            style.drawControl(QStyle.CE_ItemViewItem, opt, painter, widget)
            draw_diagonal_hatch(painter, opt.rect)
            return
        super().paint(painter, option, index)


class _FolderDropLineEdit(QLineEdit):
    """폴더를 드롭하면 folder_dropped(경로)를 방출하는 경로 입력란(편집 가능)."""
    folder_dropped = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setPlaceholderText("폴더 경로 입력(Enter) 또는 드롭·찾아보기...")

    def dragEnterEvent(self, event):
        if _extract_folder_path(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if _extract_folder_path(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        p = _extract_folder_path(event.mimeData())
        if p:
            self.folder_dropped.emit(p)
            event.acceptProposedAction()


class _FolderDropPanel(QWidget):
    """폴더 비교 한 쪽(A/B) 패널 컨테이너 — 패널 어디에 폴더를 떨궈도 등록되게 한다.
    (파일 비교의 FilePanel처럼 패널 전체가 드롭 대상. 이전엔 좁은 경로칸만 받아, 넓은
    트리 영역에 드롭하면 무반응이라 '반대 패널 등록 불가'처럼 보였다.)"""
    folder_dropped = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if _extract_folder_path(event.mimeData()):
            event.acceptProposedAction()
            self.setStyleSheet(DROP_HIGHLIGHT_QSS)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if _extract_folder_path(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, _event):
        self.setStyleSheet("")

    def dropEvent(self, event):
        self.setStyleSheet("")
        p = _extract_folder_path(event.mimeData())
        if p:
            self.folder_dropped.emit(p)
            event.acceptProposedAction()


class _FolderTree(QTreeWidget):
    """Enter 시 '선택된 모든 파일'을 열도록 신호를 낸다.
    (기본 itemActivated는 현재 항목 1개만 열어, 다중 선택 Enter가 마지막 1개만 열리던 문제 해결)"""
    enter_pressed = pyqtSignal()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.enter_pressed.emit()
            event.accept()
            return
        super().keyPressEvent(event)


class _Node:
    """폴더 트리 노드 — 파일이면 entry_index로 FolderEntry를 가리킨다."""
    __slots__ = ("name", "is_file", "entry_index", "children")

    def __init__(self, name: str, is_file: bool = False, entry_index: int = -1):
        self.name = name
        self.is_file = is_file
        self.entry_index = entry_index
        self.children: dict = {}   # name → _Node (삽입 순서 = 정렬 순서)


def _build_node_tree(entries: list) -> _Node:
    """정렬된 FolderEntry 리스트를 폴더/파일 트리로 그룹핑."""
    root = _Node("")
    for idx, e in enumerate(entries):
        parts = e.rel_path.split("/")
        node = root
        for i, part in enumerate(parts):
            is_file = (i == len(parts) - 1)
            if part not in node.children:
                node.children[part] = _Node(
                    part, is_file=is_file, entry_index=idx if is_file else -1)
            node = node.children[part]
    return root


class FolderCompareView(QWidget):
    # (path_a, path_b, title) — 파일 쌍 diff 탭 열기 요청. 한쪽만 있으면 미리보기.
    open_file_pair = pyqtSignal(str, str, str)
    title_changed = pyqtSignal()

    def __init__(self, status_bar, parent=None):
        super().__init__(parent)
        self.status = status_bar
        self._root_a = ""
        self._root_b = ""
        self._entries: list = []          # list[FolderEntry]
        self._pairs: list = []            # list[(item_a, item_b)]
        self._staged_files: dict = {}     # entry_index → 'a_to_b' | 'b_to_a' (병합 준비)
        self._merged_files: set = set()   # 병합 완료된 rel_path (연파랑 표시 + 필터 유지)
        self._diff_only = True            # 폴더 비교 기본: 변경점만 보기 ON
        self._syncing = False             # 선택/확장 미러 재귀 방지
        self._syncing_scroll = False      # 스크롤 미러 재귀 방지
        self._scan_worker = None          # 백그라운드 폴더 스캔 워커
        self._merge_worker = None         # 백그라운드 폴더 병합(복사) 워커
        self._scan_token = 0              # 낡은 스캔 결과 방어 토큰
        self._build_ui()
        self.setStyleSheet(APP_QSS)
        find_sc = QShortcut(QKeySequence("Ctrl+F"), self)
        find_sc.setContext(Qt.WidgetWithChildrenShortcut)
        find_sc.activated.connect(self._focus_find)

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # 파일 비교 뷰(DiffView + FilePanel)와 동일한 위치·구조:
        # 상단 툴바 + 좌/우 패널(제목·경로칸·구분선·리스트).
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # ── 상단 툴바 (DiffView 툴바와 동일 위치) ──
        toolbar = QHBoxLayout()
        self.diff_only_btn = QPushButton("변경점만 보기")
        self.diff_only_btn.setObjectName("toggle_btn")
        self.diff_only_btn.setCheckable(True)
        self.diff_only_btn.setFixedHeight(36)
        self.diff_only_btn.setFont(ui_font(10))
        self.diff_only_btn.setEnabled(False)   # A/B 폴더 모두 등록되면 활성화
        self.diff_only_btn.setToolTip("동일한 파일을 숨기고 변경/한쪽에만 있는 파일만 표시")
        self.diff_only_btn.toggled.connect(self._on_diff_only_toggled)
        toolbar.addWidget(self.diff_only_btn)

        self.refresh_btn = QPushButton("새로고침")
        self.refresh_btn.setFixedHeight(36)
        self.refresh_btn.setFont(ui_font(10))
        self.refresh_btn.setEnabled(False)
        self.refresh_btn.setToolTip("폴더를 다시 스캔해 목록을 갱신합니다")
        self.refresh_btn.clicked.connect(self._on_refresh)
        toolbar.addWidget(self.refresh_btn)

        # 이전/다음 변경 파일 이동 (파일 비교와 동일)
        self.prev_change_btn = QPushButton("◀ 이전 변경")
        self.prev_change_btn.setFixedHeight(36)
        self.prev_change_btn.setFont(ui_font(10))
        self.prev_change_btn.setEnabled(False)
        self.prev_change_btn.setToolTip("이전 변경 파일로 이동")
        self.prev_change_btn.clicked.connect(lambda: self._goto_change(-1))
        toolbar.addWidget(self.prev_change_btn)

        self.next_change_btn = QPushButton("다음 변경 ▶")
        self.next_change_btn.setFixedHeight(36)
        self.next_change_btn.setFont(ui_font(10))
        self.next_change_btn.setEnabled(False)
        self.next_change_btn.setToolTip("다음 변경 파일로 이동")
        self.next_change_btn.clicked.connect(lambda: self._goto_change(+1))
        toolbar.addWidget(self.next_change_btn)

        # 찾기 (파일명 검색) — 공용 구성. 폴더뷰는 '전체 단어' 토글 없음(with_word=False).
        toolbar.addSpacing(16)
        toolbar.addLayout(build_find_box(
            self, "파일명 찾기 (Ctrl+F)", self._goto_find,
            "파일 경로 검색 — Enter: 다음, Shift+Enter: 이전",
            with_word=False))

        toolbar.addStretch()
        # 범례 (공용) — 신규(연두)/변경(노랑)/준비(주황)/병합(연파랑)
        add_legend(toolbar, [
            ("신규", FOLDER_STATUS_COLORS["only_a"]),
            ("변경", FOLDER_STATUS_COLORS["modified"]),
            ("준비", DIFF_COLORS["staged"]),
            ("병합", DIFF_COLORS["merged"]),
        ])
        root.addLayout(toolbar)

        # ── 좌/우 패널 (FilePanel과 동형) ──
        splitter = QSplitter(Qt.Horizontal)
        panel_a = self._make_side_panel("a", "A 폴더 (원본)")
        panel_b = self._make_side_panel("b", "B 폴더 (비교)")
        splitter.addWidget(panel_a)
        splitter.addWidget(panel_b)
        splitter.setSizes([700, 700])
        root.addWidget(splitter, 1)

        # ── 동기화 배선 (두 트리 생성 후) ──
        self.tree_a.verticalScrollBar().valueChanged.connect(
            lambda v: self._sync_scroll(self.tree_b, v))
        self.tree_b.verticalScrollBar().valueChanged.connect(
            lambda v: self._sync_scroll(self.tree_a, v))
        for src, dst in ((self.tree_a, self.tree_b), (self.tree_b, self.tree_a)):
            src.itemSelectionChanged.connect(
                lambda s=src, d=dst: self._sync_selection(s, d))
            src.itemExpanded.connect(lambda it: self._mirror_expand(it, True))
            src.itemCollapsed.connect(lambda it: self._mirror_expand(it, False))
            src.itemActivated.connect(self._on_item_activated)

        # 기본 '변경점만 보기' ON — 트리 생성·배선 후 설정해야 _apply_filter가 안전.
        self.diff_only_btn.setChecked(True)

    def _make_side_panel(self, side: str, title: str) -> QWidget:
        """FilePanel과 동형의 한 쪽(A/B) 패널: 제목·경로칸·구분선·리스트.
        패널 전체가 폴더 드롭 대상(경로칸/트리 어디에 떨궈도 등록)."""
        w = _FolderDropPanel()
        w.folder_dropped.connect(lambda p, s=side: self._on_folder_chosen(s, p))
        lay = QVBoxLayout(w)
        lay.setContentsMargins(4, 4, 4, 4)
        lay.setSpacing(4)

        header = QHBoxLayout()
        title_lbl = QLabel(title)
        title_lbl.setFont(ui_font(10, bold=True))
        header.addWidget(title_lbl)
        hint = QLabel("  폴더를 여기에 끌어다 놓으세요")
        hint.setStyleSheet("color: #888; font-size: 9pt;")
        hint.setFont(ui_font(9))
        header.addWidget(hint)
        header.addStretch()
        lay.addLayout(header)

        path_row = QHBoxLayout()
        pe = _FolderDropLineEdit()
        pe.setFont(ui_font(10))
        pe.folder_dropped.connect(lambda p, s=side: self._on_folder_chosen(s, p))
        pe.returnPressed.connect(lambda s=side, e=pe: self._on_path_entered(s, e))
        browse_btn = QPushButton()
        browse_btn.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
        browse_btn.setFixedSize(32, 32)
        browse_btn.setIconSize(QSize(18, 18))
        browse_btn.setToolTip("폴더 찾아보기")
        browse_btn.clicked.connect(lambda _=False, s=side: self._browse(s))

        # 이 폴더로 병합 실행(파일 비교의 저장 버튼처럼 각 패널에 1개) — 주황 강조.
        merge_btn = QPushButton()
        merge_btn.setObjectName("save_btn")
        merge_btn.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        merge_btn.setFixedSize(32, 32)
        merge_btn.setIconSize(QSize(18, 18))
        merge_btn.setEnabled(False)
        merge_btn.setToolTip(f"{side.upper()} 폴더로 병합 실행 — 준비된 파일 복사")
        merge_btn.clicked.connect(lambda _=False, s=side: self._merge_execute(s))

        path_row.addWidget(pe)
        path_row.addWidget(browse_btn)
        path_row.addWidget(merge_btn)
        lay.addLayout(path_row)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        lay.addWidget(line)

        tree = self._make_tree()
        lay.addWidget(tree)

        if side == "a":
            self.path_a = pe
            self.tree_a = tree
            self.merge_btn_a = merge_btn
        else:
            self.path_b = pe
            self.tree_b = tree
            self.merge_btn_b = merge_btn
        return w

    def _make_tree(self) -> QTreeWidget:
        t = _FolderTree()
        t.setObjectName("folder_tree")
        t.setColumnCount(1)
        t.setHeaderHidden(True)   # 제목은 패널 헤더가 담당
        t.setSelectionMode(QAbstractItemView.ExtendedSelection)   # 다중 선택(병합 일괄)
        t.setUniformRowHeights(True)   # 좌우 스크롤 값 1:1 정렬을 위해
        t.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        t.setFont(ui_font(10))
        # 포커스 없는 반대쪽 트리도 선택색이 회색으로 흐려지지 않고 파랑 유지
        force_active_highlight(t)
        t.setItemDelegate(_HatchDelegate(t))   # placeholder 칸에 대각선 해치
        t.setContextMenuPolicy(Qt.CustomContextMenu)
        t.customContextMenuRequested.connect(
            lambda pos, tr=t: self._show_merge_menu(tr, pos))
        t.enter_pressed.connect(lambda tr=t: self._open_selected(tr))
        return t

    # ── 폴더 지정 ──────────────────────────────────────────────────────────────

    def _browse(self, side: str):
        start = self._root_a or self._root_b or ""
        d = QFileDialog.getExistingDirectory(
            self, f"{side.upper()} 폴더 선택", start)
        if d:
            self._on_folder_chosen(side, d)

    def _on_folder_chosen(self, side: str, path: str):
        self._set_side(side, path)
        self._rescan()

    def _on_path_entered(self, side: str, edit):
        """경로칸에 직접 입력한 폴더 경로 로드 (Enter). 폴더가 아니면 무시."""
        p = edit.text().strip().strip('"')
        if p and os.path.isdir(p):
            self._on_folder_chosen(side, p)

    def _on_refresh(self):
        """새로고침 버튼 — 병합 완료 표시(연파랑)까지 초기화하고 다시 스캔."""
        self._merged_files = set()
        self._rescan()

    def _set_side(self, side: str, path: str):
        path = os.path.abspath(path)
        # 폴더가 바뀌면 이전 병합 완료 표시는 의미가 없으므로 초기화.
        self._merged_files = set()
        if side == "a":
            self._root_a = path
            self.path_a.setText(path)
        else:
            self._root_b = path
            self.path_b.setText(path)

    # ── 스캔 & 트리 빌드 ─────────────────────────────────────────────────────────

    def _rescan(self, after=None):
        # 스캔하면 병합 준비는 초기화(엔트리 인덱스가 바뀌므로).
        self._staged_files = {}
        self._update_merge_btn()
        # 버튼 상태: 변경점만 보기는 A/B 모두 등록 시, 새로고침은 하나라도 등록 시.
        self.diff_only_btn.setEnabled(bool(self._root_a) and bool(self._root_b))
        self.refresh_btn.setEnabled(bool(self._root_a) or bool(self._root_b))
        if not self._root_a and not self._root_b:
            self._set_nav_enabled(False)
            if after:
                after()
            return
        # 백그라운드 스캔 — UI가 얼지 않도록. 최신 스캔만 반영(토큰).
        self._scan_token += 1
        token = self._scan_token
        self.status.showMessage("폴더 비교 중…")
        self.status.begin_busy()
        w = FolderScanWorker(self._root_a, self._root_b)
        w.done.connect(lambda es, t=token, cb=after: self._on_scan_done(es, t, cb))
        w.progress.connect(self._on_scan_progress)
        w.error.connect(self._on_scan_error)
        w.finished.connect(w.deleteLater)
        self._scan_worker = w
        w.start()

    def _on_scan_progress(self, done: int, total: int):
        self.status.showMessage(f"폴더 비교 중… {done:,}/{total:,}")
        self.status.set_progress(done, total)

    def _on_scan_error(self, msg: str):
        self.status.end_progress()
        self.status.showMessage(f"폴더 비교 오류: {msg}")

    def _on_scan_done(self, entries, token, after=None):
        # 더 새 스캔이 시작됐으면(낡은 결과) 무시 — 낡은 after 콜백도 폐기.
        if token != self._scan_token:
            return
        self.status.end_progress()
        self._entries = entries
        self._set_nav_enabled(bool(self._entries))
        self._populate_trees()
        c = summarize(self._entries)
        self.status.showMessage(
            f"폴더 비교 — 동일 {c[SAME]} · 변경 {c[MODIFIED]} · "
            f"A만 {c[ONLY_A]} · B만 {c[ONLY_B]}  (총 {len(self._entries)}개 파일)"
        )
        self.title_changed.emit()
        if after:
            after()

    def _populate_trees(self):
        self.tree_a.clear()
        self.tree_b.clear()
        self._pairs = []
        node_root = _build_node_tree(self._entries)
        self._add_children(
            node_root,
            self.tree_a.invisibleRootItem(),
            self.tree_b.invisibleRootItem(),
        )
        # expandAll이 itemExpanded 폭풍을 일으키지 않도록 미러 가드.
        self._syncing = True
        try:
            self.tree_a.expandAll()
            self.tree_b.expandAll()
        finally:
            self._syncing = False
        self._apply_filter()

    def _add_children(self, node: _Node, parent_a, parent_b):
        for child in node.children.values():
            ia = QTreeWidgetItem(parent_a)
            ib = QTreeWidgetItem(parent_b)
            pair_idx = len(self._pairs)
            self._pairs.append((ia, ib))
            ia.setData(0, _ROLE_PAIR, pair_idx)
            ib.setData(0, _ROLE_PAIR, pair_idx)
            if child.is_file:
                e = self._entries[child.entry_index]
                ia.setData(0, _ROLE_ENTRY, child.entry_index)
                ib.setData(0, _ROLE_ENTRY, child.entry_index)
                self._decorate_file(ia, child.name, e, "a")
                self._decorate_file(ib, child.name, e, "b")
            else:
                ia.setData(0, _ROLE_ENTRY, -1)
                ib.setData(0, _ROLE_ENTRY, -1)
                ia.setText(0, "📁 " + child.name)
                ib.setText(0, "📁 " + child.name)
                self._add_children(child, ia, ib)

    def _decorate_file(self, item, name: str, entry, side: str):
        present = entry.path_a if side == "a" else entry.path_b
        merged = entry.rel_path in self._merged_files
        item.setData(0, _ROLE_PLACEHOLDER, False)
        if present:
            item.setText(0, name)
            color = DIFF_COLORS["merged"] if merged else FOLDER_STATUS_COLORS[entry.status]
            item.setBackground(0, QBrush(color))
            item.setToolTip(0, present)
        elif merged:
            # 병합됨 표시(연파랑) — 반대쪽에 갓 복사돼 나타난 파일
            item.setText(0, name)
            item.setBackground(0, QBrush(DIFF_COLORS["merged"]))
        else:
            # 반대쪽에만 있는 파일 — 이쪽은 빈 placeholder(델리게이트가 대각선 해치)
            item.setText(0, "")
            item.setData(0, _ROLE_PLACEHOLDER, True)

    # ── 필터: 변경된 파일만 보기 ─────────────────────────────────────────────────

    def _on_diff_only_toggled(self, checked: bool):
        self._diff_only = checked
        self._apply_filter()

    def _apply_filter(self):
        for ia, ib in self._pairs:
            eidx = ia.data(0, _ROLE_ENTRY)
            if eidx is not None and eidx >= 0:
                e = self._entries[eidx]
                # 병합됨(연파랑) 파일은 동일이 됐어도 계속 표시 — 병합 결과 확인용.
                hide = (self._diff_only and e.status == SAME
                        and e.rel_path not in self._merged_files)
                ia.setHidden(hide)
                ib.setHidden(hide)
        # 보이는 파일이 하나도 없는 폴더 노드는 숨김(양쪽 동일 구조라 각각 계산해도 일치).
        self._prune_empty_folders(self.tree_a.invisibleRootItem())
        self._prune_empty_folders(self.tree_b.invisibleRootItem())

    def _prune_empty_folders(self, item) -> bool:
        """item 하위에 보이는 '파일'이 하나라도 있으면 True. 빈 폴더는 숨긴다."""
        any_visible = False
        for i in range(item.childCount()):
            ch = item.child(i)
            eidx = ch.data(0, _ROLE_ENTRY)
            if eidx is not None and eidx >= 0:      # 파일
                if not ch.isHidden():
                    any_visible = True
            else:                                    # 폴더
                vis = self._prune_empty_folders(ch)
                ch.setHidden(not vis)
                if vis:
                    any_visible = True
        return any_visible

    # ── 동기화 ──────────────────────────────────────────────────────────────────

    def _sync_scroll(self, dst, value):
        if self._syncing_scroll:
            return
        self._syncing_scroll = True
        try:
            dst.verticalScrollBar().setValue(value)
        finally:
            self._syncing_scroll = False

    def _sync_selection(self, src, dst):
        if self._syncing:
            return
        self._syncing = True
        try:
            dst.clearSelection()
            partner_cur = None
            for it in src.selectedItems():
                pidx = it.data(0, _ROLE_PAIR)
                if pidx is None:
                    continue
                ia, ib = self._pairs[pidx]
                partner = ib if dst is self.tree_b else ia
                partner.setSelected(True)
                partner_cur = partner
            cur = src.currentItem()
            if cur is not None:
                pidx = cur.data(0, _ROLE_PAIR)
                if pidx is not None:
                    ia, ib = self._pairs[pidx]
                    partner_cur = ib if dst is self.tree_b else ia
            if partner_cur is not None:
                # NoUpdate: 현재 항목만 옮기고 선택은 건드리지 않음
                # (기본 setCurrentItem은 ClearAndSelect라 다중 선택이 1개로 붕괴됨).
                dst.setCurrentItem(partner_cur, 0, QItemSelectionModel.NoUpdate)
        finally:
            self._syncing = False

    def _mirror_expand(self, item, expanded: bool):
        if self._syncing:
            return
        self._syncing = True
        try:
            pidx = item.data(0, _ROLE_PAIR)
            if pidx is not None:
                ia, ib = self._pairs[pidx]
                for it in (ia, ib):
                    if it is not item:
                        it.setExpanded(expanded)
        finally:
            self._syncing = False

    # ── 열기 ────────────────────────────────────────────────────────────────────

    def _on_item_activated(self, item, _col):
        eidx = item.data(0, _ROLE_ENTRY)
        if eidx is None or eidx < 0:
            return   # 폴더/빈칸 — 무시(폴더는 기본 확장 동작)
        e = self._entries[eidx]
        self.open_file_pair.emit(e.path_a, e.path_b, e.rel_path)

    def _open_selected(self, tree):
        """Enter — 선택된 모든 파일을 각각 비교 탭으로 연다(중복은 MainWindow가 재사용)."""
        eidxs = self._selected_file_eidxs(tree)
        if not eidxs:
            return
        for eidx in eidxs:
            e = self._entries[eidx]
            self.open_file_pair.emit(e.path_a, e.path_b, e.rel_path)

    # ── 이전/다음 변경 이동 · 찾기 (파일 비교와 동일) ─────────────────────────────

    def _set_nav_enabled(self, enabled: bool):
        for w in (self.prev_change_btn, self.next_change_btn, self.find_edit,
                  self.find_case_btn, self.find_prev_btn, self.find_next_btn):
            w.setEnabled(enabled)

    def _changed_pairs(self):
        """변경(status != same)인 보이는 파일 pair 인덱스를 순서대로."""
        out = []
        for pidx, (ia, _ib) in enumerate(self._pairs):
            eidx = ia.data(0, _ROLE_ENTRY)
            if (eidx is not None and eidx >= 0 and not ia.isHidden()
                    and self._entries[eidx].status != SAME):
                out.append(pidx)
        return out

    def _current_pair(self) -> int:
        cur = self.tree_a.currentItem() or self.tree_b.currentItem()
        if cur is not None:
            p = cur.data(0, _ROLE_PAIR)
            if p is not None:
                return p
        return -1

    def _select_pair(self, pidx: int):
        ia, ib = self._pairs[pidx]
        self._syncing = True
        try:
            self.tree_a.setCurrentItem(ia)
            ia.setSelected(True)
            self.tree_b.setCurrentItem(ib)
            ib.setSelected(True)
            self.tree_a.scrollToItem(ia, QAbstractItemView.PositionAtCenter)
            self.tree_b.scrollToItem(ib, QAbstractItemView.PositionAtCenter)
        finally:
            self._syncing = False

    def _goto_change(self, direction: int):
        if not self._entries:
            return
        order = self._changed_pairs()
        if not order:
            self.status.showMessage("변경된 파일이 없습니다.")
            return
        cur = self._current_pair()
        if direction > 0:
            target = next((p for p in order if p > cur), None)
            if target is None:
                self.status.showMessage("마지막 변경 파일입니다.")
                return
        else:
            target = next((p for p in reversed(order) if p < cur), None)
            if target is None:
                self.status.showMessage("첫 변경 파일입니다.")
                return
        self._select_pair(target)
        self.tree_a.setFocus()

    def _focus_find(self):
        if self.find_edit.isEnabled():
            self.find_edit.setFocus()
            self.find_edit.selectAll()

    def _find_matches(self, term: str):
        ignore_case = self.find_case_btn.isChecked()
        needle = term.casefold() if ignore_case else term
        out = []
        for pidx, (ia, _ib) in enumerate(self._pairs):
            eidx = ia.data(0, _ROLE_ENTRY)
            if eidx is None or eidx < 0 or ia.isHidden():
                continue
            hay = self._entries[eidx].rel_path
            hay = hay.casefold() if ignore_case else hay
            if needle in hay:
                out.append(pidx)
        return out

    def _goto_find(self, direction: int):
        term = self.find_edit.text()
        if not term or not self._entries:
            return
        matches = self._find_matches(term)
        if not matches:
            self.status.showMessage(f'"{term}" — 일치 항목이 없습니다.')
            return
        cur = self._current_pair()
        wrapped = ""
        if direction > 0:
            target = next((p for p in matches if p > cur), None)
            if target is None:
                target = matches[0]
                wrapped = " — 처음부터 다시 검색합니다."
        else:
            target = next((p for p in reversed(matches) if p < cur), None)
            if target is None:
                target = matches[-1]
                wrapped = " — 끝에서부터 다시 검색합니다."
        self._select_pair(target)
        idx = matches.index(target) + 1
        self.status.showMessage(f'찾기: "{term}" {idx}/{len(matches)}개 일치{wrapped}')

    # ── 병합(파일 복사): 우클릭 준비 → '병합 실행' ──────────────────────────────

    def _pair_for_entry(self, eidx: int):
        for ia, ib in self._pairs:
            if ia.data(0, _ROLE_ENTRY) == eidx:
                return ia, ib
        return None

    def _selected_file_eidxs(self, tree) -> list:
        out = []
        for it in tree.selectedItems():
            e = it.data(0, _ROLE_ENTRY)
            if e is not None and e >= 0:
                out.append(e)
        return out

    def _show_merge_menu(self, tree, pos):
        item = tree.itemAt(pos)
        if item is None:
            return
        eidxs = self._selected_file_eidxs(tree)
        clicked = item.data(0, _ROLE_ENTRY)
        if clicked is not None and clicked >= 0 and clicked not in eidxs:
            eidxs = [clicked]
        eidxs = [e for e in eidxs if e is not None and e >= 0]
        if not eidxs:
            return
        # 방향 유효성: A→B는 모든 대상이 A에 존재해야, B→A는 B에 존재해야.
        can_a2b = all(self._entries[e].path_a for e in eidxs)
        can_b2a = all(self._entries[e].path_b for e in eidxs)
        all_staged = all(e in self._staged_files for e in eidxs)
        has_staged = any(e in self._staged_files for e in eidxs)
        menu = QMenu(self)
        menu.setStyleSheet(MENU_QSS)
        if all_staged:
            # 이미 전부 병합 준비 상태 → 취소만 노출.
            act_a2b = act_b2a = None
            act_unstage = menu.addAction("병합 준비 취소")
        else:
            act_a2b = menu.addAction("A → B 병합 준비") if can_a2b else None
            act_b2a = menu.addAction("B → A 병합 준비") if can_b2a else None
            act_unstage = menu.addAction("병합 준비 취소") if has_staged else None
        if act_a2b is None and act_b2a is None and act_unstage is None:
            return
        chosen = menu.exec_(tree.viewport().mapToGlobal(pos))
        if chosen is None:
            return
        if chosen is act_a2b:
            self._stage_files(eidxs, DIR_A2B)
        elif chosen is act_b2a:
            self._stage_files(eidxs, DIR_B2A)
        elif chosen is act_unstage:
            self._unstage_files(eidxs)

    def _stage_files(self, eidxs, direction: str):
        for e in eidxs:
            entry = self._entries[e]
            src = entry.path_a if direction == DIR_A2B else entry.path_b
            if not src:
                continue   # 해당 방향 소스 없음 — 건너뜀
            self._staged_files[e] = direction
            self._restyle_entry(e)
        # 선택 해제 — 파란 선택색이 주황 준비색을 가리지 않도록(셀/파일 병합 준비와 동일).
        self._syncing = True
        try:
            self.tree_a.clearSelection()
            self.tree_b.clearSelection()
        finally:
            self._syncing = False
        self._update_merge_btn()
        self.status.showMessage(
            f"병합 준비 — {len(self._staged_files)}개 파일 대기 중  | '병합 실행'으로 복사")

    def _unstage_files(self, eidxs):
        for e in eidxs:
            if e in self._staged_files:
                del self._staged_files[e]
                self._restyle_entry(e)
        self._update_merge_btn()

    def _restyle_entry(self, eidx: int):
        pair = self._pair_for_entry(eidx)
        if not pair:
            return
        ia, ib = pair
        entry = self._entries[eidx]
        if eidx in self._staged_files:
            arrow = "A → B" if self._staged_files[eidx] == DIR_A2B else "B → A"
            # 반대쪽(비어 있던) 패널에도 파일명을 노출 — 무엇이 복사될지 보이도록.
            for it in (ia, ib):
                it.setText(0, entry.name)
                it.setBackground(0, QBrush(DIFF_COLORS["staged"]))
                it.setData(0, _ROLE_PLACEHOLDER, False)   # 해치 대신 준비색
                it.setToolTip(0, f"병합 준비: {arrow}")
        else:
            self._decorate_file(ia, entry.name, entry, "a")
            self._decorate_file(ib, entry.name, entry, "b")

    def _update_merge_btn(self):
        # A 폴더로 병합 = b_to_a 준비분 / B 폴더로 병합 = a_to_b 준비분.
        has_to_a = any(d == DIR_B2A for d in self._staged_files.values())
        has_to_b = any(d == DIR_A2B for d in self._staged_files.values())
        self.merge_btn_a.setEnabled(has_to_a)
        self.merge_btn_b.setEnabled(has_to_b)

    def _merge_execute(self, side: str):
        """side='a'면 B→A(b_to_a) 준비분을, 'b'면 A→B(a_to_b) 준비분을 복사.
        실제 복사는 FolderMergeWorker(백그라운드)에서 — UI 프리즈 방지."""
        target_dir = DIR_B2A if side == "a" else DIR_A2B
        items = [(e, d) for e, d in self._staged_files.items() if d == target_dir]
        if not items:
            return
        dest = "A" if side == "a" else "B"
        if QMessageBox.question(
                self, "병합 실행",
                f"준비된 {len(items)}개 파일을 {dest} 폴더로 복사(덮어쓰기)합니다.\n"
                f"계속하시겠습니까?") != QMessageBox.Yes:
            return
        # 이번에 실행하지 않는 반대 방향 준비분은 스캔 후 rel_path로 재적용.
        remaining = {self._entries[e].rel_path: d
                     for e, d in self._staged_files.items() if d != target_dir}
        # 복사 목록을 UI 스레드에서 미리 해결(스레드 밖으로 FolderEntry 접근 금지).
        pairs = []
        for eidx, direction in items:
            entry = self._entries[eidx]
            rel = entry.rel_path.replace("/", os.sep)
            if direction == DIR_A2B:
                src, dst = entry.path_a, os.path.join(self._root_b, rel)
            else:
                src, dst = entry.path_b, os.path.join(self._root_a, rel)
            pairs.append((src, dst, entry.rel_path))

        self.merge_btn_a.setEnabled(False)
        self.merge_btn_b.setEnabled(False)
        self.status.showMessage(f"병합 중… ({dest} 폴더)")
        self.status.begin_busy()

        w = FolderMergeWorker(pairs)
        w.progress.connect(self.status.set_progress)
        w.done.connect(lambda d, f, m, dest=dest, rem=remaining:
                       self._on_merge_done(d, f, m, dest, rem))
        w.error.connect(self._on_merge_error)
        w.finished.connect(w.deleteLater)
        self._merge_worker = w
        w.start()

    def _on_merge_done(self, done, fails, merged_rels, dest, remaining):
        self.status.end_progress()
        self._merged_files.update(merged_rels)   # 연파랑 '병합' 표시 유지

        def _after():
            # 재스캔(비동기) 완료 후 실행: 남은 반대 방향 준비분을 새 엔트리 인덱스로 복원.
            if remaining:
                idx_of = {e.rel_path: i for i, e in enumerate(self._entries)}
                for rel, d in remaining.items():
                    i = idx_of.get(rel)
                    if i is not None:
                        self._staged_files[i] = d
                        self._restyle_entry(i)
                self._update_merge_btn()
            if fails:
                QMessageBox.warning(
                    self, "병합 일부 실패",
                    f"{done}개 복사됨.\n실패 {len(fails)}개:\n" + "\n".join(fails[:10]))
            else:
                self.status.showMessage(f"병합 완료 — {done}개 파일 복사됨 ({dest} 폴더)")

        self._rescan(after=_after)   # 상태 갱신(복사된 파일은 동일로) — _staged_files 초기화됨

    def _on_merge_error(self, msg: str):
        self.status.end_progress()
        self._update_merge_btn()
        self.status.showMessage(f"병합 오류: {msg}")
        QMessageBox.critical(self, "병합 오류", f"파일 복사 실패:\n{msg}")

    # ── MainWindow 공개 API ─────────────────────────────────────────────────────

    def shutdown(self):
        """탭/앱 종료 시 실행 중 워커를 안전하게 정리(QThread 파괴 크래시 방지).
        취소 미지원 — 스캔/병합이 끝날 때까지 대기한다. 특히 병합은 파일을 쓰는 중이므로
        중단하지 않고 완주시킨다(저장 워커와 동일 원칙)."""
        for w in (self._scan_worker, self._merge_worker):
            if w is not None:
                try:
                    w.blockSignals(True)
                    if w.isRunning():
                        w.wait()
                except RuntimeError:
                    pass

    def folder_a(self) -> str:
        return self._root_a

    def folder_b(self) -> str:
        return self._root_b

    def set_folders(self, path_a: str = "", path_b: str = ""):
        if path_a:
            self._set_side("a", path_a)
        if path_b:
            self._set_side("b", path_b)
        self._rescan()

    def tab_title(self) -> str:
        a = os.path.basename(self._root_a.rstrip("/\\")) if self._root_a else ""
        b = os.path.basename(self._root_b.rstrip("/\\")) if self._root_b else ""
        if a and b:
            return f"📁 {a} ↔ {b}"
        if a:
            return f"📁 {a}"
        if b:
            return f"📁 {b}"
        return "📁 폴더 비교"
