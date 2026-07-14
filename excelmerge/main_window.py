"""메인 윈도우 — 탭 컨테이너 셸.

여러 개의 "파일 비교"(DiffView)와 "폴더 비교"(FolderCompareView) 탭을 담는다.
실제 비교/병합 UI·로직은 각 탭 위젯에 있고, MainWindow는 다음만 담당한다:
 - 상단 탭 스트립(QTabWidget, 닫기 버튼) + 우상단 "＋파일 / ＋폴더" 코너 버튼
 - 모든 탭이 공유하는 하단 상태바(활성 탭이 갱신)
 - 탭 열기(중복 방지)/닫기, 창 제목 갱신
 - 커맨드라인(-s/-d) 진입 — 파일이면 파일 비교 탭, 폴더면 폴더 비교 탭
"""
import os

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QPushButton, QTabWidget,
    QTabBar, QToolButton, QShortcut,
)
from PyQt5.QtGui import QKeySequence, QIcon
from PyQt5.QtCore import Qt, QSize

# 탭 닫기 버튼 — 윈도우 표준 느낌의 은은한 회색 ✕ (hover 시 연회색 배경).
_TAB_CLOSE_QSS = (
    "QToolButton { border: none; background: transparent; color: #909090;"
    " font-size: 12px; padding: 0; }"
    "QToolButton:hover { background: #d0d0d0; border-radius: 3px; color: #333; }"
    "QToolButton:pressed { background: #b8b8b8; }"
)

from . import __version__
from .diff_view import DiffView
from .statusbar import StatusBar
from .theme import APP_QSS, load_app_icon, ui_font


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # 타이틀바는 툴명 + 버전으로 고정 (탭의 파일/폴더명은 노출하지 않음).
        self.setWindowTitle(f"ExcelMerge v{__version__}")
        # 작업 표시줄·Alt+Tab에서도 타이틀바와 동일한 앱 아이콘이 보이도록 명시 지정.
        self.setWindowIcon(load_app_icon())
        self.resize(1400, 800)
        self.setStyleSheet(APP_QSS)

        # ── 공유 상태바 — 모든 탭이 이 인스턴스에 메시지를 쓴다(우측 진행률 바 내장) ──
        self.status = StatusBar()
        self.setStatusBar(self.status)

        # ── 탭 컨테이너 ──
        self.tabs = QTabWidget()
        self.tabs.setObjectName("main_tabs")
        # 기본 닫기 버튼 대신 커스텀(표준 느낌) ✕ 버튼을 탭마다 직접 단다.
        self.tabs.setTabsClosable(False)
        self.tabs.setMovable(True)
        self.tabs.setDocumentMode(True)
        self.setCentralWidget(self.tabs)

        # ── 우상단 코너: 새 비교 탭 열기 버튼 (파일/폴더 통합) ──
        corner = QWidget()
        corner_lay = QHBoxLayout(corner)
        corner_lay.setContentsMargins(4, 2, 4, 2)
        corner_lay.setSpacing(4)
        new_btn = QPushButton("＋ 새 비교")
        new_btn.setFont(ui_font(9))
        new_btn.setToolTip("새 비교 탭 열기 — 파일을 넣으면 셀 비교, 폴더를 넣으면 파일 목록 비교")
        new_btn.clicked.connect(self._new_tab)
        corner_lay.addWidget(new_btn)
        self.tabs.setCornerWidget(corner, Qt.TopRightCorner)

        # ── 탭 이동/닫기 단축키 (탭이 많을 때 빠른 순회) ──
        # Ctrl+Tab: 다음 탭 / Ctrl+Shift+Tab: 이전 탭 / Ctrl+W: 현재 탭 닫기.
        self._install_tab_shortcuts()

        # 시작 시 빈 비교 탭 하나 — 파일/폴더 아무거나 드롭하면 그에 맞게 동작
        self._new_tab()

    # ── 탭 단축키 ──────────────────────────────────────────────────────────────

    def _install_tab_shortcuts(self):
        specs = [
            (QKeySequence("Ctrl+Tab"), lambda: self._cycle_tab(1)),
            (QKeySequence("Ctrl+Shift+Tab"), lambda: self._cycle_tab(-1)),
            # 일부 환경에서 Ctrl+Tab이 위젯에 먼저 먹히는 경우를 위한 표준 대체 바인딩.
            (QKeySequence(Qt.CTRL | Qt.Key_PageDown), lambda: self._cycle_tab(1)),
            (QKeySequence(Qt.CTRL | Qt.Key_PageUp), lambda: self._cycle_tab(-1)),
            (QKeySequence("Ctrl+W"), self._close_current_tab),
        ]
        for seq, slot in specs:
            sc = QShortcut(seq, self)
            sc.setContext(Qt.ApplicationShortcut)
            sc.activated.connect(slot)

    def _cycle_tab(self, step: int):
        """활성 탭을 step(±1)만큼 이동 — 양끝에서 순환(wrap)."""
        n = self.tabs.count()
        if n <= 1:
            return
        self.tabs.setCurrentIndex((self.tabs.currentIndex() + step) % n)

    def _close_current_tab(self):
        idx = self.tabs.currentIndex()
        if idx >= 0:
            self._close_tab(idx)

    # ── 탭 열기 ────────────────────────────────────────────────────────────────

    def _add_tab(self, widget: QWidget, title: str) -> QWidget:
        idx = self.tabs.addTab(widget, title)
        self._install_close_button(widget)
        self._apply_tab_icon(widget)
        self.tabs.setCurrentIndex(idx)
        return widget

    def _apply_tab_icon(self, view: QWidget):
        """탭 라벨 좌측의 파일 형식 아이콘 갱신.
        DiffView는 tab_icon()(파일형식 PNG)을 제공하고, 폴더 비교 등은 없으므로
        null 아이콘으로 두어 자체 이모지 라벨(📁)을 유지한다."""
        idx = self.tabs.indexOf(view)
        if idx < 0:
            return
        icon = view.tab_icon() if hasattr(view, "tab_icon") else QIcon()
        self.tabs.setTabIcon(idx, icon)

    def _install_close_button(self, view: QWidget):
        """탭에 표준 느낌의 커스텀 닫기 버튼(✕)을 단다."""
        idx = self.tabs.indexOf(view)
        if idx < 0:
            return
        btn = QToolButton()
        btn.setText("✕")
        btn.setToolTip("탭 닫기")
        btn.setCursor(Qt.ArrowCursor)
        btn.setFocusPolicy(Qt.NoFocus)
        btn.setFixedSize(QSize(18, 18))
        btn.setStyleSheet(_TAB_CLOSE_QSS)
        btn.clicked.connect(lambda _=False, v=view: self._close_view(v))
        self.tabs.tabBar().setTabButton(idx, QTabBar.RightSide, btn)

    def _close_view(self, view: QWidget):
        idx = self.tabs.indexOf(view)
        if idx >= 0:
            self._close_tab(idx)

    def _new_tab(self) -> DiffView:
        """빈 비교 탭을 열고 활성화한다(그리드 UI로 시작; 폴더 드롭 시 목록으로 전환)."""
        return self.open_file_compare()

    def _wire_diff_view(self, view: "DiffView"):
        """DiffView 탭 공통 배선 — 제목 갱신 + 폴더 드롭 시 폴더 모드 전환."""
        view.panel_a.file_loaded.connect(lambda _p, v=view: self._update_tab_title(v))
        view.panel_b.file_loaded.connect(lambda _p, v=view: self._update_tab_title(v))
        view.folder_requested.connect(
            lambda side, p, v=view: self._convert_to_folder(v, side, p))

    def open_file_compare(self, path_a: str = "", path_b: str = "") -> DiffView:
        """파일 비교 탭을 연다. 동일 (A, B) 경로 탭이 이미 있으면 그 탭으로 전환한다."""
        existing = self._find_file_tab(path_a, path_b)
        if existing is not None:
            self.tabs.setCurrentWidget(existing)
            return existing

        # 경로가 주어졌고 현재 활성 탭이 '빈' 파일 탭이면 새 탭 대신 그 탭을 재사용
        # (시작 시 빈 탭 + CLI 로드로 탭이 둘로 늘어나는 것을 방지).
        if path_a or path_b:
            cur = self.tabs.currentWidget()
            if (isinstance(cur, DiffView)
                    and not cur.panel_a.get_path() and not cur.panel_b.get_path()):
                cur.load_pair(path_a, path_b)
                self._update_tab_title(cur)
                return cur

        view = DiffView(self.status)
        self._wire_diff_view(view)
        self._add_tab(view, view.tab_title())
        if path_a or path_b:
            view.load_pair(path_a, path_b)
            self._update_tab_title(view)
        return view

    def _convert_to_folder(self, view: "DiffView", side: str, path: str):
        """파일 그리드 탭에 폴더가 들어오면 그 탭 자리를 폴더 비교 뷰로 교체한다."""
        from .folder_view import FolderCompareView
        idx = self.tabs.indexOf(view)
        if idx < 0:
            return
        fv = FolderCompareView(self.status)
        fv.open_file_pair.connect(lambda a, b, _t: self.open_file_compare(a, b))
        fv.title_changed.connect(lambda v=fv: self._update_tab_title(v))
        self.tabs.removeTab(idx)
        self.tabs.insertTab(idx, fv, fv.tab_title())
        self._install_close_button(fv)
        self._apply_tab_icon(fv)
        self.tabs.setCurrentIndex(idx)
        view.deleteLater()
        fv.set_folders(path if side == "a" else "", path if side == "b" else "")
        self._update_tab_title(fv)

    def open_folder_compare(self, path_a: str = "", path_b: str = ""):
        """폴더 비교 탭을 연다. 동일 (A, B) 폴더 탭이 이미 있으면 그 탭으로 전환한다."""
        # FolderCompareView는 Step 4에서 추가 — 지연 import로 파일 비교 경로와 분리.
        from .folder_view import FolderCompareView

        existing = self._find_folder_tab(path_a, path_b)
        if existing is not None:
            self.tabs.setCurrentWidget(existing)
            return existing

        view = FolderCompareView(self.status)
        # 폴더 뷰에서 파일 더블클릭 → 파일 비교 탭 열기(중복 시 재사용).
        view.open_file_pair.connect(
            lambda a, b, _title: self.open_file_compare(a, b))
        view.title_changed.connect(lambda v=view: self._update_tab_title(v))
        self._add_tab(view, view.tab_title())
        if path_a or path_b:
            view.set_folders(path_a, path_b)
            self._update_tab_title(view)
        return view

    # ── 중복 탭 탐색 ──────────────────────────────────────────────────────────

    @staticmethod
    def _norm(path: str) -> str:
        return os.path.normcase(os.path.abspath(path)) if path else ""

    def _find_file_tab(self, path_a: str, path_b: str):
        """같은 (A, B) 파일 쌍의 DiffView 탭을 찾는다. 빈 경로면 중복 판단하지 않음."""
        if not path_a and not path_b:
            return None
        key = (self._norm(path_a), self._norm(path_b))
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if isinstance(w, DiffView):
                wk = (self._norm(w.panel_a.get_path()), self._norm(w.panel_b.get_path()))
                if wk == key:
                    return w
        return None

    def _find_folder_tab(self, path_a: str, path_b: str):
        from .folder_view import FolderCompareView
        if not path_a and not path_b:
            return None
        key = (self._norm(path_a), self._norm(path_b))
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if isinstance(w, FolderCompareView):
                wk = (self._norm(w.folder_a()), self._norm(w.folder_b()))
                if wk == key:
                    return w
        return None

    # ── 탭 닫기/제목 ──────────────────────────────────────────────────────────

    def _close_tab(self, index: int):
        w = self.tabs.widget(index)
        self.tabs.removeTab(index)
        if w is not None:
            if hasattr(w, "shutdown"):
                w.shutdown()   # 실행 중 워커 안전 종료(크래시 방지)
            w.deleteLater()
        # 항상 최소 1개 유지 — 빈 화면 방지(기존엔 창=단일 뷰였음)
        if self.tabs.count() == 0:
            self._new_tab()

    def closeEvent(self, event):
        """앱 종료 시 모든 탭의 실행 중 워커를 정리한 뒤 닫는다(QThread 파괴 크래시 방지)."""
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if w is not None and hasattr(w, "shutdown"):
                w.shutdown()
        # 업데이트 워커(있으면)도 안전 종료.
        for attr in ("_update_check_worker", "_update_dl_worker"):
            uw = getattr(self, attr, None)
            if uw is not None:
                try:
                    uw.blockSignals(True)
                    if uw.isRunning():
                        uw.wait()
                except RuntimeError:
                    pass
        super().closeEvent(event)

    def _update_tab_title(self, view: QWidget):
        # 탭 스트립의 라벨만 갱신 (타이틀바는 'ExcelMerge v<버전>'으로 고정).
        idx = self.tabs.indexOf(view)
        if idx < 0:
            return
        title = view.tab_title() if hasattr(view, "tab_title") else "탭"
        self.tabs.setTabText(idx, title)
        self._apply_tab_icon(view)
