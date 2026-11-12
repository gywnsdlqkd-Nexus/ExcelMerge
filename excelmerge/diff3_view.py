"""3파일 비교 뷰 — 왼쪽 · 가운데(기준) · 오른쪽. WinMerge · P4V 의 3-way 병합.

가운데 = 공통 기준(base). 양옆 = 그 기준에서 갈라진 두 버전(P4V 병합이면 왼쪽 theirs, 오른쪽
yours). 병합 결과는 가운데에 쓴다 — 결과 파일(P4V 의 %r)을 따로 받았으면 그 파일에 쓰고
가운데 원본은 건드리지 않는다.

**2파일 화면(DiffView)과 코드를 나누지 않는다.** 셀·행·열 대응이 3원소라 2파일 경로에 섞으면
조용히 틀린 값이 나온다(diff3_engine 머리말). 표·틀 고정·미니맵·경로칸 같은 부품만 같이 쓴다.

저장 규칙 — **결과 = 가운데 원본 + 지금까지의 모든 결정.**
  · 가운데 원본은 처음 열 때 이 화면 전용 임시 폴더에 떠 둔다(_template). 비교도 이 사본으로
    한다. 그래서 가운데 파일에 저장해도 비교 기준은 그대로고, 병합한 칸은 파랑으로 남는다.
  · 저장할 때마다 원본에서 새로 만든다. 결정 좌표가 원본 기준이라, 이미 고친 파일 위에 다시
    쓰면 지운·붙인 행만큼 어긋난다(2파일에서 실측된 사고). 병합을 취소하고 다시 저장하면
    결과에서도 빠진다.
  · 시트를 옮겨 다녀도 결정은 시트별로 남고(_sheet_cache), 저장은 모든 시트를 함께 쓴다.
  · 결과 파일이 이 화면 밖에서 바뀌었으면 덮어쓰지 않는다.
"""
import os
import shutil
import stat
import tempfile

from PyQt5.QtCore import Qt, QThread
from PyQt5.QtGui import QIcon, QKeySequence
from PyQt5.QtWidgets import (
    QAbstractItemView, QApplication, QHBoxLayout, QLabel, QMessageBox, QPushButton,
    QShortcut, QSizePolicy, QSplitter, QVBoxLayout, QWidget,
)

from .colref import get_column_letter
from .compare_toolbar import add_legend, build_find_box, make_find_matcher
from .diff3_engine import (S3_BOTH, S3_CONFLICT, S3_LEFT, S3_RIGHT, S3_SAME, SRC_LEFT,
                           SRC_RIGHT, auto_merge_plan, match_columns3, usable_col_meta3)
from .diff3_model import ROLE_BASE, ROLE_LEFT, ROLE_RIGHT, Diff3TableModel
from .diff3_workers import Diff3Worker, Load3Worker, Merge3Worker, Sheet3DiffWorker
from .diff_engine import count_changed_masked, keep_mask
from .loaders import clear_values_cache, list_sheet_names
from .merge3_build import expand_row_removals
from .panels import FilePanel
from .prefs import (add_path_history, load_last_key, load_last_key_name,
                    load_last_sheet, load_sheet_key, save_last_sheet, save_sheet_key)
from .theme import APP_QSS, DIFF_COLORS, ext_tab_icon, ui_font
from .widgets import FreezeController, MinimapScrollBar, SheetTabBar
from .xlsx_writer import _is_file_locked

_WRITABLE = (".xlsx", ".xlsm")
_SRC_IDX = {SRC_LEFT: 0, SRC_RIGHT: 2}          # row_meta3 · col_meta3 자리
_SRC_NAME = {SRC_LEFT: "왼쪽", SRC_RIGHT: "오른쪽"}
_HISTORY_SIDE = {ROLE_LEFT: "a", ROLE_BASE: "m", ROLE_RIGHT: "b"}


def _rm_readonly(func, path, _exc):
    """rmtree 가 읽기 전용 파일에서 멈추지 않게 — 속성을 풀고 한 번 더."""
    try:
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
        func(path)
    except OSError:
        pass


def _stat(path: str):
    try:
        st = os.stat(path)
        return (st.st_mtime_ns, st.st_size)
    except OSError:
        return None


class Diff3View(QWidget):
    def __init__(self, status_bar, parent=None):
        super().__init__(parent)
        self.status = status_bar
        self._load_worker = None
        self._diff_worker = None
        self._sheet_worker = None
        self._merge_worker = None
        self._diff_token = 0
        self._sheet_token = 0
        self._load_ctx = None
        self._result_path = ""
        self._auto_on_load = False         # P4V 병합(결과 파일 있음) — 처음 여는 시트는 자동 병합 준비
        self._tmpdir = None
        self._template = ""                # 가운데 원본 사본 — 비교 기준이자 저장의 바탕
        self._out_stat = None              # 결과 파일의 (수정 시각, 크기) — 바깥 변경 감지
        self._raw = ([], [], [])
        # 키(행·열)는 2파일의 전역 설정을 물려받지 않는다 — 다른 표에서 B열을 키로 잡은 뒤 ID 가 A열인
        # 표를 P4V 로 병합하면 이름 열이 키가 돼 같은 ID 행이 둘 생겼고, 전역 키 행이 2행이면 첫 데이터
        # 행을 머리글로 써서 열이 밀린 채 저장됐다(검토 실측). 시트별로 기억한 키를 쓰고
        # (_apply_remembered_key), 없으면 1행·A열. 행 순서(-1)로도 시작하지 않는다.
        self._key_row = 0
        self._key_col = 0
        self._excluded_cols: set = set()
        self._col_match = None             # match_columns3 원본(키와 무관)
        self._col_meta3 = None
        self._matrix: list = []
        self._row_meta: list = []
        self._row_masks: list = []
        self._conf_masks: list = []
        # 결정. 세 모델에 **참조로** 넘어가므로 리셋은 언제나 .clear()/.update() 로.
        self._staged: dict = {}            # 이번에 고른 것(주황) {(r, c): SRC_LEFT|SRC_RIGHT}
        self._merged: dict = {}            # 저장에 들어간 것(파랑)
        self._undo: list = []              # [[(cell, 준비였던 출처, 병합이었던 출처), ...]]
        self._dirty = False                # 마지막 저장 뒤 결정이 바뀌었나
        self._diff_only = False
        self._sheet_names: list = []
        self._current_sheet = None
        self._sheet_cache: dict = {}       # 결정이 있는 시트를 떠날 때 그 상태 그대로
        self._visited: set = set()
        self._changed_sheets = None        # Sheet3DiffWorker 결과(모르면 None)
        self._pending_sheet_diff = None
        self._carry = None                 # 키를 바꿔 다시 비교할 때 되살릴 결정
        self._computing = False            # 읽기·비교 중 — 그동안 결정·저장을 받지 않는다
        self._save_in_flight = False       # 저장 시작 ~ 완료 처리(워커 스레드가 끝나기 전에 풀린다)
        self._session_after_save = False   # 저장 중에 미뤄 둔 다시 열기
        self._moved = (0, 0)               # 지금 매트릭스에서 왼쪽·오른쪽이 순서를 바꾼 행 수
        self._loading = False              # 시트를 읽는 중 — _raw 가 아직 이 시트 것이 아니다
        self._session = 0                  # 세 파일을 새로 열 때마다 +1 — 늦게 끝난 저장 결과 거르기
        self._key_base = self._key_col     # 키 열 = **가운데 파일의** 몇 번째 열(화면 열은 열 맞춤에 따라 달라진다)
        self._matrix_key = (self._key_col, self._key_row)   # 지금 매트릭스를 만든 키
        self._matrix_col_match = None      # 지금 매트릭스를 만든 열 맞춤
        # 자동 병합이 위험한 이유("" = 없음) — 지금 매트릭스 기준. 열을 위치로 맞췄는데 머리글이
        # 다르거나, 행 순서로 맞췄는데 행 수가 다르면 짝이 어긋날 수 있다.
        self._auto_risk = ""
        self._pending_risk = ""            # 계산 중인 비교의 위험(결과가 도착하면 확정)
        self._ambiguous: set = set()       # 내용으로 짝지은 행(키가 비었거나 겹침) — 자동 병합 제외
        self._session_paths = None         # 지금 비교 중인 (왼쪽, 가운데, 오른쪽) — 경로칸 글자와 별개
        self._sheet_in: dict = {}          # 시트 → (왼쪽, 가운데, 오른쪽 에 있나)
        self._open_conf: dict = {}         # 시트 → (남은 충돌, 고르지 않은 변경, 그중 병합 불가) 칸 수
        self._syncing_scroll = False
        self._syncing_selection = False
        self._syncing_split = False
        self._find_cache = None            # (열쇠, 매트릭스 ref, 칸 목록) — _find_cells

        self._build_ui()
        for p in self.panels.values():
            p.table.set_key_col(self._key_col)
            p.table.set_key_row(self._key_row)
        self.setStyleSheet(APP_QSS)
        for seq, slot in (("Ctrl+Z", self._undo_last),
                          ("Ctrl+D", self._toggle_diff_only_shortcut),
                          ("F5", self._on_refresh_shortcut),
                          ("Alt+Up", lambda: self._goto_btn(self.prev_diff_btn, -1, False)),
                          ("Alt+Down", lambda: self._goto_btn(self.next_diff_btn, +1, False)),
                          ("Alt+Shift+Up", lambda: self._goto_btn(self.prev_conf_btn, -1, True)),
                          ("Alt+Shift+Down", lambda: self._goto_btn(self.next_conf_btn, +1, True)),
                          ("Ctrl+F", self._focus_find),
                          ("Ctrl+S", self._save)):
            sc = QShortcut(QKeySequence(seq), self)
            sc.setContext(Qt.WidgetWithChildrenShortcut)
            sc.activated.connect(slot)

    # ── UI ─────────────────────────────────────────────────────────────────
    def _button(self, text, tip, slot=None, checkable=False):
        b = QPushButton(text)
        b.setFixedHeight(36)
        b.setFont(ui_font(10))
        b.setToolTip(tip)
        b.setEnabled(False)
        b.setFocusPolicy(Qt.NoFocus)     # 눌러도 키보드 포커스는 표에 남긴다(2파일과 같음)
        if checkable:
            b.setCheckable(True)
            b.setObjectName("toggle_btn")
        if slot is not None:
            b.clicked.connect(slot)
        return b

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # ── 툴바 1행 — 2파일 화면(DiffView)과 같은 버튼·같은 순서·같은 자리 ──
        # 변경점만 보기 · 새로고침 · 이전/다음 변경 · 찾기 · (범례). 저장도 2파일처럼 패널 경로칸 옆
        # 아이콘이다(아래). 두 화면을 오가도 손이 같은 자리로 가게 한다.
        bar = QHBoxLayout()
        self.diff_only_btn = self._button("변경점만 보기", "변경된 행만 표시 / 전체 표시 전환 (Ctrl+D)",
                                          checkable=True)
        self.diff_only_btn.toggled.connect(self._on_diff_only_toggled)
        self.refresh_btn = self._button("새로고침", "세 파일을 다시 읽어 비교합니다 (F5)",
                                        self._run_refresh)
        self.prev_diff_btn = self._button("◀ 이전 변경", "이전 변경 셀로 이동 (Alt+↑)",
                                          lambda: self._goto(-1, False))
        self.next_diff_btn = self._button("다음 변경 ▶", "다음 변경 셀로 이동 (Alt+↓)",
                                          lambda: self._goto(+1, False))
        for w in (self.diff_only_btn, self.refresh_btn, self.prev_diff_btn, self.next_diff_btn):
            bar.addWidget(w)
        bar.addSpacing(16)
        bar.addLayout(build_find_box(
            self, "찾을 내용 (Ctrl+F)", self._goto_find,
            "셀 값 검색 — Enter: 다음 찾기, Shift+Enter: 이전 찾기",
            with_word=True))
        bar.addStretch()
        add_legend(bar, [(lbl, DIFF_COLORS[key]) for lbl, key in (
            ("변경", "modified"), ("같게 바뀜", "both"), ("충돌", "conflict"),
            ("준비", "staged"), ("병합", "merged"))])
        root.addLayout(bar)

        # ── 병합 줄 — 3파일에만 있는 것(충돌 이동 · 자동 병합 · 저장할 곳 안내) ──
        # 1행에 함께 두면 2파일과 버튼 자리가 달라지고, 찾기 박스까지 넣은 툴바가 기본 창 폭(1400)을
        # 넘는다(실측: 최소 폭 약 1,600 — 두 줄이면 약 1,250). 충돌과 자동 병합은 기준(가운데)이
        # 있어야 성립하는 개념이라 2파일에는 대응하는 버튼이 없다.
        merge_bar = QHBoxLayout()
        self.prev_conf_btn = self._button("◀ 이전 충돌", "아직 고르지 않은 이전 충돌로 이동 (Alt+Shift+↑)",
                                          lambda: self._goto(-1, True))
        self.next_conf_btn = self._button("다음 충돌 ▶", "아직 고르지 않은 다음 충돌로 이동 (Alt+Shift+↓)",
                                          lambda: self._goto(+1, True))
        self.auto_btn = self._button(
            "자동 병합", "충돌이 없는 변경을 모두 가운데로 가져오도록 병합 준비합니다.\n"
            "충돌(양쪽이 서로 다르게 고친 칸)은 직접 고릅니다.", self._auto_merge)
        for w in (self.prev_conf_btn, self.next_conf_btn, self.auto_btn):
            merge_bar.addWidget(w)
        merge_bar.addSpacing(8)
        self.result_label = QLabel("")
        self.result_label.setFont(ui_font(9))
        self.result_label.setStyleSheet("color: #555;")
        self.result_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        # 긴 결과 경로가 창 최소 폭을 밀어내지 않게 — 모자라면 잘리고, 전체는 툴팁으로.
        self.result_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.result_label.setMinimumWidth(0)
        merge_bar.addWidget(self.result_label, 1)
        root.addLayout(merge_bar)

        splitter = QSplitter(Qt.Horizontal)
        self.panels = {}
        for role, label in ((ROLE_LEFT, "왼쪽"), (ROLE_BASE, "가운데 (기준)"),
                            (ROLE_RIGHT, "오른쪽")):
            p = FilePanel(label, role, model=Diff3TableModel(role),
                          history_side=_HISTORY_SIDE[role], drop_hint=False)
            # 화살표 = 값이 흐르는 쪽. Alt+→ 왼쪽 → 가운데, Alt+← 오른쪽 → 가운데(세 패널 공통).
            p.table.stage_choices = (("왼쪽 → 가운데", SRC_LEFT), ("오른쪽 → 가운데", SRC_RIGHT))
            self.panels[role] = p
            splitter.addWidget(p)
        self.panel_l = self.panels[ROLE_LEFT]
        self.panel_b = self.panels[ROLE_BASE]
        self.panel_r = self.panels[ROLE_RIGHT]
        # 저장 = 2파일과 같은 자리(패널 경로칸 옆 아이콘). 쓰는 곳은 언제나 가운데(결과)뿐이라
        # 가운데 패널의 버튼만 남긴다 — 양옆은 읽기만 하는 파일이다.
        self.panel_l.save_btn.hide()
        self.panel_r.save_btn.hide()
        self.save_btn = self.panel_b.save_btn
        self.save_btn.clicked.connect(lambda _=False: self._save())
        splitter.setSizes([600, 600, 600])
        root.addWidget(splitter, 1)

        self.sheet_tabs = SheetTabBar()
        self.sheet_tabs.setObjectName("sheet_tabs")
        self.sheet_tabs.setExpanding(False)
        self.sheet_tabs.setDrawBase(True)
        self.sheet_tabs.setElideMode(Qt.ElideNone)
        self.sheet_tabs.setUsesScrollButtons(True)
        self.sheet_tabs.setFont(ui_font(9))
        self.sheet_tabs.currentChanged.connect(self._on_sheet_tab_changed)
        self.sheet_tabs.hide()
        root.addWidget(self.sheet_tabs)

        # 미니맵 스크롤바 → 틀 고정 순서(2파일과 같음 — 스크롤 동기화 배선보다 먼저 교체)
        for p in self.panels.values():
            p.table.setVerticalScrollBar(MinimapScrollBar(Qt.Vertical, p.table))
            p.table.setHorizontalScrollBar(MinimapScrollBar(Qt.Horizontal, p.table))
        self._freeze = {role: FreezeController(p.table) for role, p in self.panels.items()}

        for src in self.panels.values():
            t = src.table
            t.horizontalScrollBar().valueChanged.connect(
                lambda v, s=t: self._mirror_scroll(s, "h", v))
            t.verticalScrollBar().valueChanged.connect(
                lambda v, s=t: self._mirror_scroll(s, "v", v))
            for dst in self.panels.values():
                if dst is not src:
                    t.column_resized.connect(dst.table.apply_column_width)
                    t.row_resized.connect(dst.table.apply_row_height)
            t.selectionModel().selectionChanged.connect(
                lambda *_, s=src: self._sync_selection(s))
            t.drag_selection_finished.connect(lambda s=src: self._sync_selection(s))
            src.v_split.splitterMoved.connect(lambda _p, _i, s=src: self._sync_splitter(s))

        for role, p in self.panels.items():
            p.table.stage_requested.connect(self._stage_selected)
            p.table.unstage_requested.connect(self._unstage_selected)
            p.table.key_col_changed.connect(self._on_key_col_changed)
            p.table.key_row_changed.connect(self._on_key_row_changed)
            p.table.columns_exclude_set.connect(self._on_columns_exclude_set)
            p.file_loaded.connect(lambda path, r=role: self._on_file_loaded(r, path))
            p.folder_loaded.connect(lambda _p: self.status.showMessage(
                "3파일 비교는 파일만 다룹니다 — 폴더는 '＋ 새 비교' 탭에 넣으세요."))
            p.status_message.connect(self.status.showMessage)

        self._update_result_label()
        self.status.showMessage(
            "왼쪽·가운데(기준)·오른쪽 파일을 넣으면 비교합니다.  "
            "| 칸 선택 → Alt+→(왼쪽 값) / Alt+←(오른쪽 값) → 가운데 패널의 저장")

    # ── 진입점 ─────────────────────────────────────────────────────────────
    def load3(self, left: str = "", base: str = "", right: str = "", result: str = "",
              record: bool = True):
        """MainWindow 진입점. result 가 있으면 P4V 병합 — 그 파일에 쓰고 가운데는 읽기만 한다."""
        for role, path in ((ROLE_LEFT, left), (ROLE_BASE, base), (ROLE_RIGHT, right)):
            if path:
                self.panels[role].set_path(path)
                if record:
                    add_path_history(_HISTORY_SIDE[role], path)
        self._result_path = result or ""
        self._auto_on_load = bool(result)
        self._update_result_label()
        self._start_session()

    def _text_paths(self) -> tuple:
        return tuple(self.panels[r].get_path() for r in (ROLE_LEFT, ROLE_BASE, ROLE_RIGHT))

    def paths(self) -> tuple:
        """지금 비교 중인 세 파일. 경로칸 글자를 고쳐 두기만(Enter 없이) 해도 저장이 다른 파일을
        쓰거나 읽지 않도록, 비교를 시작할 때의 경로를 쓴다."""
        return self._session_paths or self._text_paths()

    def result_path(self) -> str:
        return self._result_path

    def out_path(self) -> str:
        """저장할 파일 — 결과 파일, 없으면 **비교 중인** 가운데 파일(경로칸 글자가 아니라)."""
        return self._result_path or self.paths()[1]

    def tab_title(self) -> str:
        l, b, r = self.paths()
        p = b or l or r
        return f"{os.path.basename(p)} · 3파일" if p else "3파일 비교"

    def tab_icon(self):
        for p in self.paths():
            if p:
                return ext_tab_icon(os.path.splitext(p)[1].lower())
        return QIcon()

    def shutdown(self):
        for w in (self._load_worker, self._diff_worker, self._sheet_worker, self._merge_worker):
            if w is None:
                continue
            try:
                w.blockSignals(True)
                if w.isRunning():
                    w.wait()
            except RuntimeError:
                pass
        self._drop_template()

    def _update_result_label(self):
        if self._result_path:
            text = f"결과 파일: {self._result_path}   — 가운데 원본은 바꾸지 않습니다"
            tip = ("병합 결과 저장 (Ctrl+S)\n"
                   f"결과 파일에 씁니다 — 가운데 원본은 바꾸지 않습니다.\n{self._result_path}")
        else:
            text = "저장하면 가운데 파일에 씁니다"
            tip = "병합 결과 저장 (Ctrl+S) — 가운데 파일에 씁니다"
        self.result_label.setText(text)
        self.result_label.setToolTip(text)
        # 저장 아이콘은 가운데 패널에 있지만, P4V 병합이면 쓰는 곳이 가운데가 아니라 결과 파일이다.
        self.save_btn.setToolTip(tip)

    # ── 세션(세 파일) ──────────────────────────────────────────────────────
    def _on_file_loaded(self, role: str, path: str):
        add_path_history(_HISTORY_SIDE[role], path)
        self._start_session()

    def _on_refresh_shortcut(self):
        if self.refresh_btn.isEnabled():
            self._run_refresh()

    def _run_refresh(self):
        for p in self.panels.values():
            p.table._user_col_widths.clear()
            p.table._user_row_heights.clear()
        self._start_session()

    def _drop_template(self):
        if self._tmpdir:
            # 이 화면이 만든 폴더 하나만. 원본이 읽기 전용(P4V 체크아웃 전·%b 임시 파일)이어도
            # 사본은 쓰기 가능하게 만들지만, 혹시 남은 속성 때문에 폴더가 쌓이지 않게 한 번 더.
            shutil.rmtree(self._tmpdir, onerror=_rm_readonly)
        self._tmpdir = None
        self._template = ""

    def _forget_inflight(self):
        """진행 중인 읽기·비교 결과를 버리게 한다 — 시트·파일이 바뀐 뒤 늦게 도착한 옛 결과가
        지금 화면을 덮지 않도록(읽기는 ctx, 비교는 token 으로 거른다)."""
        self._diff_token += 1
        self._load_ctx = None
        self._computing = False
        self._loading = False
        # 키를 바꿔 다시 비교하던 중이었다면 그 결정 되살리기도 버린다 — 남겨 두면 다음에 도착하는
        # **다른 시트의** 비교에 같은 (파일 행, 파일 열) 좌표로 붙는다(검토 실측: 손대지 않은
        # 시트의 칸이 병합 준비돼 저장됨).
        self._carry = None

    def _busy(self) -> bool:
        return self._computing or self._saving_now()

    def _refuse_if_saving(self) -> bool:
        """저장이 도는 동안엔 결정·키·시트를 바꾸지 않는다 — 저장은 시작할 때의 결정으로 쓰는데,
        끝나면 그 사이에 바꾼 것까지 '병합됨'으로 표시해 파일과 화면이 갈라졌다(검토 실측)."""
        if self._saving_now():
            self.status.showMessage("저장하는 중입니다 — 끝난 뒤에 다시 하세요.")
            return True
        return False

    def _reset_session(self):
        self._forget_inflight()          # 날아오는 옛 결과는 버린다
        self._session += 1
        self._sheet_token += 1
        self._raw = ([], [], [])
        self._col_match = None
        self._col_meta3 = None
        self._matrix = []
        self._row_meta = []
        self._row_masks = []
        self._conf_masks = []
        self._staged.clear()
        self._merged.clear()
        self._undo.clear()
        self._dirty = False
        self._carry = None
        self._sheet_cache.clear()
        self._visited.clear()
        self._changed_sheets = None
        self._open_conf.clear()
        self._sheet_in.clear()
        self._auto_risk = ""
        self._pending_risk = ""
        self._ambiguous = set()
        self._session_paths = None
        # 새 파일 — 앞 표의 키를 버린다. 같은 탭에 다른 표를 넣었을 때 앞 표의 B열이 남아, 기억이 없는
        # 새 표가 A열 대신 B열로 맞춰져 같은 ID 두 행이 저장됐다(검토 실측). 기억은 _on_loaded 에서
        # 다시 적용된다(F5 도 같다 — 사용자가 바꾼 키는 이미 기억돼 있다).
        self._key_row = 0
        self._key_col = 0
        self._key_base = 0
        for p in self.panels.values():
            p.table.set_key_col(0)
            p.table.set_key_row(self._key_row)
        self._excluded_cols.clear()
        self._clear_row_filter()
        self._clear_freeze()
        for p in self.panels.values():
            p.table.set_excluded_cols(self._excluded_cols)
            p.table._safe_clear()
        self.sheet_tabs.clear_changed()
        self._update_minimap()
        self._set_buttons_enabled(False)

    def _start_session(self):
        if self._save_in_flight:
            # 사본(_template)을 지우기 전에 저장이 끝나야 한다. 그 자리에서 기다리면 남은 저장
            # 시간 동안 화면이 멈춘다(검토 실측) — 저장이 끝나면 다시 연다.
            self._session_after_save = True
            self.status.showMessage("저장하는 중입니다 — 끝나면 다시 비교합니다.")
            return
        self._session_after_save = False
        self._reset_session()
        l, b, r = self._text_paths()
        self.refresh_btn.setEnabled(bool(l or b or r))
        if not (l and b and r):
            self.status.showMessage("왼쪽·가운데(기준)·오른쪽 세 파일을 모두 넣으면 비교합니다.")
            return
        missing = [p for p in (l, b, r) if not os.path.isfile(p)]
        if missing:
            self.status.showMessage("파일을 찾을 수 없습니다: " + ", ".join(missing))
            return
        try:
            self._drop_template()
            self._tmpdir = tempfile.mkdtemp(prefix="em3_")
            self._template = os.path.join(self._tmpdir, "base" + os.path.splitext(b)[1])
            shutil.copyfile(b, self._template)     # 속성(읽기 전용)은 따라오지 않게
            os.chmod(self._template, stat.S_IWRITE | stat.S_IREAD)
        except OSError as e:
            self._drop_template()
            self._on_error(f"가운데 파일을 읽을 수 없습니다: {e}")
            return
        self._session_paths = (l, b, r)
        self._out_stat = _stat(self.out_path())
        self._rebuild_sheet_tabs()
        self._run_compare()

    def _rebuild_sheet_tabs(self):
        l, b, r = self.paths()
        ln, bn, rn = list_sheet_names(l), list_sheet_names(self._template), list_sheet_names(r)
        names = list(bn)
        for n in ln + rn:
            if n not in names:
                names.append(n)
        self._sheet_in = {n: (n in ln, n in bn, n in rn) for n in names}
        if self._current_sheet in names:
            idx = names.index(self._current_sheet)
        else:
            idx = 0
            remembered = load_last_sheet(b)
            if remembered in names:
                idx = names.index(remembered)
            else:
                idx = next((i for i, n in enumerate(names) if not n.startswith("!")), 0)
        self._sheet_names = names
        self._current_sheet = names[idx] if names else None
        blocked = self.sheet_tabs.blockSignals(True)
        try:
            while self.sheet_tabs.count():
                self.sheet_tabs.removeTab(0)
            for n in names:
                self.sheet_tabs.addTab(n)
            if names:
                self.sheet_tabs.setCurrentIndex(idx)
        finally:
            self.sheet_tabs.blockSignals(blocked)
        self.sheet_tabs.setVisible(len(names) > 1)
        self._pending_sheet_diff = list(names) if len(names) > 1 else None

    def _flush_sheet_diff(self):
        names, self._pending_sheet_diff = self._pending_sheet_diff, None
        if not names:
            return
        l, _b, r = self.paths()
        self._sheet_token += 1
        token = self._sheet_token
        w = Sheet3DiffWorker((l, self._template, r), names)
        w.done.connect(lambda changed, t=token, nm=names: self._on_sheet_diff_done(changed, t, nm))
        w.finished.connect(w.deleteLater)
        self._sheet_worker = w
        w.start(QThread.LowestPriority)

    def _on_sheet_diff_done(self, changed, token, names):
        if token != self._sheet_token:
            return
        self._changed_sheets = set(changed)
        self.sheet_tabs.set_changed_indices([i for i, n in enumerate(names) if n in changed])

    # ── 시트 ───────────────────────────────────────────────────────────────
    def _on_sheet_tab_changed(self, index: int):
        if not (0 <= index < len(self._sheet_names)):
            return
        name = self._sheet_names[index]
        if name == self._current_sheet:
            return
        if self._refuse_if_saving():
            blocked = self.sheet_tabs.blockSignals(True)
            try:
                if self._current_sheet in self._sheet_names:
                    self.sheet_tabs.setCurrentIndex(self._sheet_names.index(self._current_sheet))
            finally:
                self.sheet_tabs.blockSignals(blocked)
            return
        self._stash_sheet()
        self._forget_inflight()
        self._current_sheet = name
        save_last_sheet(self.panel_b.get_path(), name)
        snap = self._sheet_cache.pop(name, None)
        if snap is not None:
            self._restore_sheet(snap)
        else:
            self._run_compare()

    def _stash_sheet(self):
        """결정이 있는 시트는 떠날 때 상태째 둔다 — 저장은 모든 시트의 결정을 함께 쓴다."""
        name = self._current_sheet
        if not self._matrix or not (self._staged or self._merged):
            self._sheet_cache.pop(name, None)
            return
        # 키는 **지금 매트릭스를 만든 키**로 둔다. 키를 바꿔 다시 비교하는 중에 떠나면
        # _key_col 은 이미 새 키인데 매트릭스·결정은 옛 키 것이다(검토 실측).
        self._sheet_cache[name] = {
            "raw": self._raw, "key": self._matrix_key,
            "excluded": set(self._excluded_cols), "col_match": self._matrix_col_match,
            "col_meta3": self._col_meta3, "matrix": self._matrix,
            "row_meta": self._row_meta, "row_masks": self._row_masks,
            "risk": self._auto_risk, "ambiguous": self._ambiguous, "moved": self._moved,
            "conf_masks": self._conf_masks,
            "staged": dict(self._staged), "merged": dict(self._merged),
        }

    def _restore_sheet(self, snap):
        self._raw = snap["raw"]
        self._key_col, self._key_row = snap["key"]
        self._matrix_key = snap["key"]
        self._matrix_col_match = snap["col_match"]
        # 자동 병합 위험·내용으로 짝지은 행도 그 매트릭스 것으로 — 안 되돌리면 막아 둔 자동
        # 병합이 어긋난 열로 돌았다(검토 실측).
        self._auto_risk = snap.get("risk", "")
        self._ambiguous = snap.get("ambiguous", set())
        self._moved = snap.get("moved", (0, 0))
        # 키 기준 열(가운데 파일의 몇 번째 열)도 이 시트 것으로 — 열을 위치로 맞춘 시트(cm 없음)에서 앞
        # 시트 값이 남으면, 키 행만 바꿔도 키 열이 앞 시트의 열 번호로 옮겨지고 그 틀린 키가 이 시트의
        # 기억으로 저장됐다(검토 실측).
        cm = self._matrix_col_match
        kc = self._key_col
        if cm and 0 <= kc < len(cm) and cm[kc][1] is not None:
            self._key_base = cm[kc][1]
        else:
            self._key_base = max(0, kc)
        self._excluded_cols.clear()
        self._excluded_cols.update(snap["excluded"])
        self._col_match = snap["col_match"]
        self._col_meta3 = snap["col_meta3"]
        self._matrix = snap["matrix"]
        self._row_meta = snap["row_meta"]
        self._row_masks = snap["row_masks"]
        self._conf_masks = snap["conf_masks"]
        self._staged.clear()
        self._staged.update(snap["staged"])
        self._merged.clear()
        self._merged.update(snap["merged"])
        self._undo.clear()
        for p in self.panels.values():
            p.table.set_key_col(self._key_col)
            p.table.set_key_row(self._key_row)
            p.table.set_excluded_cols(self._excluded_cols)
        self._refresh_tables()
        self._set_buttons_enabled(True)
        self.status.showMessage(self._summary())

    def _sheets_to_write(self) -> list:
        out = []
        if self._matrix:
            out.append((self._current_sheet, self._matrix, self._row_meta, self._col_meta3,
                        {**self._merged, **self._staged}))
        for name, s in self._sheet_cache.items():
            if name != self._current_sheet:
                out.append((name, s["matrix"], s["row_meta"], s["col_meta3"],
                            {**s["merged"], **s["staged"]}))
        return out

    # ── 비교 ───────────────────────────────────────────────────────────────
    def _run_compare(self):
        l, _b, r = self.paths()
        if not self._template:
            return
        self._set_buttons_enabled(False)
        self._matrix = []
        self._row_meta = []
        self._row_masks = []
        self._conf_masks = []
        self._staged.clear()
        self._merged.clear()
        self._undo.clear()
        self._excluded_cols.clear()
        # 옛 시트의 원본·열 맞춤도 비운다 — 읽는 동안 키를 바꾸면 옛 시트 데이터로 다시 비교해
        # 새 시트 탭 아래에 옛 시트를 '비교 완료'로 보여 줬다(검토 실측).
        self._raw = ([], [], [])
        self._col_match = None
        for p in self.panels.values():
            p.table.set_excluded_cols(self._excluded_cols)
        self._computing = True
        self._loading = True
        self.status.showMessage("파일 로딩 중...")
        paths = (l, self._template, r)
        ctx = (paths, self._current_sheet)
        self._load_ctx = ctx
        w = Load3Worker(paths, self._current_sheet)
        w.done.connect(lambda a, b, c, x=ctx: self._on_loaded(a, b, c, x))
        w.error.connect(self._on_error)
        w.progress.connect(self.status.showMessage)
        w.finished.connect(w.deleteLater)
        self._load_worker = w
        w.start()

    def _on_loaded(self, left, base, right, ctx):
        if ctx != self._load_ctx:
            return                         # 그 사이 파일·시트가 바뀌었다
        self._loading = False
        self._raw = (left, base, right)
        ncols = max((len(row) for d in self._raw for row in d), default=0)
        nrows = max(len(d) for d in self._raw)
        if self._key_col >= 0 and ncols and self._key_col >= ncols:
            self._key_col = 0
            for p in self.panels.values():
                p.table.set_key_col(0)
        if self._key_row > 0 and nrows and self._key_row >= nrows:
            self._key_row = 0
            for p in self.panels.values():
                p.table.set_key_row(0)
        self._apply_remembered_key(nrows)
        self._col_match = match_columns3(left, base, right, self._key_row)
        self._follow_key_column()
        self._start_diff("load")

    def _key_memory_paths(self) -> list:
        """키를 기억해 두는 파일 — 결과(P4V %r 은 작업 공간의 실제 파일), 가운데, 왼쪽, 오른쪽 순."""
        l, b, r = self.paths()
        return [p for p in (self._result_path, b, l, r) if p]

    def _apply_remembered_key(self, nrows: int):
        """기억해 둔 키를 적용한다 — **시트별 기억**이 먼저, 없으면 A열.

        파일 단위 기억(2파일이 남긴 것)은 시트를 가리지 않는다. 한 시트의 키 이름(ItemID)이 다른 시트에
        외래 키 열로 있으면 그 열이 키가 됐고, 이름이 없으면 번호(B)로 따라가 엉뚱한 열이 키가 됐다 —
        3파일은 P4V 첫 방문에 자동 병합을 걸어 같은 ID 두 행이 저장됐다(검토 실측). 그래서 파일 단위
        기억은 **시트가 하나뿐인 파일**에서만 쓴다(그때는 어느 시트 것인지 헷갈릴 일이 없다).
        시트 안에서는 이름이 먼저다 — 열이 끼어 자리가 바뀌어도 따라간다. 이름을 못 찾으면 번호."""
        sheet = self._current_sheet
        mem = None
        for p in self._key_memory_paths():
            mem = load_sheet_key(p, sheet)
            if mem is not None:
                break
        if mem is None and len(self._sheet_names) <= 1:
            for p in self._key_memory_paths():
                kk = load_last_key(p)
                if kk is not None:
                    mem = (kk[0], kk[1], load_last_key_name(p) or "")
                    break
        if mem is None:
            # 기억이 없는 시트는 1행·A열로 — 앞 시트의 키 행이 남으면 첫 데이터 행을 머리글로 써서 열을
            # 맞췄고, 한쪽이 열 하나만 끼워도 P4V 자동 병합이 열을 한 칸씩 밀어 썼다(검토 실측). 앞 시트에서
            # 고른 '행 순서'도 넘어오지 않게 한다.
            mem = (0, 0, "")
        kr, kc, name = mem
        kr = kr if 0 <= kr < max(1, nrows) else 0
        if kr != self._key_row:
            self._key_row = kr
            for pn in self.panels.values():
                pn.table.set_key_row(kr)
        want_col = -1 if kc < 0 else 0     # 실제 화면 열은 _follow_key_column 이 가운데 열로 맞춘다
        if want_col != self._key_col:
            self._key_col = want_col
            for pn in self.panels.values():
                pn.table.set_key_col(want_col)
        base = self._raw[1]
        head = [str(v).strip() for v in base[self._key_row]] if self._key_row < len(base) else []
        if name and name in head:
            self._key_base = head.index(name)
        else:
            self._key_base = max(0, kc)

    def _remember_key(self):
        """사용자가 고른 키를 결과·가운데 파일의 **이 시트**에 기억한다(이름 + 가운데 파일 열 번호).
        2파일의 파일 단위 기억은 건드리지 않는다 — 3파일에서 고른 키가 2파일 동작을 바꾸지 않게."""
        base = self._raw[1]
        kr, kb = self._key_row, self._key_base
        name = ""
        if self._key_col >= 0 and kr < len(base) and 0 <= kb < len(base[kr]):
            name = str(base[kr][kb]).strip()
        l, b, r = self.paths()
        for p in {self._result_path, b} - {""}:
            save_sheet_key(p, self._current_sheet, kr, kb if self._key_col >= 0 else -1, name)

    def _follow_key_column(self):
        """키 열을 **가운데 파일의 같은 열**로 맞춘다. 한쪽이 키 열 앞에 열을 끼우면 화면 열 번호가
        밀려, 화면 0열이 한쪽에만 있는 새 열이 된다 — 그대로 두면 위치 기준으로 떨어져 파일마다
        다른 열을 키로 써서 자동 병합이 진짜 행을 지운다(검토 실측). 2파일의 '키 이름 기억'과
        같은 일을 가운데 파일 열 번호로 한다."""
        cm = self._col_match
        if self._key_col < 0:
            return
        if not cm:
            # 열을 위치로 맞추는 표(머리글 가운데 빈 칸 등) — 화면 열이 곧 가운데 파일 열이다. 여기서
            # 그냥 돌아가면 기억한 키(B열)를 무시하고 A열로 비교했다(검토 실측).
            width = max((len(r) for r in self._raw[1]), default=0)
            if 0 <= self._key_base < width and self._key_base != self._key_col:
                self._key_col = self._key_base
                for p in self.panels.values():
                    p.table.set_key_col(self._key_base)
            return
        want = next((c for c, t in enumerate(cm) if t[1] == self._key_base), None)
        if want is not None and want != self._key_col:
            self._key_col = want
            for p in self.panels.values():
                p.table.set_key_col(want)

    def _start_diff(self, mode: str):
        self._computing = True
        self.status.showMessage("비교 계산 중...")
        self.status.begin_busy()
        self._diff_token += 1
        cm3 = usable_col_meta3(self._col_match, self._key_col)
        self._pending_risk = self._risk_of(cm3)
        w = Diff3Worker(*self._raw, self._key_col, self._key_row, self._excluded_cols,
                        self._diff_token, mode, cm3)
        w.done.connect(self._on_diff_ready)
        w.error.connect(self._on_error)
        w.finished.connect(w.deleteLater)
        self._diff_worker = w
        w.start()

    def _risk_of(self, cm3) -> str:
        """이 비교로 자동 병합을 해도 되나 — 안 되면 그 이유.

        열을 위치로 맞추면(이름으로 못 맞췄거나 키 열이 한쪽에 없음) 한쪽이 열 하나만 끼워도 오른쪽
        열이 모두 밀린 채 '한쪽만 바뀜'이 된다. 행 순서로 맞추면 한쪽이 행 하나만 끼워도 아래가
        모두 밀린다. 그때 자동 병합은 값을 엉뚱한 열·행에 쓴다(검토 실측) — 사람이 보고 고르게 한다.
        """
        kr = self._key_row
        # 그 시트가 아예 없는 파일은 뺀다 — 그쪽은 어차피 병합 출처로 쓰지 않는다(_src_blocked).
        present = [d for d, here in zip(self._raw, self._presence()) if here and d]
        if cm3 is None:
            heads = set()
            for d in present:
                h = list(d[kr]) if kr < len(d) else []
                while h and h[-1] == "":
                    h.pop()
                heads.add(tuple(h))
            if len(heads) > 1:
                return "열을 위치로 맞췄는데 세 파일의 머리글이 다릅니다"
        if self._key_col < 0:
            # 행 수가 같아도 한쪽이 행을 하나 끼우고 하나 지우면 그 사이가 모두 밀린다 — 다른 시트에서
            # 고른 행 순서가 이어져 P4V 자동 병합이 다른 행에 썼다(검토 실측). 행 순서는 사람이 본다.
            return "키 열 없이 행 순서로 맞추고 있습니다"
        return ""

    def _on_diff_ready(self, token, matrix, row_meta, col_meta, masks, cmasks, _changed,
                       dropped, mode, ambiguous=None):
        if token != self._diff_token:
            return
        self._computing = False
        self.status.end_progress()
        self._matrix, self._row_meta, self._col_meta3 = matrix, row_meta, col_meta
        self._row_masks, self._conf_masks = masks, cmasks
        self._auto_risk = self._pending_risk       # 이 매트릭스의 위험 — 결과가 도착해야 확정
        self._ambiguous = set(ambiguous or ())
        self._moved = self._count_moved()
        self._matrix_key = (self._key_col, self._key_row)
        self._matrix_col_match = self._col_match
        note = self._apply_carry() if self._carry is not None else ""
        self._refresh_tables()
        note += self._sheet_note()
        if mode == "load":
            first = self._current_sheet not in self._visited
            self._visited.add(self._current_sheet)
            if self._auto_on_load and first:
                if self._auto_risk:
                    note += f"  | {self._auto_risk} — 자동 병합 준비 안 함"
                else:
                    n, skipped = self._stage_plan(
                        auto_merge_plan(matrix, row_meta, self._ambiguous), undo=False)
                    if n:
                        note += f"  | 충돌 없는 변경 {n}칸을 자동으로 병합 준비함"
                    if skipped:
                        note += f"  | 병합할 수 없는 칸 {skipped}개"
            if self.diff_only_btn.isChecked():
                self._apply_diff_filter()
            else:
                self.diff_only_btn.setChecked(True)
        self._set_buttons_enabled(True)
        warn = (f"  | ⚠ 키가 겹치거나 빈 행 {dropped}개는 내용으로 짝지음 — 자동 병합에서 뺐으니 "
                "직접 확인하세요" if dropped else "")
        if self._auto_risk and not (mode == "load" and self._auto_on_load):
            warn += f"  | ⚠ {self._auto_risk} — 자동 병합 안 함"
        self.status.showMessage(self._summary() + warn + note)
        self._flush_sheet_diff()
        self._focus_grid()

    def _count_moved(self) -> tuple:
        """왼쪽·오른쪽이 **순서만 바꾼** 행 수 — 결과는 가운데 순서를 따르므로 그 이동은 들어가지 않는다.
        세 파일 모두에 있는 행을 본다 — 내용으로 짝지은 행(같은 키 묶음)도 넣는다. 묶음 안에서는 내용이
        같고 양쪽에 한 번씩만 있는 행만 순서와 상관없이 짝짓고(정렬·맞바꿈) 나머지는 순서를 지켜 짝짓는다.
        그래서 묶음 안의 정렬과 묶음끼리 자리를 바꾼 것(묶음 합치기)이 잡힌다 — 빼면 그런 정렬이 화면에도
        질문에도 안 나온 채 사라졌다(검토 실측). 가운데 순서로 늘어놓고, 각 쪽
        행 번호의 가장 긴 증가 부분열 밖에 있는 행 수."""
        import bisect
        rows = sorted((m for r, m in enumerate(self._row_meta)
                       if None not in m and r > self._key_row),
                      key=lambda m: m[1])
        out = []
        for i in (0, 2):
            tails = []
            for m in rows:
                k = bisect.bisect_left(tails, m[i])
                if k == len(tails):
                    tails.append(m[i])
                else:
                    tails[k] = m[i]
            out.append(len(rows) - len(tails))
        return tuple(out)

    def _count_open(self) -> tuple:
        """저장 확인용 — (고르지 않은 충돌, 고르지 않은 변경, 그중 어느 쪽에서도 병합할 수 없는 칸).
        제외 열도 센다(결과에선 그 칸도 가운데 값으로 남는다)."""
        conf = other = stuck = 0
        # 새 행(가운데에 없는 행)은 결과에서 시트 끝에 붙는다 — 화면에서 그 뒤에 가운데 행이 있으면
        # 제자리가 아니다(검토 실측: theirs 가 중간에 넣은 행이 끝으로 감).
        last_b = max((r for r, m in enumerate(self._row_meta) if m[1] is not None), default=-1)
        misplaced = sum(1 for r in range(min(last_b, len(self._row_meta)))
                        if self._row_meta[r][1] is None
                        and any(self._decision_of((r, c)) for c in range(len(self._matrix[r]))))
        for r, m in enumerate(self._row_masks):
            c = 0
            while m:
                if m & 1 and (r, c) not in self._staged and (r, c) not in self._merged:
                    if self._matrix[r][c][0] == S3_CONFLICT:
                        conf += 1
                    else:
                        other += 1
                        if self._auto_src(r, c) is None:
                            stuck += 1
                m >>= 1
                c += 1
        return conf, other, stuck, misplaced, sum(self._moved)

    def _count_conflicts(self, keep: int):
        conf = open_conf = 0
        for r, m in enumerate(self._conf_masks):
            m &= keep
            c = 0
            while m:
                if m & 1:
                    conf += 1
                    if (r, c) not in self._staged and (r, c) not in self._merged:
                        open_conf += 1
                m >>= 1
                c += 1
        return conf, open_conf

    def _summary(self) -> str:
        rows = len(self._matrix)
        cols = len(self._matrix[0]) if self._matrix else 0
        keep = self._keep_mask()
        changed = count_changed_masked(self._row_masks, keep)
        conf, open_conf = self._count_conflicts(keep)
        if self._matrix:
            self._open_conf[self._current_sheet] = self._count_open()
        return (f"비교 완료 — {rows}행 × {cols}열 | 변경 {changed}칸 · 충돌 {conf}칸"
                f"(남은 {open_conf}) · 준비 {len(self._staged)}칸 · 병합 {len(self._merged)}칸")

    def _presence(self) -> tuple:
        """지금 시트가 (왼쪽, 가운데, 오른쪽) 파일에 있나. 시트가 없는 형식(CSV 등)은 모두 참."""
        return self._sheet_in.get(self._current_sheet, (True, True, True))

    def _sheet_note(self) -> str:
        pl, pb, pr = self._presence()
        if not pb:
            return "  | 가운데 파일에 없는 시트 — 비교만 합니다(병합 안 됨)"
        gone = [n for n, here in (("왼쪽", pl), ("오른쪽", pr)) if not here]
        if gone:
            return (f"  | {'·'.join(gone)} 파일에 이 시트가 없음 — 시트 삭제는 병합하지 않습니다"
                    "(가운데 내용 유지)")
        return ""

    def _src_blocked(self, src) -> str:
        """이 시트에서 src 쪽 값을 가져올 수 없는 이유. 가져올 수 있으면 ""."""
        pl, pb, pr = self._presence()
        if not pb:
            return "가운데 파일에 이 시트가 없어 병합할 수 없습니다 — 비교만 합니다."
        if not (pl if src == SRC_LEFT else pr):
            return (f"{_SRC_NAME[src]} 파일에 이 시트가 없습니다 — 시트를 통째로 지우는 병합은 "
                    "하지 않습니다(가운데 내용을 유지합니다).")
        return ""

    # ── 키 · 제외 ──────────────────────────────────────────────────────────
    def _col_ids(self) -> list:
        """화면 열 → (왼쪽, 가운데, 오른쪽 파일 열). 위치 기준이면 (c, c, c) — 이름 기준과
        같은 꼴이어야 키를 바꿔 위치 기준으로 넘어가도 같은 열을 같은 열로 알아본다."""
        cm = self._col_meta3
        if cm:
            return [tuple(t) for t in cm]
        return [(c, c, c) for c in range(len(self._matrix[0]) if self._matrix else 0)]

    def _cell_ident(self, r, c, col_ids):
        return (tuple(self._row_meta[r]), col_ids[c] if c < len(col_ids) else None)

    def _capture_carry(self):
        if not self._matrix:
            return None
        ids = self._col_ids()
        return {"staged": [(self._cell_ident(r, c, ids), d) for (r, c), d in self._staged.items()],
                "merged": [(self._cell_ident(r, c, ids), d) for (r, c), d in self._merged.items()]}

    def _apply_carry(self) -> str:
        """키를 바꿔 다시 비교한 뒤 결정을 (파일 행 번호, 파일 열 번호)로 다시 찾아 건다.
        새 키로 행이 다르게 짝지어져 못 찾은 결정은 풀고 알린다 — 병합이 풀리면 다음 저장
        결과에서도 빠지므로 저장 대기로 둔다."""
        carry, self._carry = self._carry, None
        rows = {tuple(m): r for r, m in enumerate(self._row_meta)}
        cols = {t: c for c, t in enumerate(self._col_ids())}
        lost_merged = lost = 0
        for kind, target in (("staged", self._staged), ("merged", self._merged)):
            target.clear()
            for (rid, cid), d in carry[kind]:
                r, c = rows.get(rid), cols.get(cid)
                if r is None or c is None or self._matrix[r][c][0] == S3_SAME:
                    lost += 1
                    lost_merged += kind == "merged"
                    continue
                target[(r, c)] = d
        if lost_merged:
            self._dirty = True
        # 되살린 결정이 행의 일부만 덮으면(못 찾은 열이 있으면) 행째로 맞춘다 — 삭제 행이 반쪽만
        # 걸리면 키 없는 행이 저장됐다(검토 실측).
        rows = {k[0] for k in list(self._staged) + list(self._merged)}
        filled = self._complete_rows(rows)
        for k, d in filled.items():
            self._staged[k] = d
            self._merged.pop(k, None)
        if filled:
            self._dirty = True
        if not lost:
            return f"  | 행을 맞추려 {len(filled)}칸을 함께 준비" if filled else ""
        return (f"  | 새 키로 다시 찾지 못한 결정 {lost}개는 풀림"
                + (" — 병합이 풀린 칸은 다음 저장 결과에서도 빠집니다" if lost_merged else ""))

    def _refuse_if_loading(self) -> bool:
        if self._loading:
            self.status.showMessage("시트를 읽는 중입니다 — 다 읽은 뒤에 키를 바꾸세요.")
            return True
        return False

    def _on_key_col_changed(self, col: int):
        if col == self._key_col or self._refuse_if_saving() or self._refuse_if_loading():
            return
        if self._carry is None:          # 이미 다시 비교 중이면 처음 잡은 것을 그대로 쓴다
            self._carry = self._capture_carry()
        self._key_col = col
        cm = self._col_match
        self._key_base = (cm[col][1] if (cm and 0 <= col < len(cm) and cm[col][1] is not None)
                          else col)
        self._remember_key()
        self._excluded_cols.clear()        # 위치 기준으로 바뀌면 같은 번호가 다른 열이 된다
        for p in self.panels.values():
            p.table.set_key_col(col)
            p.table.set_excluded_cols(self._excluded_cols)
        self._undo.clear()
        self._set_buttons_enabled(False)
        # 첫 비교가 아직 안 끝났으면(읽은 직후) 그 비교를 대신한다 — 'load' 로 돌려야 처음 여는
        # 시트의 자동 병합 준비·필터가 걸린다.
        self._start_diff("recompute" if self._matrix else "load")

    def _on_key_row_changed(self, row: int):
        if row == self._key_row or self._refuse_if_saving() or self._refuse_if_loading():
            return
        if self._carry is None:
            self._carry = self._capture_carry()
        self._key_row = row
        self._excluded_cols.clear()
        for p in self.panels.values():
            p.table.set_key_row(row)
            p.table.set_excluded_cols(self._excluded_cols)
        self._col_match = match_columns3(*self._raw, row)
        self._follow_key_column()
        self._remember_key()
        self._undo.clear()
        self._set_buttons_enabled(False)
        # 첫 비교가 아직 안 끝났으면(읽은 직후) 그 비교를 대신한다 — 'load' 로 돌려야 처음 여는
        # 시트의 자동 병합 준비·필터가 걸린다.
        self._start_diff("recompute" if self._matrix else "load")

    def _on_columns_exclude_set(self, cols: list, exclude: bool):
        if exclude:
            self._excluded_cols.update(c for c in cols if c != self._key_col)
        else:
            self._excluded_cols.difference_update(cols)
        for p in self.panels.values():
            p.table.set_excluded_cols(self._excluded_cols)
        self._silent_clear_selection()
        self._apply_diff_filter()
        excl = ", ".join(get_column_letter(c + 1) for c in sorted(self._excluded_cols))
        self.status.showMessage(f"검사 제외 열: {excl or '없음'}  |  {self._summary()}")

    # ── 표 갱신 · 필터 ─────────────────────────────────────────────────────
    def _refresh_tables(self):
        for p in self.panels.values():
            p.populate(self._matrix, self._merged, self._staged, self._row_meta,
                       self._excluded_cols, self._col_meta3)
        self._apply_diff_filter()

    def _notify(self, cells):
        for p in self.panels.values():
            p.table.model().notify_cells(cells)

    def _keep_mask(self) -> int:
        return keep_mask(len(self._matrix[0]) if self._matrix else 0, self._excluded_cols)

    def _toggle_diff_only_shortcut(self):
        if self.diff_only_btn.isEnabled():
            self.diff_only_btn.setChecked(not self.diff_only_btn.isChecked())

    def _on_diff_only_toggled(self, checked: bool):
        self._diff_only = checked
        self._apply_diff_filter()

    def _apply_diff_filter(self):
        """2파일 _apply_diff_filter 와 같은 규칙. 결정이 걸린 칸은 언제나 바뀐 칸이라
        (저장해도 비교 기준이 그대로) 병합한 행이 필터에 숨는 일이 없다."""
        if not self._matrix:
            self._update_minimap()
            self._refresh_freeze()
            return
        desired = set(range(min(self._key_row + 1, len(self._matrix))))
        if self._diff_only:
            keep = self._keep_mask()
            for r, m in enumerate(self._row_masks):
                if r > self._key_row and not (m & keep):
                    desired.add(r)
        flipped = {}
        for role, p in self.panels.items():
            tbl = p.table
            changed = []
            flipped[role] = changed
            tbl._applying_sizes = True
            prev_upd = tbl.updatesEnabled()
            tbl.setUpdatesEnabled(False)
            prev_mode = tbl.verticalScrollMode()
            tbl.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)   # 대량 숨김 O(R²) 방지
            try:
                for r in range(tbl.rowCount()):
                    want = r in desired
                    if tbl.isRowHidden(r) != want:
                        tbl.setRowHidden(r, want)
                        changed.append(r)
            finally:
                tbl.setVerticalScrollMode(prev_mode)
                tbl._applying_sizes = False
                tbl.setUpdatesEnabled(prev_upd)
        self._update_minimap()
        self._refresh_freeze(flipped)

    def _clear_row_filter(self):
        for p in self.panels.values():
            tbl = p.table
            tbl._applying_sizes = True
            prev_mode = tbl.verticalScrollMode()
            tbl.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
            try:
                for r in range(tbl.rowCount()):
                    if tbl.isRowHidden(r):
                        tbl.setRowHidden(r, False)
            finally:
                tbl.setVerticalScrollMode(prev_mode)
                tbl._applying_sizes = False

    def _refresh_freeze(self, changed_rows=None):
        for role, fc in getattr(self, "_freeze", {}).items():
            fc.refresh(None if changed_rows is None else changed_rows.get(role, ()))

    def _clear_freeze(self):
        for fc in getattr(self, "_freeze", {}).values():
            fc.clear()

    def _update_minimap(self):
        row_ratios, col_ratios = [], []
        if self._matrix:
            keep = self._keep_mask()
            hidden = self.panel_b.table.isRowHidden
            visible = [r for r in range(len(self._matrix)) if not hidden(r)]
            n = len(visible)
            colmask = 0
            for vi, r in enumerate(visible):
                m = self._row_masks[r] & keep
                if m:
                    colmask |= m
                    row_ratios.append(0.0 if n == 1 else vi / (n - 1))
            total = len(self._matrix[0])
            c = 0
            while colmask:
                if colmask & 1:
                    col_ratios.append(0.0 if total == 1 else c / (total - 1))
                colmask >>= 1
                c += 1
        for p in self.panels.values():
            v, h = p.table.verticalScrollBar(), p.table.horizontalScrollBar()
            if isinstance(v, MinimapScrollBar):
                v.set_change_ratios(row_ratios)
            if isinstance(h, MinimapScrollBar):
                h.set_change_ratios(col_ratios)

    # ── 동기화 ─────────────────────────────────────────────────────────────
    def _mirror_scroll(self, src, orient: str, value: int):
        if self._syncing_scroll:
            return
        self._syncing_scroll = True
        try:
            for p in self.panels.values():
                if p.table is not src:
                    bar = (p.table.horizontalScrollBar() if orient == "h"
                           else p.table.verticalScrollBar())
                    bar.setValue(value)
        finally:
            self._syncing_scroll = False

    def _sync_selection(self, src_panel):
        src = src_panel.table
        if self._syncing_selection or src._populating or getattr(src, "_drag_selecting", False):
            return
        self._syncing_selection = True
        try:
            for p in self.panels.values():
                if p is not src_panel and not p.table._populating:
                    p.table.mirror_selection_from(src)
                    p._sync_cell_edit()
        finally:
            self._syncing_selection = False

    def _sync_splitter(self, src_panel):
        if self._syncing_split:
            return
        self._syncing_split = True
        try:
            for p in self.panels.values():
                if p is not src_panel:
                    p.v_split.setSizes(src_panel.v_split.sizes())
        finally:
            self._syncing_split = False

    def _silent_clear_selection(self):
        for p in self.panels.values():
            tbl = p.table
            tbl._populating = True
            try:
                tbl.clearSelection()
                sm = tbl.selectionModel()
                if sm is not None:
                    sm.clearCurrentIndex()
            finally:
                tbl._populating = False
            p._selected_cell = None
            p.cell_edit.clear()
            p.cell_edit.setEnabled(False)

    def _focus_grid(self):
        fw = QApplication.focusWidget()
        if fw is not None:
            if fw is self.find_edit:
                return                   # 찾을 내용을 치는 중 — 포커스를 뺏지 않는다(2파일과 같음)
            for p in self.panels.values():
                for w in (p.cell_edit, p.path_edit):
                    if w is fw or w.isAncestorOf(fw):
                        return
        self.panel_b.table.setFocus()

    # ── 찾기(2파일과 같은 동작) ────────────────────────────────────────────
    def _focus_find(self):
        """Ctrl+F — 찾기 입력란 포커스 + 전체 선택."""
        if self.find_edit.isEnabled():
            self.find_edit.setFocus()
            self.find_edit.selectAll()

    def _iter_find_matches(self, match):
        """검색어와 맞는 (r, c) 를 행 우선 순서로. **화면에 보이는 글자**로 판정한다 — 세 패널 중
        하나라도 맞으면 그 칸이고, 숨은 행·검사 제외 열은 뺀다(2파일과 같은 규칙).

        가운데 패널은 결정이 걸린 칸에서 고른 쪽 값을 보인다(Diff3TableModel.display_text). 그 칸의
        기준값은 화면 어디에도 없으므로 판정에서 뺀다 — 넣으면 '보이지 않는 글자를 찾는' 오동작이 된다.
        """
        excl = self._excluded_cols
        staged, merged = self._staged, self._merged
        decided = bool(staged or merged)
        hidden = self.panel_b.table.isRowHidden
        for r, row in enumerate(self._matrix):
            if hidden(r):
                continue
            for c, cell in enumerate(row):
                if c in excl:
                    continue
                lv, bv, rv = cell[1], cell[2], cell[3]
                if match(lv) or (rv != lv and match(rv)):
                    yield (r, c)
                elif bv != lv and bv != rv and match(bv) \
                        and not (decided and ((r, c) in staged or (r, c) in merged)):
                    yield (r, c)

    def _find_cells(self, term: str) -> list:
        """검색어에 맞는 칸 목록 — 같은 조건으로 다시 물으면 표를 다시 훑지 않는다(2파일
        DiffView._find_cells 와 같은 이유: Enter 를 누를 때마다 표 전체를 훑으면 큰 표에서 굼뜨다).
        결과를 바꾸는 것은 검색어·옵션·표·필터·제외 열·결정이 걸린 칸이다. 표는 identity 로 대조한다
        (새 비교·시트 전환·키 변경이면 다른 객체)."""
        case, word = self.find_case_btn.isChecked(), self.find_word_btn.isChecked()
        key = (term, case, word, self._diff_only, tuple(sorted(self._excluded_cols)),
               len(self._staged), len(self._merged),
               hash(frozenset(self._staged)) if self._staged else 0,
               hash(frozenset(self._merged)) if self._merged else 0)
        hit = self._find_cache
        if hit is not None and hit[0] == key and hit[1] is self._matrix:
            return hit[2]
        cells = list(self._iter_find_matches(make_find_matcher(term, case, word)))
        self._find_cache = (key, self._matrix, cells)
        return cells

    def _goto_find(self, direction: int):
        """direction=+1: 다음 찾기, -1: 이전 찾기. 끝에 닿으면 반대편에서 이어 찾는다."""
        term = self.find_edit.text()
        if not term or not self._matrix:
            return
        cells = self._find_cells(term)
        if not cells:
            self.status.showMessage(f'"{term}" — 일치 항목이 없습니다.')
            return
        anchor = self._current_anchor()
        wrapped = ""
        if direction > 0:
            target = next((x for x in cells if x > anchor), None)
            if target is None:
                target = cells[0]
                wrapped = " — 처음부터 다시 검색합니다."
        else:
            target = next((x for x in reversed(cells) if x < anchor), None)
            if target is None:
                target = cells[-1]
                wrapped = " — 끝에서부터 다시 검색합니다."
        self._land_on(*target)
        idx = cells.index(target) + 1
        self.status.showMessage(f'찾기: "{term}" {idx}/{len(cells)}개 일치{wrapped}')

    def _current_anchor(self) -> tuple:
        """이동(변경·충돌·찾기)의 출발 칸 — 가운데, 왼쪽, 오른쪽 순으로 현재 칸이 있는 패널."""
        for p in (self.panel_b, self.panel_l, self.panel_r):
            r, c = p.table._current_cell()
            if r >= 0 and c >= 0:
                return (r, c)
        return (0, -1)

    def _land_on(self, r: int, c: int):
        """세 패널을 같은 칸으로 옮기고 화면 가운데로 스크롤."""
        for p in self.panels.values():
            p.table._move_current_cell(r, c)
            if p.table.model().is_data_cell(r, c):
                p.table.scrollTo(p.table.model().index(r, c), QAbstractItemView.PositionAtCenter)

    # ── 이동 ───────────────────────────────────────────────────────────────
    def _goto_btn(self, btn, direction, conflicts):
        if btn.isEnabled():
            self._goto(direction, conflicts)

    def _goto(self, direction: int, conflicts: bool):
        if not self._matrix:
            return
        masks = self._conf_masks if conflicts else self._row_masks
        keep = self._keep_mask()
        hidden = self.panel_b.table.isRowHidden
        cells = []
        for r, m in enumerate(masks):
            m &= keep
            if not m or hidden(r):
                continue
            c = 0
            while m:
                if m & 1 and not (conflicts and ((r, c) in self._staged
                                                 or (r, c) in self._merged)):
                    cells.append((r, c))
                m >>= 1
                c += 1
        what = "남은 충돌" if conflicts else "변경 칸"
        if not cells:
            self.status.showMessage(f"{what}이 없습니다." if conflicts else "변경된 칸이 없습니다.")
            return
        anchor = self._current_anchor()
        if direction > 0:
            target = next((x for x in cells if x > anchor), None)
        else:
            target = next((x for x in reversed(cells) if x < anchor), None)
        if target is None:
            self.status.showMessage(f"{'마지막' if direction > 0 else '첫'} {what}입니다.")
            return
        self._land_on(*target)
        self.panel_b.table.setFocus()

    # ── 병합 준비 ──────────────────────────────────────────────────────────
    def _selected_cells(self) -> set:
        out = set()
        for p in self.panels.values():
            out |= p.table.get_selected_cells()
        n = len(self._matrix)
        return {(r, c) for (r, c) in out if 0 <= r < n and 0 <= c < len(self._matrix[r])}

    def _row_has(self, r, src) -> bool:
        return r < len(self._row_meta) and self._row_meta[r][_SRC_IDX[src]] is not None

    def _unmergeable(self, cell, src) -> bool:
        """가운데에 쓸 자리가 없거나(그 열이 가운데에 없음), 쓰면 멀쩡한 열을 비우게 되는 칸
        (소스에 그 열이 없는데 행은 있음). 2파일 _unmergeable_cells · build_side_patches 와 같은 규칙."""
        cm = self._col_meta3
        if not cm:
            return False
        r, c = cell
        if not (0 <= c < len(cm)):
            return True
        src_col, base_col = cm[c][_SRC_IDX[src]], cm[c][1]
        if base_col is None:
            return True
        return src_col is None and self._row_has(r, src)

    def _noop(self, cell, src) -> bool:
        """가져와도 가운데가 그대로인 칸 — 그쪽 값이 기준과 같다(예: 왼쪽만 바뀐 칸에 '오른쪽 →').
        충돌 칸은 아니다 — 기준값과 같은 쪽을 고르는 것도 충돌을 푸는 결정이다(행 단위 충돌에서
        남긴 쪽의 안 바뀐 칸을 고를 수 없어 충돌이 영영 남던 문제)."""
        r, c = cell
        st, lv, bv, rv = self._matrix[r][c]
        if st == S3_SAME:
            return True
        if st == S3_CONFLICT:
            return False
        val = lv if src == SRC_LEFT else rv
        removal = (not self._row_has(r, src)) and self._row_meta[r][1] is not None
        return val == bv and not removal

    def _decision_of(self, cell):
        return self._staged.get(cell) or self._merged.get(cell)

    def _auto_src(self, r, c):
        """충돌 없는 칸을 채울 출처 — auto_merge_plan 과 같은 규칙. '양쪽이 같게' 칸은 값이 같으니
        가져올 수 있는 쪽으로(한쪽 파일에 시트·열이 없으면 다른 쪽에서)."""
        st = self._matrix[r][c][0]
        cands = {S3_LEFT: (SRC_LEFT,), S3_RIGHT: (SRC_RIGHT,),
                 S3_BOTH: (SRC_LEFT, SRC_RIGHT)}.get(st, ())
        for src in cands:
            if not self._src_blocked(src) and not self._unmergeable((r, c), src):
                return src
        return None

    def _complete_rows(self, rows) -> dict:
        """행째로만 의미가 있는 결정을 행 전체로 맞춘다 → 새로 걸 {(r, c): 출처}.

        · 한쪽이 **지운 행**을 그쪽으로 골랐으면, 그 행의 바뀐 칸 전체를 같은 쪽으로(일부만이면
          키 없는 반쪽 행이 남는다 — 키를 바꿔 결정을 되살릴 때 열 일부를 못 찾아 생겼다).
        · **가운데에 없는 행**(새 행)에 결정이 하나라도 있으면, 아직 안 고른 충돌 없는 칸을 자동
          병합 규칙으로 채운다 — 충돌 칸 하나만 골라 저장하면 키 없는 행이 들어갔다(검토 실측).
        """
        add = {}
        for r in rows:
            row = self._matrix[r]
            meta = self._row_meta[r]
            decided = {c: self._decision_of((r, c)) for c in range(len(row))}
            decided = {c: d for c, d in decided.items() if d}
            if not decided:
                continue
            if meta[1] is not None:
                for src, i in _SRC_IDX.items():
                    if meta[i] is None and src in decided.values():
                        for c, cell in enumerate(row):
                            if cell[0] != S3_SAME and decided.get(c) != src \
                                    and not self._unmergeable((r, c), src):
                                add[(r, c)] = src
            else:
                for c, cell in enumerate(row):
                    if c in decided or cell[0] in (S3_SAME, S3_CONFLICT):
                        continue
                    src = self._auto_src(r, c)
                    if src:
                        add[(r, c)] = src
        return add

    def _row_cells(self, rows) -> set:
        return {(r, c) for r in rows for c in range(len(self._matrix[r]))}

    def _record_undo(self, cells):
        self._undo.append([(cell, self._staged.get(cell), self._merged.get(cell))
                           for cell in cells])

    def _stage_plan(self, plan: dict, undo: bool = True):
        """{(r, c): 출처} 중 아직 결정이 없고 병합할 수 있는 칸을 준비. (준비한 수, 못 하는 수)."""
        cells = {k: v for k, v in plan.items()
                 if k not in self._staged and k not in self._merged}
        bad = set()
        for k, v in list(cells.items()):
            if self._src_blocked(v) or self._unmergeable(k, v):
                alt = self._auto_src(*k)     # '양쪽이 같게' 는 다른 쪽에서라도
                if alt is None:
                    bad.add(k)
                    del cells[k]
                else:
                    cells[k] = alt
        if not cells:
            return 0, len(bad)
        if undo:
            self._record_undo(cells)
        self._staged.update(cells)
        self._dirty = True
        self._notify(cells)
        self._set_buttons_enabled(True)
        return len(cells), len(bad)

    def _removal_rows(self, rows, src) -> set:
        """rows 중 src 쪽이 지운 행(그쪽엔 없고 가운데엔 있음)."""
        return {r for r in rows if not self._row_has(r, src) and self._row_meta[r][1] is not None}

    def _stage_selected(self, src: str):
        if not self._matrix or src not in _SRC_IDX or self._computing or self._refuse_if_saving():
            return
        why = self._src_blocked(src)
        if why:
            QMessageBox.information(self, "알림", why)
            return
        sel = self._selected_cells()
        changed = {k for k in sel if self._matrix[k[0]][k[1]][0] != S3_SAME}
        if not changed:
            QMessageBox.information(self, "알림", "선택한 칸 중 바뀐 칸이 없습니다.")
            return
        cells = expand_row_removals(self._matrix, self._row_meta, changed, src)
        noop = {k for k in cells if self._noop(k, src)}
        bad = {k for k in cells - noop if self._unmergeable(k, src)}
        cells -= noop | bad
        if not cells:
            why = (f"고른 칸은 {_SRC_NAME[src]} 값이 가운데와 같아 가져올 것이 없습니다."
                   if noop and not bad else
                   "고른 칸이 모두 가운데에 없는 열이거나 열 구성이 달라 병합할 수 없습니다.")
            QMessageBox.information(self, "알림", why)
            return
        # 이 행을 **반대쪽이 지운 것**으로 골라 두었다면, 이쪽 값을 고르는 순간 그 행은 남기는
        # 것이다 — 반대쪽 삭제 결정을 행째 푼다. 남겨 두면 몇 칸만 비워진 키 없는 행이 저장됐다
        # (검토 실측). 지운 쪽을 고를 때는 expand_row_removals 가 이미 행 전체를 덮는다.
        other = SRC_RIGHT if src == SRC_LEFT else SRC_LEFT
        released = set()
        for r in self._removal_rows({k[0] for k in cells}, other):
            released |= {(r, c) for c in range(len(self._matrix[r]))
                         if (self._staged.get((r, c)) or self._merged.get((r, c))) == other
                         and (r, c) not in cells}
        rows = {k[0] for k in cells}
        self._record_undo(self._row_cells(rows))     # 행째 되돌릴 수 있게
        for k in released:
            self._staged.pop(k, None)
            self._merged.pop(k, None)
        for k in cells:
            self._staged[k] = src
            self._merged.pop(k, None)
        filled = self._complete_rows(rows)
        for k, d in filled.items():
            self._staged[k] = d
            self._merged.pop(k, None)
        self._dirty = True
        self._notify(self._row_cells(rows))
        self._silent_clear_selection()
        self._set_buttons_enabled(True)
        note = ""
        if released:
            note += (f"  | 같은 행의 {_SRC_NAME[other]} 삭제 결정 {len(released)}칸은 풀림"
                     "(행을 남김)")
        if filled:
            note += f"  | 새 행의 나머지 {len(filled)}칸도 함께 준비(키 없는 행 방지)"
        if noop:
            note += f"  | {_SRC_NAME[src]} 값이 가운데와 같은 {len(noop)}칸은 뺌"
        if bad:
            note += f"  | 열 구성이 달라 병합할 수 없는 {len(bad)}칸은 뺌"
        self.status.showMessage(
            f"{_SRC_NAME[src]} → 가운데 병합 준비 {len(cells)}칸{note}  | {self._summary()}")

    def _unstage_selected(self):
        if self._computing or self._refuse_if_saving():
            return
        sel = self._selected_cells()
        cells = {k for k in sel if k in self._staged or k in self._merged}
        if not cells:
            QMessageBox.information(self, "알림", "선택한 칸 중 병합 준비·병합된 칸이 없습니다.")
            return
        # 지운 행을 가져온 결정은 행 전체로 — 일부만 풀면 반쪽짜리 빈 행이 남는다.
        # 새 행(가운데에 없는 행)도 행째로 — 키 칸만 풀면 키 없는 행이 들어간다(검토 실측).
        for src in _SRC_IDX:
            rows = [k for k in cells
                    if self._decision_of(k) == src
                    and not self._row_has(k[0], src) and self._row_meta[k[0]][1] is not None]
            for r in {k[0] for k in rows}:
                cells |= {(r, c) for c in range(len(self._matrix[r]))
                          if self._decision_of((r, c)) == src}
        for r in {k[0] for k in cells if self._row_meta[k[0]][1] is None}:
            cells |= {(r, c) for c in range(len(self._matrix[r])) if self._decision_of((r, c))}
        had_merged = sum(1 for k in cells if k in self._merged)
        self._record_undo(cells)
        for k in cells:
            self._staged.pop(k, None)
            self._merged.pop(k, None)
        self._dirty = True
        self._notify(cells)
        self._silent_clear_selection()
        self._set_buttons_enabled(True)
        note = (f"  | 병합을 푼 {had_merged}칸은 다음 저장 때 결과에서도 빠집니다"
                if had_merged else "")
        self.status.showMessage(f"병합 준비 취소 {len(cells)}칸{note}  | {self._summary()}")

    def _auto_merge(self):
        if not self._matrix or self._computing or self._refuse_if_saving():
            return
        if self._auto_risk:
            QMessageBox.information(
                self, "알림", f"{self._auto_risk} — 짝이 어긋날 수 있어 자동 병합을 하지 않습니다.\n"
                "세 파일에 모두 있는 열을 키 열로 정하거나, 칸을 보고 직접 고르세요.")
            return
        n, bad = self._stage_plan(auto_merge_plan(self._matrix, self._row_meta, self._ambiguous))
        note = f"  | 병합할 수 없는 칸 {bad}개" if bad else ""
        if not n:
            self.status.showMessage("자동으로 가져올 변경이 더 없습니다." + note)
            return
        self.status.showMessage(f"자동 병합 준비 {n}칸{note}  | {self._summary()}")

    def _undo_last(self):
        if self._computing or self._refuse_if_saving():
            return
        if not self._undo:
            self.status.showMessage("되돌릴 병합 준비 동작이 없습니다.")
            return
        entry = self._undo.pop()
        for cell, staged, merged in entry:
            self._staged.pop(cell, None)
            self._merged.pop(cell, None)
            if staged is not None:
                self._staged[cell] = staged
            if merged is not None:
                self._merged[cell] = merged
        self._dirty = True
        self._notify([cell for cell, _s, _m in entry])
        self._silent_clear_selection()
        self._set_buttons_enabled(True)
        self.status.showMessage(f"되돌리기 — {len(entry)}칸  | {self._summary()}")

    # ── 저장 ───────────────────────────────────────────────────────────────
    def _saving_now(self) -> bool:
        if self._save_in_flight:
            return True
        w = self._merge_worker
        if w is None:
            return False
        try:
            return w.isRunning()
        except RuntimeError:
            self._merge_worker = None
            return False

    def _unvisited_changed_sheets(self):
        """P4V 결과 파일 모드에서, 바뀐 곳이 있는데 열어 보지 않은 시트. 아직 모르면 None.
        가운데에 없는 시트는 어차피 병합할 수 없으니 뺀다."""
        if len(self._sheet_names) <= 1:
            return []
        if self._changed_sheets is None:
            return None
        return [n for n in self._sheet_names
                if n in self._changed_sheets and n not in self._visited
                and self._sheet_in.get(n, (True, True, True))[1]]

    def _left_out_report(self) -> list:
        """P4V 결과에 **들어가지 않는 것** — 사람이 읽을 줄 목록. 결과는 가운데 원본에서 다시
        만들므로, 고르지 않은 칸은 모두 가운데 값으로 남는다(충돌뿐 아니라 한쪽만 바꾼 칸, 가운데에
        없는 열처럼 병합할 수 없는 칸도). 예전엔 충돌만 물어서 나머지가 소리 없이 빠졌다(검토 실측)."""
        if self._matrix:
            self._open_conf[self._current_sheet] = self._count_open()
        lines = [f"[{n}] {why}" for n, why in self._sheet_structure_changes()]
        for n, (conf, other, stuck, misplaced, moved) in self._open_conf.items():
            if not self._sheet_in.get(n, (True, True, True))[1]:
                continue                 # 가운데에 없는 시트 — 위의 시트 줄로 알린다
            name = n if n is not None else "시트"
            if conf:
                lines.append(f"[{name}] 고르지 않은 충돌 {conf}칸 — 가운데 값으로 남음")
            if other:
                extra = f"(그중 병합할 수 없는 칸 {stuck}개 — 가운데에 없는 열 등)" if stuck else ""
                lines.append(f"[{name}] 고르지 않은 변경 {other}칸{extra} — 가운데 값으로 남음")
            if misplaced:
                lines.append(f"[{name}] 새 행 {misplaced}개는 화면 자리가 아니라 시트 끝에 붙습니다")
            if moved:
                lines.append(f"[{name}] 왼쪽·오른쪽이 바꾼 행 순서({moved}개 행)는 들어가지 않습니다"
                             "(가운데 순서 유지)")
        return lines

    def _sheet_structure_changes(self) -> list:
        """세 파일 중 일부에만 있는 시트 — [(시트, 설명)]. 시트를 통째로 더하거나 지우는 병합은
        하지 않으므로, P4V 결과에서는 이 시트들이 소리 없이 빠지거나 남는다(검토 실측)."""
        out = []
        for n in self._sheet_names:
            pl, pb, pr = self._sheet_in.get(n, (True, True, True))
            if (pl, pb, pr) == (True, True, True):
                continue
            if not pb:
                who = "·".join(w for w, here in (("왼쪽", pl), ("오른쪽", pr)) if here)
                out.append((n, f"{who}에서 새로 만든 시트 — 결과에 들어가지 않습니다"))
            else:
                who = "·".join(w for w, here in (("왼쪽", pl), ("오른쪽", pr)) if not here)
                out.append((n, f"{who}에서 지운 시트 — 결과에 가운데 내용이 남습니다"))
        return out

    def _save(self):
        if self._saving_now() or not self._matrix or self._computing:
            return
        if self._carry is not None:
            return                       # 키를 바꿔 다시 비교하는 중
        if not self._dirty:
            self.status.showMessage("저장할 변경이 없습니다.")
            return
        if self._session_paths and self._text_paths() != self._session_paths:
            QMessageBox.warning(
                self, "저장하지 않음",
                "경로칸이 지금 비교한 파일과 다릅니다 — 다른 파일에 쓰게 될 수 있어 저장하지 "
                "않았습니다.\n경로칸에서 Enter 로 다시 비교하거나 경로를 되돌린 뒤 저장하세요.")
            return
        l, b, r = self.paths()
        out = self.out_path()
        for p in (b, out):
            if os.path.splitext(p)[1].lower() not in _WRITABLE:
                QMessageBox.information(
                    self, "저장 미지원",
                    "3파일 병합 저장은 .xlsx/.xlsm 파일에서만 지원합니다.\n\n" + p)
                return
        if _is_file_locked(out):
            QMessageBox.warning(self, "저장할 수 없음", f"파일을 저장할 수 없습니다.\n\n{out}")
            return
        if _stat(out) != self._out_stat:
            QMessageBox.warning(
                self, "저장하지 않음",
                "저장할 파일이 이 화면 밖에서 바뀌었습니다 — 덮어쓰지 않았습니다.\n"
                "새로고침(F5)으로 다시 읽은 뒤 병합하세요.\n\n" + out)
            return
        if self._result_path:
            pending = self._unvisited_changed_sheets()
            if pending is None:
                self.status.showMessage("시트별 변경을 확인하는 중입니다 — 잠시 뒤 다시 저장하세요.")
                return
            if pending:
                QMessageBox.warning(
                    self, "열어 보지 않은 시트",
                    "아래 시트에도 바뀐 곳이 있습니다. 결과 파일은 가운데 원본에서 새로 만들므로, "
                    "열어 보지 않은 시트에는 양쪽 변경이 들어가지 않습니다.\n"
                    "시트를 열어 병합을 확인한 뒤 저장하세요.\n\n" + "\n".join(pending))
                return
            # P4V 는 이 결과를 '해결됨'으로 받는다 — 결과에 들어가지 않는 것을 모두 보여 주고 묻는다.
            left_out = self._left_out_report()
            if left_out:
                ans = QMessageBox.question(
                    self, "결과에 빠지는 변경",
                    "아래는 결과 파일에 들어가지 않습니다(가운데 원본 그대로 남습니다).\n\n"
                    + "\n".join(left_out) + "\n\n그래도 저장할까요?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if ans != QMessageBox.Yes:
                    self.status.showMessage(
                        "저장하지 않음 — Alt+Shift+↓ 로 남은 충돌로, Alt+↓ 로 변경 칸으로 갈 수 있습니다.")
                    return
        self._set_buttons_enabled(False)
        self.status.showMessage("저장 중...")
        self.status.begin_busy()
        w = Merge3Worker(self._template, out, self._sheets_to_write(),
                         {SRC_LEFT: l, SRC_RIGHT: r})
        w.done.connect(lambda n, o=out, ss=self._session: self._on_saved(n, o, ss))
        w.error.connect(self._on_save_error)
        w.finished.connect(w.deleteLater)
        self._merge_worker = w
        self._save_in_flight = True
        w.start()

    def _on_save_error(self, msg: str):
        self._save_in_flight = False
        self._on_error(msg)
        if self._session_after_save:
            self._start_session()

    def _on_saved(self, count: int, out: str, session=None):
        self._save_in_flight = False
        self.status.end_progress()
        clear_values_cache()
        if self._session_after_save:
            # 저장 중에 파일을 바꾸거나 새로고침했다 — 미뤄 둔 다시 열기를 지금 한다
            self._start_session()
            QMessageBox.information(self, "저장 완료", f"병합 결과를 저장했습니다.\n\n{out}")
            return
        if session is not None and session != self._session:
            return                       # 저장 뒤 파일을 새로 열었다 — 지금 화면과 무관
        moved = list(self._staged)
        self._merged.update(self._staged)
        self._staged.clear()
        for s in self._sheet_cache.values():
            s["merged"].update(s["staged"])
            s["staged"].clear()
        self._undo.clear()
        self._dirty = False
        self._out_stat = _stat(out)
        self._notify(moved)
        self._set_buttons_enabled(True)
        msg = f"저장 완료 — {os.path.basename(out)} | 병합 {len(self._merged)}칸"
        self.status.showMessage(f"{msg}  | {self._summary()}")
        QMessageBox.information(self, "저장 완료", f"병합 결과를 저장했습니다.\n\n{out}")

    def _on_error(self, msg: str):
        self._computing = False
        self.status.end_progress()
        self._set_buttons_enabled(True)
        self.status.showMessage(f"오류: {msg}")
        QMessageBox.critical(self, "오류", f"작업 실패:\n{msg}")

    def _set_buttons_enabled(self, enabled: bool):
        has = enabled and bool(self._matrix)
        for b in (self.diff_only_btn, self.prev_diff_btn, self.next_diff_btn,
                  self.prev_conf_btn, self.next_conf_btn, self.auto_btn,
                  self.find_edit, self.find_case_btn, self.find_word_btn,
                  self.find_prev_btn, self.find_next_btn):
            b.setEnabled(has)
        self.save_btn.setEnabled(has and self._dirty)
