"""비교 뷰(DiffView/FolderCompareView) 공용 툴바 조각 — 찾기 박스·범례.

두 뷰가 거의 동일한 찾기 박스(검색란 + 대소문자/전체단어 토글 + 이전/다음 버튼)와
색 범례를 각자 인라인으로 중복 구성하던 것을 여기로 통합했다. 순수 위젯 구성만 담당하며,
동작(검색 로직)은 각 뷰의 콜백(on_goto)·상태에 위임한다.
"""
from PyQt5.QtWidgets import QHBoxLayout, QLineEdit, QPushButton, QLabel, QShortcut
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QKeySequence

from .theme import ui_font
from .widgets import make_find_icon

_CASE_TIP = ("대소문자 무시 (Ignore case)\n"
             "켜짐: 대소문자를 구분하지 않고 검색\n"
             "꺼짐: 대소문자가 정확히 일치할 때만 검색")
_WORD_TIP = ("전체 단어 일치 (Match whole word only)\n"
             "켜짐: 검색어가 독립된 단어로 존재할 때만 찾음\n"
             "꺼짐: 부분 문자열도 찾음")


def build_find_box(owner, placeholder: str, on_goto, edit_tooltip: str,
                   with_word: bool = True) -> QHBoxLayout:
    """공용 찾기 박스를 만들고 위젯을 owner 속성으로 노출한다.

    owner.find_edit / find_case_btn / (with_word면) find_word_btn / find_prev_btn /
    find_next_btn 를 세팅하고, 검색란 Enter·Shift+Enter 및 이전/다음 버튼을 on_goto(±1)에 배선.
    반환: 툴바에 addLayout 할 QHBoxLayout.
    """
    box = QHBoxLayout()
    box.setSpacing(4)

    edit = QLineEdit()
    edit.setObjectName("find_edit")
    edit.setPlaceholderText(placeholder)
    edit.setFixedHeight(36)
    edit.setFixedWidth(200)
    edit.setFont(ui_font(10))
    edit.setClearButtonEnabled(True)
    edit.setEnabled(False)
    edit.setToolTip(edit_tooltip)
    edit.returnPressed.connect(lambda: on_goto(+1))
    sc = QShortcut(QKeySequence("Shift+Return"), edit)
    sc.setContext(Qt.WidgetShortcut)
    sc.activated.connect(lambda: on_goto(-1))
    box.addWidget(edit)
    owner.find_edit = edit
    owner._find_prev_shortcut = sc   # QShortcut 참조 유지(GC 방지)

    def _btn(kind: str, tip: str, checkable: bool) -> QPushButton:
        b = QPushButton()
        b.setObjectName("find_btn")
        b.setFixedSize(36, 36)
        b.setIcon(make_find_icon(kind))
        b.setIconSize(QSize(22, 22))
        b.setCheckable(checkable)
        b.setEnabled(False)
        b.setToolTip(tip)
        box.addWidget(b)
        return b

    owner.find_case_btn = _btn("case", _CASE_TIP, True)
    owner.find_case_btn.setChecked(True)
    if with_word:
        owner.find_word_btn = _btn("word", _WORD_TIP, True)
    owner.find_prev_btn = _btn("prev", "이전 찾기 (Shift+Enter)", False)
    owner.find_prev_btn.clicked.connect(lambda: on_goto(-1))
    owner.find_next_btn = _btn("next", "다음 찾기 (Enter)", False)
    owner.find_next_btn.clicked.connect(lambda: on_goto(+1))
    return box


def add_legend(toolbar: QHBoxLayout, entries) -> None:
    """(label, QColor) 목록을 색 스와치 + 라벨로 toolbar 에 추가."""
    for lbl, color in entries:
        dot = QLabel("  ")
        dot.setFixedSize(20, 20)
        dot.setStyleSheet(
            f"background:{color.name()}; border:1px solid #aaa; border-radius:3px;")
        txt = QLabel(lbl)
        txt.setFont(ui_font(9))
        toolbar.addWidget(dot)
        toolbar.addWidget(txt)
        toolbar.addSpacing(8)
