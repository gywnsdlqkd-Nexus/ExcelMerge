"""UI 테마 — 색상/폰트/QSS/앱 아이콘 단일 출처 (excel_diff_merge.py에서 분리)."""
import base64 as _b64
import os
import sys
import functools

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QFont, QIcon, QImage, QPalette, QPixmap

FONT_FAMILY = "맑은 고딕"


def ui_font(size: int = 9, bold: bool = False) -> QFont:
    if bold:
        return QFont(FONT_FAMILY, size, QFont.Bold)
    return QFont(FONT_FAMILY, size)


def force_active_highlight(widget):
    """비활성/비포커스 상태에서도 선택 하이라이트를 활성(파랑)과 동일하게 고정한다.
    포커스가 다른 위젯(버튼/단축키)으로 옮겨가도 선택색이 회색으로 흐려지지 않게 한다.
    그리드/트리 뷰가 공통으로 쓰는 팔레트 오버라이드(중복 제거용 헬퍼)."""
    pal = widget.palette()
    for grp in (QPalette.Inactive, QPalette.Disabled):
        pal.setColor(grp, QPalette.Highlight,
                     pal.color(QPalette.Active, QPalette.Highlight))
        pal.setColor(grp, QPalette.HighlightedText,
                     pal.color(QPalette.Active, QPalette.HighlightedText))
    widget.setPalette(pal)


# diff 상태별 색상 원본 테이블 — 모든 파생 색 상수는 여기서만 나온다
DIFF_RGB = {
    "added":    (198, 239, 206),   # 연두   - 한쪽 파일에만 값이 있음 (A 전용/B 전용)
    "modified": (255, 235, 156),   # 노랑   - 양쪽 값이 서로 다름
    "staged":   (255, 185,  80),   # 주황   - 저장 대기 중
    "merged":   (173, 216, 230),   # 연파랑 - 병합 완료
    "same":     (255, 255, 255),   # 흰색   - 동일
}
DIFF_COLORS = {k: QColor(*v) for k, v in DIFF_RGB.items()}

# 변경점이 있는 하단 시트 탭 배경 강조(반투명 노랑) — 기본 탭 렌더 위에 덧칠해 텍스트 가독 유지.
SHEET_TAB_CHANGED_BG = QColor(255, 235, 156, 140)   # modified 노랑 계열, alpha

MINIMAP_MARKER_COLOR = QColor(255, 140, 0, 220)   # 진한 주황 — 범례의 주황(staged)보다 채도↑

# 셀값란/그리드에서 A/B 값이 다른 문자 구간 배경 강조색 (연핑크).
CELL_DIFF_HL = QColor(0xFF, 0xC9, 0xC9)   # #FFC9C9
# 강조 구간의 폰트 색 (빨강) — 선택 셀(흰 글자)에서도 핑크 배경 위 글자가 보이도록.
CELL_DIFF_FG = QColor(0xC6, 0x28, 0x28)   # #C62828
# 수식으로 계산된 값 셀의 폰트 색 (파랑) — 값이 수식 결과임을 그리드에서 구분.
# 변경(diff) 구간은 이 파랑 대신 CELL_DIFF_FG(빨강)로 덮어써 변경 위치를 우선 표시한다.
CELL_FORMULA_FG = QColor(0x15, 0x65, 0xC0)   # #1565C0
# 변경 검사에서 제외된 열의 셀 배경(회색) — 흰색과 구분되어 '검사 안 함'을 한눈에 표시.
EXCLUDED_CELL_BG = QColor(0xD9, 0xD9, 0xD9)   # #D9D9D9

# 폴더 비교(FolderCompareView) 파일 목록 상태별 배경색 — folder_compare 상태 키에 매핑.
# only_a/only_b(한쪽만) = 연두(added 재사용), modified = 노랑(재사용),
# same = 옅은 회색(중립), placeholder(반대쪽에 없는 칸) = 더 옅은 회색.
FOLDER_STATUS_RGB = {
    "same":     (245, 245, 245),   # 옅은 회색 - 동일
    "modified": DIFF_RGB["modified"],
    "only_a":   DIFF_RGB["added"],
    "only_b":   DIFF_RGB["added"],
}
FOLDER_STATUS_COLORS = {k: QColor(*v) for k, v in FOLDER_STATUS_RGB.items()}

# 신규(added) 셀의 반대쪽 빈 칸에 그리는 대각선 해치 색 (Beyond Compare식 매칭 표시).
HATCH_COLOR = QColor(0xB0, 0xB0, 0xB0)   # 회색

# 그리드 헤더(가로/세로) 배경·글자 색 — 키 열/행(노랑)·검사 제외(회색)·일반 구분.
HEADER_KEY_BG = QColor(255, 213, 0)      # 키 열/행 헤더 배경(노랑)
HEADER_EXCL_BG = QColor(220, 220, 220)   # 검사 제외 열 헤더 배경(회색)
HEADER_NORMAL_BG = QColor(232, 234, 240) # 일반 헤더 배경
HEADER_FG = QColor(0, 0, 0)              # 일반 헤더 글자(검정)
HEADER_EXCL_FG = QColor(140, 140, 140)   # 검사 제외 열 헤더 글자(회색)

MENU_QSS = f"QMenu {{ font-family: '{FONT_FAMILY}'; font-size: 9pt; }}"

APP_QSS = """
            QMainWindow { background: #f5f5f5; }
            QPushButton {
                background: #0078d4; color: white;
                border: none; border-radius: 4px; padding: 6px 14px;
            }
            QPushButton:hover   { background: #005fa3; }
            QPushButton:pressed { background: #004880; }
            QPushButton:disabled { background: #b0b0b0; color: #e0e0e0; }
            QPushButton#save_btn { background: #c47700; }
            QPushButton#save_btn:hover   { background: #a06000; }
            QPushButton#save_btn:pressed { background: #7d4a00; }
            QPushButton#save_btn:disabled { background: #b0b0b0; color: #e0e0e0; }
            QPushButton#toggle_btn { background: #555; }
            QPushButton#toggle_btn:checked { background: #0078d4; }
            QPushButton#toggle_btn:hover   { background: #333; }
            QPushButton#toggle_btn:disabled { background: #b0b0b0; color: #e0e0e0; }
            QPushButton#find_btn {
                background: #ffffff; border: 1px solid #c8c8c8; padding: 0;
            }
            QPushButton#find_btn:hover   { background: #eef5fc; border-color: #0078d4; }
            QPushButton#find_btn:pressed { background: #dcebf9; }
            QPushButton#find_btn:checked { background: #0078d4; border-color: #005fa3; }
            QPushButton#find_btn:checked:hover { background: #005fa3; }
            QPushButton#find_btn:disabled { background: #f2f2f2; border-color: #dedede; }
            QLineEdit#find_edit:focus { border: 1px solid #0078d4; }
            QLineEdit#find_edit:disabled { background: #f2f2f2; color: #aaa; }
            QLineEdit {
                border: 1px solid #ccc; border-radius: 4px;
                padding: 4px 8px; background: #fff;
            }
            CellEditWidget {
                border: 1px solid #ccc; border-radius: 4px;
                padding: 2px 6px; background: #fff;
            }
            CellEditWidget:disabled { background: #f5f5f5; color: #aaa; }
            QTableView { border: 1px solid #ddd; gridline-color: #e0e0e0; }
            QHeaderView::section {
                background: #e8eaf0; border: none;
                border-right: 1px solid #ccc; border-bottom: 1px solid #ccc;
                padding: 3px 6px; font-weight: bold;
            }
            QTabBar#sheet_tabs { background: #f0f0f0; }
            QTabBar#sheet_tabs::tab {
                background: #e4e6ec; color: #333;
                border: 1px solid #ccc; border-bottom: none;
                padding: 3px 14px; margin-right: 2px;
                min-width: 40px;
            }
            QTabBar#sheet_tabs::tab:selected {
                background: #fff; color: #0078d4;
                border-top: 2px solid #0078d4;
            }
            QTabBar#sheet_tabs::tab:hover:!selected { background: #eef5fc; }
        """

DROP_HIGHLIGHT_QSS = (
    "FilePanel { border: 2px dashed #0078d4; border-radius: 6px; background: #e8f4ff; }"
)


_APP_ICON_B64 = "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"

def load_app_icon():
    """앱 아이콘을 단일 소스(app_icon.ico)에서 로드.
    - PyInstaller 번들: sys._MEIPASS 임시 폴더의 ico 사용
    - 개발 모드: 스크립트와 같은 폴더의 ico 사용
    ICO는 멀티 사이즈(16/32/48/256 등)를 포함하므로 작업 표시줄·타이틀바 모두
    동일 이미지에서 적합한 사이즈를 가져온다. 실패 시 base64 PNG로 폴백.
    """
    for d in _icon_search_dirs():
        ico_path = os.path.join(d, "app_icon.ico")
        if os.path.isfile(ico_path):
            icon = QIcon(ico_path)
            if not icon.isNull():
                return icon
    # 폴백: 임베디드 base64 PNG (ico 누락 시)
    raw = _b64.b64decode(_APP_ICON_B64)
    pix = QPixmap()
    pix.loadFromData(raw, "PNG")
    return QIcon(pix)


# ── PNG 아이콘 로더 (헤더 키/제외 표시, 탭 파일형식 표시) ──────────────────────
def _icon_search_dirs():
    """PNG/ICO 아이콘 탐색 후보 디렉터리.
    이미지는 프로젝트 루트의 images/ 폴더에 있고, spec이 번들 시 _MEIPASS/images/로 넣는다.
    - 프리즈: sys._MEIPASS/images 및 _MEIPASS 루트
    - dev: (프로젝트 루트 | excelmerge/)의 images/ 및 각 폴더 자체(구버전 호환)
    각 후보의 images/ 하위를 먼저 본다.
    """
    bases = []
    mei = getattr(sys, "_MEIPASS", None)
    if mei:
        bases.append(mei)
    here = os.path.dirname(os.path.abspath(__file__))
    bases.append(here)
    bases.append(os.path.dirname(here))
    dirs = []
    for b in bases:
        dirs.append(os.path.join(b, "images"))
        dirs.append(b)
    return dirs


def _icon_with_transparent_white(path: str, thr: int = 240) -> QIcon:
    """PNG를 로드하되 (거의)흰색 배경 픽셀을 투명 처리한 QIcon 반환.
    원본이 큰 편(수백 px)이라 128px로 축소 후 픽셀 순회 — 헤더/메뉴 소형 표시에 충분하고
    순회 비용도 작다(1회 로드·캐시). 실패 시 null QIcon."""
    img = QImage(path)
    if img.isNull():
        return QIcon()
    img = img.convertToFormat(QImage.Format_ARGB32)
    if max(img.width(), img.height()) > 128:
        img = img.scaled(128, 128, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        img = img.convertToFormat(QImage.Format_ARGB32)
    w, h = img.width(), img.height()
    for y in range(h):
        for x in range(w):
            c = img.pixelColor(x, y)
            if c.red() >= thr and c.green() >= thr and c.blue() >= thr:
                c.setAlpha(0)
                img.setPixelColor(x, y, c)
    return QIcon(QPixmap.fromImage(img))


@functools.lru_cache(maxsize=None)
def load_png_icon(name: str, transparent_white: bool = False) -> QIcon:
    """파일명(예: 'Key.png')으로 QIcon을 1회 로드·캐시. 못 찾으면 null QIcon(폴백).
    transparent_white=True면 흰색 배경을 투명 처리해 채색 배경(헤더 노랑/회색) 위에서
    흰 사각형이 보이지 않도록 한다."""
    for d in _icon_search_dirs():
        p = os.path.join(d, name)
        if os.path.isfile(p):
            if transparent_white:
                ic = _icon_with_transparent_white(p)
            else:
                ic = QIcon(p)
            if not ic.isNull():
                return ic
    return QIcon()


def key_header_icon() -> QIcon:
    """키 열 헤더 표시 아이콘 (흰 배경 투명)."""
    return load_png_icon("Key.png", True)


def exclude_header_icon() -> QIcon:
    """변경 검사 제외 열 헤더/메뉴 표시 아이콘 (흰 배경 투명)."""
    return load_png_icon("Exception.png", True)


def reset_header_icon() -> QIcon:
    """초기화/해제(검사 제외 해제 등) 메뉴 표시 아이콘 (흰 배경 투명)."""
    return load_png_icon("Reset.png", True)


# 확장자 → 탭 파일형식 아이콘 파일명. (uasset은 요청대로 Excel.png 사용)
_TAB_ICON_FILES = {
    ".xlsx": "Excel.png", ".xls": "Excel.png", ".xlsm": "Excel.png", ".xlsb": "Excel.png",
    ".json": "JSON.png",
    ".uasset": "Excel.png",
}


def ext_tab_icon(ext: str) -> QIcon:
    """확장자에 해당하는 탭 아이콘. 매핑 없으면 null QIcon."""
    name = _TAB_ICON_FILES.get((ext or "").lower())
    return load_png_icon(name) if name else QIcon()
