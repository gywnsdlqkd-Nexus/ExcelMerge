"""테이블/스크롤/입력 위젯 (excel_diff_merge.py에서 분리)."""
import os
import re

from PyQt5.QtWidgets import (
    QApplication, QTableView, QAbstractItemView, QLineEdit, QPlainTextEdit,
    QHeaderView, QMenu, QStyle, QScrollBar, QStyleOptionSlider,
    QStyledItemDelegate, QStyleOptionViewItem,
    QTabBar, QStylePainter, QStyleOptionTab,
)
from PyQt5.QtCore import (
    Qt, pyqtSignal, QItemSelection, QItemSelectionRange, QItemSelectionModel,
    QRect, QPoint, QSize, QObject, QEvent,
)
from PyQt5.QtGui import (
    QPainter, QPalette, QTextCursor, QTextCharFormat, QTextDocument, QTextOption,
    QIcon, QPixmap, QFont, QPen, QColor, QKeySequence,
)
from PyQt5 import sip
from .colref import get_column_letter

from .diff_model import DiffTableModel
from .loaders import _SUPPORTED_EXTS
from .constants import DIR_A2B, DIR_B2A
from .theme import (
    CELL_DIFF_HL, CELL_DIFF_FG, HATCH_COLOR, MENU_QSS, SHEET_TAB_CHANGED_BG,
    MINIMAP_MARKER_COLOR, ui_font, key_header_icon, exclude_header_icon,
    reset_header_icon, force_active_highlight,
    HEADER_NORMAL_BG, HEADER_TINT_ALPHA,
)



# 셀 표시 문자열의 '줄바꿈' 감지. QStyledItemDelegate.initStyleOption 은 DisplayRole 의
# '\n' 을 U+2028(LineSeparator)로 바꿔 넣으므로, opt.text 에는 '\n' 이 **없다**.
# '\n' 만 검사하면 멀티라인 셀까지 단일 라인 경로로 새어 들어가 한 줄로 뭉개진다.
_LINEBREAK_RE = re.compile("[\n\r\u2028\u2029]")


def draw_diagonal_hatch(painter, rect, color=HATCH_COLOR, step=6, width=1):
    """rect 안에 등간격 대각선(↗)을 그린다. clip은 이 함수가 rect로 건다.
    신규 셀의 반대쪽 빈 칸에 매칭 표시(Beyond Compare식)로 사용."""
    painter.save()
    painter.setClipRect(rect)
    painter.setPen(QPen(color, width))
    h = rect.height()
    x = rect.left() - h
    while x < rect.right():
        painter.drawLine(x, rect.bottom(), x + h, rect.top())
        x += step
    painter.restore()


def make_find_icon(kind: str) -> QIcon:
    """찾기 버튼용 아이콘을 QPainter로 렌더링 (HiDPI 2배 해상도).
    체크 시 배경이 파란색으로 바뀌므로 Off=진회색 / On=흰색 두 벌을 등록.
    (DiffView·FolderCompareView 공용)"""
    def render(color: QColor) -> QPixmap:
        s = 32
        pm = QPixmap(s * 2, s * 2)
        pm.setDevicePixelRatio(2)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.TextAntialiasing)
        if kind == "case":
            p.setPen(color)
            p.setFont(QFont("Segoe UI", 12, QFont.Bold))
            p.drawText(QRect(0, 0, s, s), Qt.AlignCenter, "Aa")
        elif kind == "word":
            p.setPen(color)
            p.setFont(QFont("Segoe UI", 10, QFont.Bold))
            p.drawText(QRect(0, 0, s, s - 8), Qt.AlignCenter, "ab")
            p.setPen(QPen(color, 1.8, Qt.SolidLine, Qt.RoundCap))
            y = s - 7
            p.drawLine(QPoint(7, y), QPoint(s - 7, y))
            p.drawLine(QPoint(7, y), QPoint(7, y - 4))
            p.drawLine(QPoint(s - 7, y), QPoint(s - 7, y - 4))
        else:  # "prev" / "next" — 셰브론 화살표
            p.setPen(QPen(color, 2.6, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
            m = s // 2
            d = -1 if kind == "prev" else 1
            p.drawLine(QPoint(m - 3 * d, m - 7), QPoint(m + 4 * d, m))
            p.drawLine(QPoint(m + 4 * d, m), QPoint(m - 3 * d, m + 7))
        p.end()
        return pm

    ic = QIcon()
    ic.addPixmap(render(QColor("#3b3b3b")), QIcon.Normal, QIcon.Off)
    ic.addPixmap(render(QColor("#ffffff")), QIcon.Normal, QIcon.On)
    ic.addPixmap(render(QColor("#b8b8b8")), QIcon.Disabled, QIcon.Off)
    ic.addPixmap(render(QColor("#b8b8b8")), QIcon.Disabled, QIcon.On)
    return ic


# 자동 컬럼 너비 상한 — 150px
# 데이터가 긴 셀 때문에 열이 화면을 가리지 않도록 제한.
# 사용자가 헤더 드래그로 직접 넓힌 열은 _user_col_widths 에 기록되어 이 상한 무시.
# 새로고침 시에는 _run_refresh()가 _user_col_widths를 비우므로 모든 열이 디폴트로 복귀.
MAX_AUTO_COL_WIDTH_PX = 150


class MinimapScrollBar(QScrollBar):
    """수직 또는 수평 스크롤바 위에 변경된 셀(행/열)의 위치를 색상 마커로 오버레이.
    paintEvent에서 super 호출 후, orientation에 맞춰 트랙(groove) 영역에
    비율 위치(0.0~1.0)별로 가는 막대를 그린다.

    마커는 **바깥쪽 가장자리의 얇은 띠**에만 그리고(세로바=오른쪽, 가로바=아래쪽),
    **핸들과 겹치는 구간은 건너뛴다**. 예전에는 트랙 폭의 대부분(14px 중 10px)을
    핸들 위에 덮어 그려서, 변경이 많은 파일에서는 스크롤바가 주황 벽이 되고 핸들이
    어디 있는지 보이지 않았다. 클릭·드래그 영역은 예전과 같다(칠하는 자리만 바뀐다).
    핸들 구간을 비워도 정보 손실이 적다 — 그 구간은 지금 화면에 떠 있는 부분이라
    변경 위치를 표에서 직접 볼 수 있다.
    """
    _MARKER_COLOR = MINIMAP_MARKER_COLOR
    _GUTTER_PX = 4      # 마커 띠 폭

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self._ratios: list = []   # 0.0~1.0 사이 변경 위치 목록
        # 비율 -> [(픽셀 오프셋, 겹친 개수)] 캐시. denom(트랙 길이)별로 다르다.
        self._px_denom: int = -1
        self._px_offsets: list | None = None
        self._stack_colors: dict = {}   # 겹친 개수 -> 합성색

    def set_change_ratios(self, ratios):
        # 변경된 경우에만 repaint (불필요한 페인트 방지)
        if list(ratios) != self._ratios:
            self._ratios = list(ratios)
            self._px_offsets = None   # 픽셀 오프셋 캐시 무효화
            self.update()

    def _pixel_offsets(self, denom: int) -> list:
        """비율 목록을 [(픽셀 오프셋, 겹친 개수)] 로 접는다.

        변경 행이 많으면 마커 수가 트랙 픽셀 수를 크게 넘는다(실측: 2,479개가
        고유 615px 에 4배 중복) — 같은 자리에 같은 사각형을 몇 번씩 그렸다.
        접어서 자리마다 한 번만 그리면 fillRect 호출이 그만큼 줄어든다.

        겹친 개수를 함께 들고 다니는 이유: 마커 색에 알파가 있어(220/255) 겹쳐
        그리면 점점 진해진다 — 그냥 중복만 버리면 변경이 밀집한 구간이 눈에 띄게
        연해진다. _stacked_color 가 그 누적을 한 번의 합성으로 재현한다.

        denom(트랙 길이)이 바뀌면 해상도가 달라지므로 다시 계산한다 — 창 크기나
        스플리터를 바꾼 뒤에도 마커 밀도가 유지된다(미리 접어 두면 트랙이 길어졌을 때
        마커가 실제보다 듬성해진다).
        """
        if self._px_offsets is None or self._px_denom != denom:
            counts: dict = {}
            for ratio in self._ratios:
                off = int(ratio * denom)
                counts[off] = counts.get(off, 0) + 1
            self._px_denom = denom
            self._px_offsets = list(counts.items())   # dict 는 삽입 순서 보존
        return self._px_offsets

    def _stacked_color(self, k: int) -> QColor:
        """같은 자리에 k 번 겹쳐 그렸을 때와 같은 색을 1회 합성으로 낸다.

        source-over 를 k 번 적용하면 배경 계수가 (1-a)^k 이므로
        alpha_eff = 1 - (1-a)^k 짜리 색으로 한 번 그리면 같은 결과가 된다.
        (8비트 반올림 때문에 채널당 최대 1 정도의 오차는 남는다.)
        """
        hit = self._stack_colors.get(k)
        if hit is None:
            base = self._MARKER_COLOR
            a = base.alpha() / 255.0
            eff = 1.0 - (1.0 - a) ** k
            hit = QColor(base.red(), base.green(), base.blue(),
                         max(0, min(255, round(eff * 255))))
            self._stack_colors[k] = hit
        return hit

    def _marker_band(self, groove) -> tuple:
        """마커 띠의 (시작 좌표, 두께) — 스크롤바 바깥쪽 가장자리에 1px 띄워 붙인다.

        세로바는 오른쪽, 가로바는 아래쪽 = 표에서 먼 쪽. 좁은 트랙에서도 최소 1px 은
        남기고, 띠가 트랙보다 두꺼워지지 않게 자른다.
        """
        g = max(1, min(self._GUTTER_PX, (groove.width() if self.orientation() == Qt.Vertical
                                         else groove.height()) - 1))
        edge = groove.right() if self.orientation() == Qt.Vertical else groove.bottom()
        return edge - g, g

    @staticmethod
    def _hits_slider(pos: int, slider, vertical: bool) -> bool:
        """두께 2px 짜리 마커가 핸들 구간과 겹치는가 — 겹치면 그리지 않는다."""
        lo, hi = ((slider.top(), slider.bottom()) if vertical
                  else (slider.left(), slider.right()))
        return pos + 1 >= lo and pos <= hi

    def paintEvent(self, e):
        super().paintEvent(e)
        if not self._ratios:
            return
        # QStyle을 통해 정확한 trough(groove)·핸들(slider) 영역을 얻는다
        opt = QStyleOptionSlider()
        self.initStyleOption(opt)
        st = self.style()
        groove = st.subControlRect(
            QStyle.CC_ScrollBar, opt, QStyle.SC_ScrollBarGroove, self)
        if groove.width() <= 0 or groove.height() <= 0:
            return
        slider = st.subControlRect(
            QStyle.CC_ScrollBar, opt, QStyle.SC_ScrollBarSlider, self)
        band, thick = self._marker_band(groove)
        vertical = self.orientation() == Qt.Vertical
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, False)
        painter.setPen(Qt.NoPen)
        painter.setBrush(self._MARKER_COLOR)
        if vertical:
            track_top = groove.top()
            for off, k in self._pixel_offsets(max(0, groove.height() - 2)):
                y = track_top + off
                if self._hits_slider(y, slider, True):
                    continue
                painter.fillRect(band, y, thick, 2, self._stacked_color(k))
        else:
            track_left = groove.left()
            for off, k in self._pixel_offsets(max(0, groove.width() - 2)):
                x = track_left + off
                if self._hits_slider(x, slider, False):
                    continue
                painter.fillRect(x, band, 2, thick, self._stacked_color(k))
        painter.end()


class DiffHighlightDelegate(QStyledItemDelegate):
    """modified 셀에서 A/B가 다른 문자 구간을 CELL_DIFF_HL(핑크)로 강조 렌더한다.
    강조 구간이 없는 셀(대다수)은 기본 델리게이트 렌더 그대로 — 빠른 경로.

    텍스트를 QTextDocument로 그려(셀값란과 동일 방식) 여러 줄('\\n')·문자 인덱스·
    수직 정렬을 그대로 반영한다. 단일 라인 가정으로 위치를 계산하면 멀티라인 셀에서
    강조가 엉뚱한 곳에 찍히므로, 문자 단위 배경은 QTextCharFormat으로 지정한다."""

    def paint(self, painter, option, index):
        model = index.model()
        r, c = index.row(), index.column()
        ranges = (model.diff_char_ranges(r, c)
                  if hasattr(model, "diff_char_ranges") else [])
        # 강조 구간(빨강)이 없으면 기본 델리게이트 렌더 — 빠른 경로.
        # (수식 셀의 파랑 폰트는 모델 ForegroundRole로 처리되어 QTextDocument가 불필요하다.)
        if not ranges:
            # 신규(added) 셀의 빈 쪽 → 대각선 해치로 매칭 표시.
            if (hasattr(model, "is_added_placeholder")
                    and model.is_added_placeholder(r, c)):
                opt = QStyleOptionViewItem(option)
                self.initStyleOption(opt, index)
                widget = opt.widget
                style = widget.style() if widget is not None else QApplication.style()
                opt.text = ""
                style.drawControl(QStyle.CE_ItemViewItem, opt, painter, widget)
                draw_diagonal_hatch(painter, opt.rect)
                return
            super().paint(painter, option, index)
            return

        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        widget = opt.widget
        style = widget.style() if widget is not None else QApplication.style()

        text = opt.text
        # 1) 텍스트 없이 배경/선택/프레임만 그린다 (기본 델리게이트와 동일한 배경/선택색).
        opt.text = ""
        style.drawControl(QStyle.CE_ItemViewItem, opt, painter, widget)

        text_rect = style.subElementRect(QStyle.SE_ItemViewItemText, opt, widget)
        fg = self._text_color(opt, index)
        if not _LINEBREAK_RE.search(text):
            self._paint_inline(painter, opt, text_rect, text, ranges, fg)
            return

        # 2) QTextDocument로 텍스트 구성 — '\n'은 줄바꿈, 소프트랩 없음(뷰와 동일).
        doc = QTextDocument()
        doc.setDocumentMargin(0)
        doc.setDefaultFont(opt.font)
        to = QTextOption(opt.displayAlignment)
        to.setWrapMode(QTextOption.NoWrap)
        doc.setDefaultTextOption(to)
        doc.setPlainText(text)

        # 기본 전경색 (선택 시 흰색, 그 외 ForegroundRole/기본).
        # ForegroundRole은 수식 결과 셀에 파랑을 돌려준다. 선택 셀은 흰색 우선(파란 선택 배경
        # 위 가독), 변경 구간(빨강)은 아래에서 덮어쓴다.
        base_fmt = QTextCharFormat()
        base_fmt.setForeground(fg)
        cur = QTextCursor(doc)
        cur.select(QTextCursor.Document)
        cur.mergeCharFormat(base_fmt)

        # 3) 다른 문자 구간에 핑크 배경 + 빨강 폰트. QTextDocument 위치는 '\n'을 1칸으로
        #    세어 flat 인덱스와 1:1 정렬되므로 diff_char_ranges 값을 그대로 쓴다.
        #    빨강 전경은 base_fmt(선택 시 흰색)를 덮어써 선택 셀에서도 강조가 보이게 한다.
        hl_fmt = QTextCharFormat()
        hl_fmt.setBackground(CELL_DIFF_HL)
        hl_fmt.setForeground(CELL_DIFF_FG)
        n = len(text)
        for start, end in ranges:
            start = max(0, start)
            end = min(n, end)
            if start >= end:
                continue
            c = QTextCursor(doc)
            c.setPosition(start)
            c.setPosition(end, QTextCursor.KeepAnchor)
            c.mergeCharFormat(hl_fmt)

        # 4) 텍스트 영역으로 clip + 수직 중앙 정렬해 그린다.
        painter.save()
        painter.setClipRect(text_rect)
        doc_h = doc.size().height()
        y = text_rect.top() + max(0, (text_rect.height() - doc_h) / 2)
        painter.translate(text_rect.left(), y)
        doc.drawContents(painter)
        painter.restore()


    @staticmethod
    def _text_color(opt, index):
        """셀 기본 전경색 — 선택 시 흰색, 그 외 ForegroundRole(수식=파랑)/기본색.
        변경 구간의 빨강은 호출부가 이 위에 덮어쓴다."""
        if opt.state & QStyle.State_Selected:
            return opt.palette.color(QPalette.Active, QPalette.HighlightedText)
        return index.data(Qt.ForegroundRole) or opt.palette.color(
            QPalette.Active, QPalette.Text)

    @staticmethod
    def _paint_inline(painter, opt, text_rect, text, ranges, fg):
        """단일 라인 강조 렌더 — QTextDocument 없이 QPainter 로 직접 그린다.

        ★ 강조 구간을 부분 문자열로 그리면 커닝/자간이 달라져 글자가 미세하게 밀린다.
          그래서 **전체 문자열을 두 번** 그린다: 기본색으로 한 번, 그 다음 강조 구간
          rect 로 clip 한 상태에서 빨강으로 한 번. 두 번 모두 같은 시작 x/baseline 을
          쓰므로 글리프 위치가 픽셀 단위로 동일하고, 강조 경계만 색이 갈린다.

        수직 정렬은 기본 델리게이트(drawItemText)와 같은 규칙(AlignVCenter)으로 맞춘다.
        긴 텍스트는 기존 QTextDocument 경로와 동일하게 **자르지 않고 clip** 한다 —
        문자 인덱스(ranges)가 생략기호(...) 없는 원문 기준이라 그래야 강조가 안 밀린다."""
        fm = opt.fontMetrics
        left = text_rect.left()
        # 강조 배경은 '셀 전체 높이'가 아니라 **글자 줄 높이**만 덮는다 —
        # QTextDocument 경로(문자 배경 = 줄 높이)와 픽셀 단위로 같게 맞추기 위함.
        line_h = fm.height()
        line_top = text_rect.top() + (text_rect.height() - line_h) // 2
        baseline = line_top + fm.ascent()
        n = len(text)

        spans = []
        for start, end in ranges:
            start = max(0, start)
            end = min(n, end)
            if start >= end:
                continue
            x1 = left + fm.horizontalAdvance(text, start)
            x2 = left + fm.horizontalAdvance(text, end)
            if x2 > x1:
                spans.append(QRect(x1, line_top, x2 - x1, line_h))

        painter.save()
        painter.setClipRect(text_rect)
        painter.setFont(opt.font)
        for rect in spans:                      # 1) 강조 배경(핑크)
            painter.fillRect(rect, CELL_DIFF_HL)
        painter.setPen(fg)                      # 2) 전체 텍스트 기본색
        painter.drawText(left, baseline, text)
        if spans:                               # 3) 강조 구간만 빨강으로 덮어 그리기
            painter.setPen(CELL_DIFF_FG)
            for rect in spans:
                painter.save()
                painter.setClipRect(rect, Qt.IntersectClip)
                painter.drawText(left, baseline, text)
                painter.restore()
        painter.restore()


class _Axis:
    """열/행 대칭 로직의 '축' 서술자 — 같은 알고리즘을 두 축에 공유하기 위한 파라미터.

    이 파일의 과거 버그 다수가 열 쪽만 고치고 행 쪽을 빼먹은(또는 그 반대) **축 불일치**
    였다. 그래서 대칭 연산의 알고리즘은 `_axis_*` 한 곳에만 두고, 기존 이름
    (`_select_col`/`_select_row` 등)은 축만 넘기는 얇은 위임으로 남긴다 — 호출부와
    테스트 표면은 그대로 유지하면서 한쪽만 고쳐지는 일을 구조적으로 막는다.

    주의: 아래 좌표/오버레이 매핑은 '이 축의 고정 밴드를 **실제로 그리는** 오버레이'를
    쓴다(열=left, 행=top). 본체에서는 고정 밴드가 숨김이라 매핑이 틀어지기 때문이다.
    """
    __slots__ = ("name", "is_col")

    def __init__(self, is_col: bool):
        self.is_col = is_col
        self.name = "col" if is_col else "row"

    # ── 뷰/모델 크기 ─────────────────────────────────────────────────────────
    def count(self, view) -> int:
        """이 축의 섹션 수(열축이면 열 수)."""
        return view.columnCount() if self.is_col else view.rowCount()

    def cross_count(self, view) -> int:
        """직교 축의 섹션 수(열축이면 행 수) — 전체 열/행 선택의 길이."""
        return view.rowCount() if self.is_col else view.columnCount()

    def cross_data_count(self, model) -> int:
        """직교 축의 데이터 개수(열축이면 데이터 행 수)."""
        return model.data_rows if self.is_col else model.data_cols

    def index(self, model, i: int, cross: int):
        """이 축 i, 직교 축 cross 위치의 QModelIndex."""
        return model.index(cross, i) if self.is_col else model.index(i, cross)

    def kind(self, model, i: int, cross: int) -> str:
        return model.cell_kind(cross, i) if self.is_col else model.cell_kind(i, cross)

    def of_staged(self, coord) -> int:
        """staged 좌표 (r, c) 에서 이 축 성분만."""
        r, c = coord
        return c if self.is_col else r

    # ── 선택 range ───────────────────────────────────────────────────────────
    def lo(self, rng) -> int:
        return rng.left() if self.is_col else rng.top()

    def hi(self, rng) -> int:
        return rng.right() if self.is_col else rng.bottom()

    # ── 좌표 / 히트테스트 ────────────────────────────────────────────────────
    def pos(self, point) -> int:
        return point.x() if self.is_col else point.y()

    def section_at(self, view, v: int) -> int:
        return view.columnAt(v) if self.is_col else view.rowAt(v)

    def overlay(self, fc):
        """이 축의 고정 밴드를 실제로 그리는 오버레이(열=left, 행=top)."""
        return fc.left if self.is_col else fc.top

    def frozen_count(self, fc) -> int:
        return fc._n_cols if self.is_col else fc._n_rows

    # ── 섹션 크기 / 숨김 (틀 고정 크기 반영용) ───────────────────────────────
    def hidden(self, view, i: int) -> bool:
        return view.isColumnHidden(i) if self.is_col else view.isRowHidden(i)

    def set_hidden(self, view, i: int, h: bool) -> None:
        if self.is_col:
            view.setColumnHidden(i, h)
        else:
            view.setRowHidden(i, h)

    def size(self, view, i: int) -> int:
        """이 축 섹션의 크기(열=폭, 행=높이)."""
        return view.columnWidth(i) if self.is_col else view.rowHeight(i)

    def set_size(self, view, i: int, v: int) -> None:
        if self.is_col:
            view.setColumnWidth(i, v)
        else:
            view.setRowHeight(i, v)

    def user_sizes(self, host) -> dict:
        """사용자가 직접 조절한 크기 기록(축별 dict)."""
        return host._user_col_widths if self.is_col else host._user_row_heights

    @property
    def span_attr(self) -> str:
        """FreezeController 의 고정 밴드 총 크기 속성명(_fw/_fh)."""
        return "_fw" if self.is_col else "_fh"

    def resized_signal(self, host):
        return host.column_resized if self.is_col else host.row_resized

    # ── 뷰의 축별 헬퍼 위임 ──────────────────────────────────────────────────
    def full_selected(self, view):
        return (view._full_columns_selected() if self.is_col
                else view._full_rows_selected())

    def touched(self, view):
        return view._touched_cols() if self.is_col else view._touched_rows()


AX_COL = _Axis(True)
AX_ROW = _Axis(False)


class _TintHeaderView(QHeaderView):
    """모델이 준 헤더 색을 실제로 칠하는 헤더 — 본체와 고정 밴드 오버레이가 함께 쓴다."""

    def paintSection(self, painter, rect, logicalIndex):
        """모델이 준 헤더 색을 **실제로** 칠한다.

        앱에 스타일시트가 걸려 있으면 Qt(QStyleSheetStyle)가 헤더 섹션 배경을 직접
        그리면서 모델의 BackgroundRole 을 무시한다. 그래서 키 열 노랑은 **한 번도
        칠해진 적이 없었고**(실측: 전 열이 QSS 색 #e8eaf0), 나중에 넣은 '한쪽에만 있는
        열' 연초록도 같은 이유로 죽어 있었다. QSS 에서 background 를 빼 봐도 모델 색이
        살아나지 않는다 — 스타일이 그리는 것 자체가 바뀌지 않기 때문이다.

        그래서 **덧칠**한다. super() 가 테두리·글자·아이콘까지 다 그린 뒤에 색을 얹되,
        덮어 칠하면 글자가 사라지므로 반투명(HEADER_TINT_ALPHA)으로 올린다.
        일반 색(HEADER_NORMAL_BG)은 얹지 않는다 — 얹어 봐야 같은 색이고 공짜도 아니다.

        ★ super() 호출을 save/restore 로 **감싸야** 한다. QHeaderView.paintSection 은
        painter 에 클립을 남기고 돌아와, 그대로 칠하면 전부 잘려 나간다(실측: 불투명
        빨강을 칠해도 화면에 단 한 픽셀도 안 나왔다). 이 한 줄이 기능의 전부다.
        """
        painter.save()
        super().paintSection(painter, rect, logicalIndex)
        painter.restore()
        model = self.model()
        if model is None:
            return
        bg = model.headerData(logicalIndex, self.orientation(), Qt.BackgroundRole)
        if not isinstance(bg, QColor) or bg == HEADER_NORMAL_BG:
            return
        tint = QColor(bg)
        tint.setAlpha(HEADER_TINT_ALPHA)
        painter.save()
        painter.fillRect(rect, tint)
        painter.restore()


class _FrozenView(QTableView):
    """틀 고정 헬퍼 뷰 — 본체 모델/선택모델을 공유한다. 자체 스크롤바 없음.
    휠 이벤트는 본체로 전달해 본체가 스크롤되고 컨트롤러가 헬퍼를 되동기하게 한다.
    선택 모델을 본체와 공유하므로(FreezeController._make_view) 고정 셀(키 열/행)을
    이 뷰에서 클릭하면 본체 선택이 갱신돼 셀값란·A/B 미러 등 기존 로직이 그대로 동작한다."""
    def __init__(self, host):
        super().__init__(host)
        self._host = host
        self._drag_anchor = None   # (row, col) 모델 좌표 — 오버레이↔본체 경계를 넘는 드래그 선택 앵커
        self._drag_custom = False  # 이번 드래그에서 _extend_drag로 직접 선택을 만들었는가

    def wheelEvent(self, event):
        self._host.wheelEvent(event)

    def scrollTo(self, index, hint=QAbstractItemView.EnsureVisible):
        """스크롤 위치는 전적으로 FreezeController가 제어한다. 선택/현재 셀 변경 시
        Qt가 자동 호출하는 scrollTo가 고정 뷰를 제 위치에서 밀어내지 않도록 무력화."""
        return

    # ── 경계를 넘는 드래그 선택 ───────────────────────────────────────────────
    # 고정 열/행은 본체와 분리된 오버레이 위젯이라, 오버레이에서 시작한 드래그는 Qt가 그
    # 위젯에 마우스를 grab 해 오버레이의 '보이는' 셀(고정 밴드)만 선택된다(스크롤 열/행은
    # 오버레이에서 숨김). 본체와 selection model 을 공유하므로, 드래그 중 커서를 본체 좌표로
    # 매핑해 공유 모델에 사각형 선택을 직접 만들어 경계를 넘는 선택을 가능하게 한다.
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            idx = self.indexAt(event.pos())
            self._drag_anchor = (idx.row(), idx.column()) if idx.isValid() else None
            self._drag_custom = False
            # 오버레이 드래그 중에도 A↔B 미러를 억제한다(본체와 선택 모델을 공유하므로
            # 호스트 플래그를 쓴다). 종료 시 1회만 미러된다.
            self._host._drag_selecting = True
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        # Ctrl/Shift 드래그는 Qt 기본(ExtendedSelection)에 맡긴다 — 다중 영역 누적/확장 시맨틱을
        # 보존한다. 그러지 않으면 _extend_drag의 ClearAndSelect가 기존 선택을 지운다(Ctrl+클릭에
        # 1px 지터가 섞여도 마찬가지). 평범한 좌드래그만 경계 넘는 사각형 선택으로 처리한다.
        if (self._drag_anchor is not None and (event.buttons() & Qt.LeftButton)
                and not (event.modifiers() & (Qt.ControlModifier | Qt.ShiftModifier))):
            self._extend_drag(event.globalPos())
            self._drag_custom = True
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        # 커스텀 드래그(_extend_drag)로 이동을 소비했으면 Qt는 드래그 상태로 진입하지 못한 채라,
        # super()의 릴리즈가 이를 '클릭'으로 보고 선택을 릴리즈 셀 하나로 리셋(고정 셀이면 0으로
        # 붕괴)해 방금 만든 다중 선택을 지운다(A열 키+ A1,B1 선택 시 풀림 버그). 커스텀 드래그였다면
        # super()를 건너뛰고 상태만 정리해 선택을 보존한다.
        if self._drag_custom:
            self._drag_custom = False
            self._drag_anchor = None
            self.setState(QAbstractItemView.NoState)
            event.accept()
            self._host._end_drag_select()
            return
        self._drag_anchor = None
        super().mouseReleaseEvent(event)
        self._host._end_drag_select()

    def _extend_drag(self, gp):
        """커서 전역좌표(gp)로 목표 (row, col)을 정하고 본체 공유 모델에 사각형 선택을 만든다.
        커서가 본체 스크롤 영역이면 host 좌표(끝 clamp/오토스크롤), 고정 밴드 위면 고정 열/행
        오버레이(left/top) 기준으로 매핑한다. ★ self(=이 오버레이)의 columnAt/rowAt을 쓰면 안 된다:
        top 오버레이는 고정 열이 숨김이라 고정 밴드에서 -1→0(A)으로 붕괴하고(키 행에서 K→D 드래그가
        A~D 전체 선택되는 버그), left 오버레이는 고정 행이 숨김이라 대칭 문제가 생긴다. 항상 고정
        열은 fc.left, 고정 행은 fc.top 기준(_frozen_col_at/_frozen_row_at)으로 매핑한다."""
        host = self._host
        fc = host._freeze
        hv = host.viewport().mapFromGlobal(gp)   # 본체(스크롤 셀) 뷰포트 좌표
        ar, ac = self._drag_anchor

        # 목표 열: 본체 영역이면 host 매핑, 고정 밴드(hv.x()<0)면 고정 열 오버레이(fc.left) 기준
        if hv.x() >= 0:
            col = host._drag_col_at(hv.x())
        elif fc is not None and getattr(fc, "_n_cols", 0) > 0:
            col = host._frozen_col_at(fc, gp)
        else:
            col = ac
        # 목표 행 (열과 대칭): 고정 행 밴드(hv.y()<0)면 고정 행 오버레이(fc.top) 기준
        if hv.y() >= 0:
            row = host._drag_row_at(hv.y())
        elif fc is not None and getattr(fc, "_n_rows", 0) > 0:
            row = host._frozen_row_at(fc, gp)
        else:
            row = ar

        host._select_range(ar, ac, row, col)
        host._set_current_cell_no_update(row, col)


class FreezeController(QObject):
    """엑셀 '틀 고정'(정식) — 본체(ExcelTableView)를 우하단 사분면으로 쓰고, 예약된 상단/좌측
    여백에 헬퍼 뷰 3개를 배치한다. 본체가 고정 행/열을 setRowHidden/setColumnHidden으로 숨겨
    스크롤 영역만 렌더하므로 '가림'이 없고, corner 헤더가 고정 눈금('2'/'B')을 정확히 표시한다.

    좌표(본체 기준, fr=frameWidth, hdrW=세로헤더폭, hdrH=가로헤더높이, fw=고정열폭합, fh=고정행높이합):
      corner:(fr, fr, hdrW+fw, hdrH+fh)          헤더 표시 → 고정 눈금 + 고정 코너 셀
      top   :(fr+hdrW+fw, fr+hdrH, 본체폭, fh)    헤더 숨김, 가로 동기 → 고정 행 데이터
      left  :(fr+hdrW, fr+hdrH+fh, fw, 본체높이)  헤더 숨김, 세로 동기 → 고정 열 데이터
    본체 여백은 ExcelTableView.updateGeometries() 오버라이드가 reposition()과 함께 적용한다
    (여백 예약은 _BandHeaderView 의 sizeHint — 이유는 그 오버라이드 주석 참고).
    """
    _MIN_BODY_W = 80    # 고정 영역이 본체를 다 먹으면 freeze 중단(잔여 본체 최소치)
    _MIN_BODY_H = 44

    def __init__(self, host):
        super().__init__(host)
        self.host = host
        host._freeze = self
        self._n_rows = 1          # 고정 행 수 = key_row + 1
        self._n_cols = 1          # 고정 열 수 = key_col + 1 (key_col<0 → 0)
        self._fw = 0              # 고정 열 폭합(refresh에서 캡처 — 본체 숨김과 무관하게 안정)
        self._fh = 0              # 고정 행 높이합
        self._active = False
        # left 오버레이의 행 숨김이 본체와 일치하는지. False 면 다음 미러를 전 행으로
        # 돌린다(델타 미러의 전제 조건). 모델 리셋/freeze 해제로 무너진다.
        self._rows_synced = False
        self._fixing = False      # 밴드 어긋남 보정 중 재진입 방지
        self.corner = self._make_view(headers=True)
        self.top = self._make_view(headers=False)
        self.left = self._make_view(headers=False)
        self._views = (self.corner, self.top, self.left)
        # 오버레이의 '스크롤하는 축'만 픽셀 단위로 — 본체(ScrollPerItem)의 헤더 픽셀 오프셋을
        # 그대로 복사해 픽셀 단위로 정렬한다(_sync_top_h/_sync_left_v). 본체와 오버레이는
        # 스크롤바 표시 여부가 달라 ScrollPerItem '최댓값'이 1 어긋나는데, 값만 미러하면
        # 스크롤 맨 끝에서 한 칸 밀린다. 픽셀 오프셋 복사는 최댓값 quirk와 무관하게 정확하다.
        self.top.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.left.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        # 본체를 열/행 단위 스크롤로 고정 → 스크롤바 값이 곧 열/행 인덱스라 헬퍼와 값 동기로 정렬된다.
        host.setHorizontalScrollMode(QAbstractItemView.ScrollPerItem)
        host.setVerticalScrollMode(QAbstractItemView.ScrollPerItem)
        host.horizontalScrollBar().valueChanged.connect(self._on_h_scroll)
        host.verticalScrollBar().valueChanged.connect(self._on_v_scroll)
        host.horizontalHeader().sectionResized.connect(self._on_col_resized)
        host.verticalHeader().sectionResized.connect(self._on_row_resized)
        # 모델 리셋(populate)은 뷰의 행 숨김을 초기화할 수 있고 그 동작이 플랫폼마다
        # 다르다 — 본체와 left 가 서로 다른 상태에서 출발할 수 있으므로 델타 미러를
        # 금지하고 다음 1회를 전 행 스캔으로 돌린다. (모델 객체는 교체되지 않는다.)
        host.model().modelReset.connect(self._on_model_reset)
        # corner(고정 눈금) 헤더 우클릭 → 키 열/행 설정·해제 메뉴.
        # ★ 반드시 popup()(비모달)로 띄운다. 자식 오버레이 헤더 이벤트 처리 중 모달 메뉴(exec_ 중첩
        #   이벤트루프)를 쓰면 Qt 상태가 깨져 access violation(Windows fatal exception)이 난다.
        #   본체 헤더는 exec_ 그대로(문제 없음). corner는 최소 키 메뉴만 popup으로 제공.
        ch = self.corner.horizontalHeader()
        cv = self.corner.verticalHeader()
        ch.setContextMenuPolicy(Qt.CustomContextMenu)
        cv.setContextMenuPolicy(Qt.CustomContextMenu)
        ch.customContextMenuRequested.connect(self._corner_col_menu)
        cv.customContextMenuRequested.connect(self._corner_row_menu)
        # 키 열/행은 본체에서 숨겨져 있어 '틀 고정' corner 오버레이 헤더에서만 드래그로
        # 크기 조절이 가능하다. 이 헤더의 sectionResized 를 처리하지 않으면 corner 내부
        # 열폭/행높이만 바뀌고 _fw/_fh·본체 여백·헬퍼·반대 패널이 갱신되지 않아 고정 밴드가
        # 본체와 어긋나고 셀이 겹쳐 보인다 → 호스트/전 뷰/미러로 전파한다.
        ch.sectionResized.connect(self._on_corner_col_resized)
        cv.sectionResized.connect(self._on_corner_row_resized)
        # 고정 밴드 경계를 넘는 헤더 드래그 선택: corner 헤더도 host 필터/anchor에 등록.
        host.register_frozen_headers(self.corner)

    def _corner_col_menu(self, pos):
        """고정 열 헤더(corner) 우클릭 — 병합 준비/취소 + 키 열 설정/해제(비모달 popup).
        ★ 반드시 popup(비모달). exec_(모달)은 자식 오버레이 이벤트 중 access violation 유발."""
        if not self._alive():
            return
        host = self.host
        ch = self.corner.horizontalHeader()
        col = ch.logicalIndexAt(pos)
        if col < 0:
            return
        m = QMenu(host)
        m.setStyleSheet(MENU_QSS)
        m.setAttribute(Qt.WA_DeleteOnClose)

        # 병합 준비/취소 — 본체 헤더 메뉴와 동일(대상 열 = 선택 반영). 키 열도 변경/스테이징
        # 셀이 있으면 노출(예: 신규 행의 키 값). 없으면 항목 자체가 안 뜬다.
        target_cols = host._selected_header_cols(col)
        cols_label = ", ".join(get_column_letter(c + 1) for c in target_cols)
        has_changed = any(host._col_has_changed(c) for c in target_cols)
        has_staged = host._cols_have_staged(target_cols)
        if has_changed:
            m.addAction(f"A → B  병합 준비  [{cols_label}열]").triggered.connect(
                lambda _=False, cs=target_cols:
                (host._select_cols(cs), host.stage_requested.emit(DIR_A2B)))
            m.addAction(f"B → A  병합 준비  [{cols_label}열]").triggered.connect(
                lambda _=False, cs=target_cols:
                (host._select_cols(cs), host.stage_requested.emit(DIR_B2A)))
        if has_staged:
            if has_changed:
                m.addSeparator()
            m.addAction(f"병합 준비 취소  [{cols_label}열]").triggered.connect(
                lambda _=False, cs=target_cols:
                (host._select_cols(cs), host.unstage_requested.emit()))
        if has_changed or has_staged:
            m.addSeparator()

        if col == host._key_col:
            m.addAction("🔓  키 열 해제 (ROW 순서 기반 비교)").triggered.connect(
                lambda: host.key_col_changed.emit(-1))
        else:
            letter = get_column_letter(col + 1)
            m.addAction(key_header_icon(), f"키 열로 설정  [{letter}열]").triggered.connect(
                lambda: host.key_col_changed.emit(col))

        # ── 변경 검사 제외/해제 — 키 열보다 좌측(고정 밴드) 열도 제외 가능하게 노출 ──
        # 키 열 자신은 키 매칭 기준이라 제외 대상에서 뺀다(본체 헤더 메뉴와 동일 규칙).
        # 선택에 '제외됨'과 '비제외'가 섞이면 두 항목을 모두 노출한다.
        excl_cols = [c for c in target_cols if c != host._key_col]
        to_exclude = [c for c in excl_cols if c not in host._excluded_cols]
        to_unexclude = [c for c in excl_cols if c in host._excluded_cols]
        if to_exclude or to_unexclude:
            m.addSeparator()
        if to_exclude:
            lbl = ", ".join(get_column_letter(c + 1) for c in to_exclude)
            m.addAction(exclude_header_icon(),
                        f"변경 검사에서 제외  [{lbl}열]").triggered.connect(
                lambda _=False, cs=to_exclude: host.columns_exclude_set.emit(cs, True))
        if to_unexclude:
            lbl = ", ".join(get_column_letter(c + 1) for c in to_unexclude)
            m.addAction(reset_header_icon(),
                        f"검사 제외 해제  [{lbl}열]").triggered.connect(
                lambda _=False, cs=to_unexclude: host.columns_exclude_set.emit(cs, False))
        m.popup(ch.mapToGlobal(pos))

    def _corner_row_menu(self, pos):
        """고정 행 헤더(corner) 우클릭 — 병합 준비/취소 + 키 행 설정(비모달 popup).
        ★ 반드시 popup(비모달). exec_(모달)은 자식 오버레이 이벤트 중 access violation 유발."""
        if not self._alive():
            return
        host = self.host
        cv = self.corner.verticalHeader()
        row = cv.logicalIndexAt(pos)
        if row < 0:
            return
        orig = host.model().orig_row(row)   # display→원본 파일 행
        if orig is None:
            return
        m = QMenu(host)
        m.setStyleSheet(MENU_QSS)
        m.setAttribute(Qt.WA_DeleteOnClose)

        # 병합 준비/취소 — 본체 행 헤더 메뉴와 동일(대상 행 = 선택 반영). 키 행에서도 노출.
        target_rows = host._selected_header_rows(row)
        suffix = f"  [{len(target_rows)}개 행]" if len(target_rows) > 1 else ""
        has_changed = any(host._row_has_changed(r) for r in target_rows)
        has_staged = host._rows_have_staged(target_rows)
        if has_changed:
            m.addAction(f"A → B  병합 준비{suffix}").triggered.connect(
                lambda _=False, rs=target_rows:
                (host._select_rows(rs), host.stage_requested.emit(DIR_A2B)))
            m.addAction(f"B → A  병합 준비{suffix}").triggered.connect(
                lambda _=False, rs=target_rows:
                (host._select_rows(rs), host.stage_requested.emit(DIR_B2A)))
        if has_staged:
            if has_changed:
                m.addSeparator()
            m.addAction(f"병합 준비 취소{suffix}").triggered.connect(
                lambda _=False, rs=target_rows:
                (host._select_rows(rs), host.unstage_requested.emit()))

        # 키 행 지정 — 현재 키 행이 아닌 단일 행에서만(초기화 기능은 제거됨).
        if orig != host._key_row and len(target_rows) == 1:
            if has_changed or has_staged:
                m.addSeparator()
            m.addAction(key_header_icon(), f"키 행으로 설정  [{orig + 1}행]").triggered.connect(
                lambda _=False, o=orig: host.key_row_changed.emit(o))

        if m.isEmpty():
            return   # 키 행인데 병합할 것도 없음 → 메뉴 미표시
        m.popup(cv.mapToGlobal(pos))

    def _overlay_cell_menu(self, view, pos):
        """고정 밴드(오버레이) 셀 우클릭 → 본체의 셀 메뉴를 그대로 띄운다.
        오버레이는 본체와 선택 모델을 공유하므로 판단·동작이 본체와 동일하다.
        (헤더는 각자 자기 정책을 쓰므로 corner 헤더 메뉴와 충돌하지 않는다.)"""
        if not self._alive():
            return
        self.host._popup_cell_menu(view.viewport().mapToGlobal(pos))

    def _make_view(self, headers: bool):
        host = self.host
        v = _FrozenView(host)
        # 키 열 헤더는 본체에서 숨겨지고 **이 오버레이가** 그린다 — 색을 칠하는
        # 헤더로 바꿔 두지 않으면 키 노랑이 여기서만 또 안 나온다.
        # (QTableView 는 sectionsClickable/highlightSections 를 자기가 만든 기본
        #  헤더에만 켠다 — 교체하면서 직접 켜 줘야 헤더 클릭이 산다.)
        for _h, _set in ((_TintHeaderView(Qt.Horizontal, v), v.setHorizontalHeader),
                         (_TintHeaderView(Qt.Vertical, v), v.setVerticalHeader)):
            _h.setSectionsClickable(True)
            _h.setHighlightSections(True)
            _set(_h)
        v.setModel(host.model())
        # 선택 모델을 본체와 공유 — 고정 셀(키 열/행)을 이 뷰에서 클릭하면 본체 선택이 갱신되고
        # selectionChanged가 발화해 셀값란·A/B 미러·수식 플래그 등 기존 배선이 그대로 동작한다.
        v.setSelectionModel(host.selectionModel())
        v.setItemDelegate(DiffHighlightDelegate(v))
        v.setFocusPolicy(Qt.NoFocus)
        v.setSelectionBehavior(QAbstractItemView.SelectItems)
        v.setSelectionMode(QAbstractItemView.ExtendedSelection)
        # 선택색 통일(비포커스에도 파랑) — 본체와 동일 팔레트.
        force_active_highlight(v)
        v.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # 고정 밴드 셀 우클릭 → 본체와 동일한 병합 준비/취소 메뉴. 이 배선이 없으면 키 열
        # 및 그 좌측 열·키 행 셀은 본체에서 숨겨져 있어(오버레이가 그림) 우클릭해도 아무
        # 메뉴가 안 뜬다 — 그 열들을 셀 우클릭으로 병합 준비할 방법이 없어진다.
        v.setContextMenuPolicy(Qt.CustomContextMenu)
        v.customContextMenuRequested.connect(
            lambda pos, _v=v: self._overlay_cell_menu(_v, pos))
        v.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        v.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        v.setHorizontalScrollMode(QAbstractItemView.ScrollPerItem)
        v.setVerticalScrollMode(QAbstractItemView.ScrollPerItem)
        v.setAlternatingRowColors(False)
        v.setFrameShape(QTableView.NoFrame)
        v.setStyleSheet("QTableView { border: none; }")
        v.setFont(ui_font(9))
        v.verticalHeader().setDefaultSectionSize(
            host.verticalHeader().defaultSectionSize())
        v.horizontalHeader().setVisible(headers)
        v.verticalHeader().setVisible(headers)
        if headers:   # corner — 고정 눈금(키 아이콘 포함)
            v.horizontalHeader().setIconSize(QSize(14, 14))
            v.verticalHeader().setIconSize(QSize(14, 14))
        v.hide()
        return v

    def _on_model_reset(self):
        self._rows_synced = False

    def _alive(self) -> bool:
        """teardown 중(헬퍼 C++ 객체 삭제됨) 시그널이 도착해도 크래시하지 않도록 가드.
        스크롤·리사이즈 핸들러가 매번 부르는 핫패스라 제너레이터 없이 펼쳐 쓴다."""
        isdeleted = sip.isdeleted
        if isdeleted(self) or isdeleted(self.host):
            return False
        corner, top, left = self._views
        return not (isdeleted(corner) or isdeleted(top) or isdeleted(left))

    @property
    def active(self) -> bool:
        return self._active

    @staticmethod
    def frozen_span(key_row: int, key_col: int, rows: int, cols: int) -> tuple:
        """앵커(key_row/key_col)와 데이터 크기로 고정 (행수, 열수) 계산.
        key_col<0(ROW 순서)면 열 고정 없음(0)."""
        n_rows = min((key_row if key_row and key_row > 0 else 0) + 1, max(rows, 0))
        n_cols = 0 if key_col is None or key_col < 0 else min(key_col + 1, max(cols, 0))
        return n_rows, n_cols

    def frozen_px(self) -> tuple:
        # 본체가 고정 행/열을 숨기면 폭·높이가 0이 되므로, refresh에서 캡처한 값을 쓴다.
        return self._fw, self._fh

    # ── 외부 API ─────────────────────────────────────────────────────────────
    def clear(self):
        """freeze 비활성화 — 헬퍼 숨김 + 본체 열숨김/여백 원복(행 숨김은 필터가 소유)."""
        if not self._alive():
            return
        self._active = False
        self._rows_synced = False   # 본체 행 숨김이 left 없이 바뀔 수 있다(_clear_row_filter)
        for v in self._views:
            v.hide()
        host = self.host
        for c in range(host.model().columnCount()):
            if host.isColumnHidden(c):
                host.setColumnHidden(c, False)
        host.updateGeometries()   # super()가 기본 여백 복원

    def refresh(self, changed_rows=None):
        """오버레이 재적용. changed_rows 를 주면 좌측 오버레이의 행 미러를 그 행들로
        한정한다(호출부가 실제로 뒤집힌 행을 이미 알고 있을 때). None 이면 전 행 스캔."""
        if not self._alive():
            return
        host = self.host
        model = host.model()
        rows = getattr(model, "data_rows", 0)
        if not rows or getattr(model, "_mode", None) != "diff":
            self.clear()
            return
        self._n_rows, self._n_cols = self.frozen_span(
            getattr(host, "_key_row", 0), getattr(host, "_key_col", 0),
            rows, getattr(model, "data_cols", 0))
        self._active = True
        ncols = model.columnCount()
        # 벌크 뮤테이션 구간: _applying_sizes로 freeze 핸들러/오버라이드의 per-op 무거운 작업을 억제하고
        # (40k행 O(R) 미러가 폭풍이 되지 않도록) 화면 갱신도 잠시 끈다. 끝난 뒤 1회만 full reposition.
        prev_host = host.updatesEnabled()
        prev_v = [v.updatesEnabled() for v in self._views]
        prev_flag = getattr(host, "_applying_sizes", False)
        host._applying_sizes = True
        host.setUpdatesEnabled(False)
        for v in self._views:
            v.setUpdatesEnabled(False)
        try:
            # 1) 본체 고정 열을 잠시 해제해 실제 폭을 캡처(숨김 열은 폭 0이라 캡처 불가).
            for c in range(ncols):
                if host.isColumnHidden(c):
                    host.setColumnHidden(c, False)
            colw = [host.columnWidth(c) for c in range(ncols)]
            self._fw = sum(colw[c] for c in range(self._n_cols))
            dh = host.verticalHeader().defaultSectionSize()
            urh = getattr(host, "_user_row_heights", {})
            self._fh = sum(urh.get(r, dh) for r in range(self._n_rows))
            # 2) 본체/헬퍼 열 숨김.
            self._apply_col_hidden()
            # 3) 헬퍼 열 폭/행 높이를 '숨김 이후' 캡처값으로 명시 적용 — setColumnHidden 부작용이
            #    폭을 0으로 만드는 것을 되돌린다(마지막에 적용해야 덮어쓰이지 않음).
            for c in range(ncols):
                for v in self._views:
                    if v.columnWidth(c) != colw[c]:
                        v.setColumnWidth(c, colw[c])
            for r, h in urh.items():
                for v in self._views:
                    if v.rowHeight(r) != h:
                        v.setRowHeight(r, h)
            self._mirror_hidden_rows(changed_rows)
        finally:
            host._applying_sizes = prev_flag
            host.setUpdatesEnabled(prev_host)
            for v, pv in zip(self._views, prev_v):
                v.setUpdatesEnabled(pv)
        host.updateGeometries()   # 벌크 종료(_applying_sizes 복원) 후 1회 여백 재적용 + reposition()
        self._sync_scroll()

    # ── 숨김/크기/스크롤 미러 ────────────────────────────────────────────────
    def _apply_col_hidden(self):
        """본체·헬퍼의 열 숨김: 본체/top은 고정 열 숨김(스크롤 열만), left/corner는 스크롤 열 숨김."""
        host = self.host
        for c in range(host.model().columnCount()):
            frozen = c < self._n_cols
            if host.isColumnHidden(c) != frozen:
                host.setColumnHidden(c, frozen)
            if self.top.isColumnHidden(c) != frozen:
                self.top.setColumnHidden(c, frozen)
            if self.left.isColumnHidden(c) != (not frozen):
                self.left.setColumnHidden(c, not frozen)
            if self.corner.isColumnHidden(c) != (not frozen):
                self.corner.setColumnHidden(c, not frozen)

    def _sync_sizes(self):
        host = self.host
        for c in range(host.model().columnCount()):
            if host.isColumnHidden(c):
                continue   # 숨겨진(고정) 열은 폭 0 → 복사하면 헬퍼의 고정 열 폭이 사라짐
            w = host.columnWidth(c)
            for v in self._views:
                if v.columnWidth(c) != w:
                    v.setColumnWidth(c, w)
        for r, h in getattr(host, "_user_row_heights", {}).items():
            for v in self._views:
                if v.rowHeight(r) != h:
                    v.setRowHeight(r, h)

    def _mirror_hidden_rows(self, rows=None):
        """left 오버레이의 행 숨김을 본체와 맞춘다.
        (top/corner 는 고정 행만 노출하므로 높이 클립으로 충분 — left 만 정렬이 필요하다.)

        rows: 본체에서 **실제로 숨김이 뒤집힌 행**. 주면 그 행만 훑는다 — 열 제외처럼
        행 가시성이 그대로인 갱신에서 미러 비용이 0 이 된다(실측 6,328행 x 71열,
        A/B 4회: 열 제외 ON+OFF 46.5ms -> 32.1ms). 반대로 '변경점만 보기' 토글은
        대부분의 행이 실제로 뒤집혀 델타가 곧 전 행이라 이득이 없다(측정 노이즈 안).
        단, left 가 본체와 동기라는 보장이 없으면(_rows_synced False) 델타를 무시하고
        전 행을 훑는다.
        """
        host = self.host
        left = self.left
        full = rows is None or not self._rows_synced
        self._rows_synced = True
        if full:
            targets = range(host.model().rowCount())
        elif rows:
            # ★ 반드시 오름차순으로 훑는다 — 행 인덱스가 뒤죽박죽이면 Qt 의 행 숨김
            #   갱신이 순차 접근일 때보다 눈에 띄게 느리다(실측: 정렬 없이 set 을 그대로
            #   돌리면 토글이 기준선보다 20% 느렸다). 호출부는 이미 오름차순 리스트를
            #   주므로 이 sorted() 는 사실상 공짜이고, 안전망으로만 남긴다.
            targets = sorted(rows)
        else:
            return
        # 대량 행 숨김을 ScrollPerItem에서 하면 per-item 재계산으로 ~O(R²) → 루프 동안
        # ScrollPerPixel로 전환 후 복원(본체 필터 루프와 동일 취지).
        prev_vmode = left.verticalScrollMode()
        left.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        # setRowHidden 하나하나가 sectionResized 를 쏘고, 그 슬롯(_on_row_resized)은
        # _applying_sizes 가드로 어차피 아무 일도 하지 않는다 — 시그널 자체를 막아
        # 행마다 나던 슬롯 디스패치(토글 1회당 ~12,700회)를 없앤다.
        # ※ cProfile 로는 이 디스패치가 커 보였지만 프로파일러 오버헤드였다 — 실제
        #   A/B 로는 유의미한 차이가 없었다. 헛일을 줄이는 정리로만 남긴다.
        vh = left.verticalHeader()
        prev_blocked = vh.blockSignals(True)
        try:
            for r in targets:
                hidden = host.isRowHidden(r)
                if left.isRowHidden(r) != hidden:
                    left.setRowHidden(r, hidden)
        finally:
            vh.blockSignals(prev_blocked)
            left.setVerticalScrollMode(prev_vmode)
            # 위에서 헤더 시그널을 막고 스크롤 모드를 바꿔 뒀기 때문에 Qt 가 이 뷰의
            # **스크롤 범위를 다시 계산하지 않는다.** 범위가 옛값(짧은 쪽)으로 남으면
            # _sync_left_v 의 setValue 가 거기서 잘려(clamp) 고정 열이 본체를 못 따라가고
            # 화면 위쪽에 붙박인다 — 아래로 스크롤할수록 키 값이 밀려 보인다.
            # 실측: 전체 행 표시에서 본체는 4460 까지 가는데 오버레이 최댓값이 82 였다.
            left.doItemsLayout()

    def _sync_scroll(self):
        if not self._alive():
            return
        self._sync_top_h()
        self.top.verticalScrollBar().setValue(0)
        self._sync_left_v()
        self.left.horizontalScrollBar().setValue(0)
        self.corner.horizontalScrollBar().setValue(0)
        self.corner.verticalScrollBar().setValue(0)

    def _sync_top_h(self):
        """상단(고정 행) 오버레이의 가로 위치를 본체와 픽셀 단위로 정렬한다.
        본체는 ScrollPerItem이라 최좌측 열이 픽셀 경계에 스냅되고, 그 픽셀 오프셋은
        가로 헤더의 offset()으로 얻는다. 오버레이는 가로 ScrollPerPixel이므로 그 값을
        그대로 넣으면 동일 콘텐츠(같은 열·폭)가 픽셀 정확히 겹친다. 스크롤바 최댓값이
        서로 1 어긋나도(표시/AlwaysOff 차이) 무관 — 마지막 페이지에서도 안 밀린다."""
        if not self._alive():
            return
        self.top.horizontalScrollBar().setValue(self.host.horizontalHeader().offset())

    def _sync_left_v(self):
        """좌측(고정 열) 오버레이의 세로 위치를 본체와 픽셀 단위로 정렬한다(_sync_top_h의 세로
        대칭). 본체 세로 헤더의 offset()이 곧 콘텐츠 세로 픽셀 오프셋이고, 오버레이는 세로
        ScrollPerPixel이라 그 값을 그대로 넣으면 같은 행이 픽셀 정확히 겹친다. 스크롤바 최댓값
        차이(표시/AlwaysOff)와 무관하므로 스크롤 맨 아래에서도 고정 열이 밀리지 않는다."""
        if not self._alive():
            return
        self.left.verticalScrollBar().setValue(self.host.verticalHeader().offset())
        self._fix_left_drift()

    def _fix_left_drift(self):
        """밴드가 **그려지는 위치**로도 본체와 맞는지 보고, 어긋나면 자리를 다시 잡는다.

        오프셋만 맞추는 것으로는 부족하다. 밴드는 별도 위젯이고 그 y 는 reposition()
        이 `fr + hdrH + fh` 로 정하는데, reposition 은 **크기 변경 때만** 돈다. 그 뒤
        가로 헤더 높이나 고정 행 높이(fh)가 바뀌면 밴드만 몇 픽셀 어긋난 채 남는다.
        그래도 오프셋은 서로 같으니 동기화는 "맞다"고 보고 넘어간다 — 그래서 **스크롤을
        아무리 해도 안 고쳐지고** 창 크기를 바꿔야 돌아왔다.

        실측(36↔40 빌드 Data_MailBox_CS.xlsx): 키 열이 반 행 아래로 밀린 채 유지됐고,
        행 번호·본문은 서로 맞아서 "키 값만 아래로 밀린다" 로 보였다. 창을 조금만
        키웠다 줄이면 즉시 정상으로 돌아왔다.

        한 행의 화면 위치만 비교하므로 스크롤 1회당 비용은 무시할 만하다.
        """
        if self._fixing or not self._active or not self._alive():
            return
        host, left = self.host, self.left
        if not left.isVisible():
            return
        r = host.rowAt(0)          # 화면 맨 위에 걸린 행 — 양쪽이 같은 행을 그려야 한다
        # left.isRowHidden 은 **부르지 않는다.** 행 미러링이 델타로 도는지 세는 테스트가
        # 오버레이 접근 횟수를 보고 있어, 여기서 한 번만 더 봐도 그 보장이 깨진 것처럼
        # 보인다. 높이로 판정해도 충분하다(숨긴 행은 높이가 0).
        if r < 0 or left.rowHeight(r) == 0:
            return
        hy = host.viewport().mapToGlobal(QPoint(0, host.rowViewportPosition(r))).y()
        ly = left.viewport().mapToGlobal(QPoint(0, left.rowViewportPosition(r))).y()
        if hy == ly:
            return
        self._fixing = True
        try:
            self.reposition()      # 처음부터 다시 계산 — 창 크기 변경이 하던 일
        finally:
            self._fixing = False

    def reposition(self):
        """헬퍼 3개의 위치/크기를 본체 헤더·고정 크기 기준으로 재계산. updateGeometries에서 호출."""
        if not self._alive() or not self._active:
            return
        host = self.host
        fr = host.frameWidth()
        hdrW = host.verticalHeader().width()
        hdrH = host.horizontalHeader().height()
        fw, fh = self.frozen_px()
        W, H = host.width(), host.height()
        body_w = W - 2 * fr - hdrW - fw
        body_h = H - 2 * fr - hdrH - fh
        # 고정 영역이 본체를 다 먹으면 freeze 표시 중단(스크롤 영역 확보 불가)
        if body_w < self._MIN_BODY_W or body_h < self._MIN_BODY_H:
            for v in self._views:
                v.hide()
            return
        # 본체는 스크롤바가 뜨면 그만큼 셀 렌더 영역이 줄어드는데, 고정 헬퍼 뷰는 스크롤바가
        # 항상 꺼져 있어(ScrollBarAlwaysOff) 그만큼 더 많은 행/열을 그린다. 그러면 ScrollPerItem
        # 최댓값이 서로 달라져 스크롤 끝(하단/우단)에서 고정 열/행이 본체와 한 칸 어긋난다.
        # → 헬퍼 뷰 크기에서 본체 스크롤바 두께를 빼 '보이는 행/열 수'를 본체와 맞춘다.
        hbar = host.horizontalScrollBar()
        vbar = host.verticalScrollBar()
        hbar_h = hbar.height() if hbar.isVisible() else 0
        vbar_w = vbar.width() if vbar.isVisible() else 0
        self.corner.verticalHeader().setFixedWidth(hdrW)
        self.corner.horizontalHeader().setFixedHeight(hdrH)
        self.corner.setGeometry(fr, fr, hdrW + fw, hdrH + fh)
        self.corner.setVisible(True)
        self.top.setGeometry(fr + hdrW + fw, fr + hdrH, max(0, body_w - vbar_w), fh)
        self.top.setVisible(True)
        if self._n_cols > 0:
            self.left.setGeometry(fr + hdrW, fr + hdrH + fh, fw, max(0, body_h - hbar_h))
            self.left.setVisible(True)
        else:
            self.left.hide()
        self.corner.raise_()
        self.top.raise_()
        self.left.raise_()
        self._sync_scroll()

    # ── 시그널 핸들러 (teardown 가드 필수) ───────────────────────────────────
    def _on_h_scroll(self, value):
        if self._alive() and self._active:
            self._sync_top_h()

    def _on_v_scroll(self, value):
        if self._alive() and self._active:
            self._sync_left_v()

    def _on_col_resized(self, idx, old, new):
        # 벌크 작업(_applying_sizes: 필터의 setRowHidden 폭풍, 크기 일괄 적용) 중엔 발화 무시 —
        # 그러지 않으면 숨김 해제되는 행마다 sectionResized→여기서 O(C)+reposition이 돌아 O(R×C) 폭주.
        if not self._alive() or not self._active or getattr(self.host, "_applying_sizes", False):
            return
        self._sync_sizes()
        self.host.updateGeometries()

    def _on_row_resized(self, idx, old, new):
        if new <= 0 or not self._alive() or not self._active \
                or getattr(self.host, "_applying_sizes", False):
            return
        self._sync_sizes()
        self.host.updateGeometries()

    def _on_corner_resized(self, ax: _Axis, idx, new):
        """corner 오버레이에서 고정(키) 열 폭 / 행 높이를 드래그로 조절한 경우."""
        if new <= 0 or not self._alive() or not self._active \
                or idx >= ax.frozen_count(self) \
                or getattr(self.host, "_applying_sizes", False):
            return
        self._apply_frozen_size(ax, idx, new, mirror=True)

    def _on_corner_col_resized(self, idx, old, new):
        self._on_corner_resized(AX_COL, idx, new)

    def _on_corner_row_resized(self, idx, old, new):
        self._on_corner_resized(AX_ROW, idx, new)

    def _apply_frozen_size(self, ax: _Axis, idx, new, mirror=False):
        """고정(키) 열 폭 / 행 높이 변경을 호스트·전 헬퍼 뷰·_fw(_fh)·본체 여백에 반영
        (+선택적 미러). 열/행 공용 구현 — 한쪽만 고쳐지는 축 불일치를 막는다.

        키 열/행은 본체에서 숨겨져 있어 host.setColumnWidth/setRowHeight 가 무시되므로,
        잠시 숨김을 풀고 크기를 심어 Qt 가 기억하게 한다(이후 refresh 가 그 크기를 캡처)."""
        if not self._alive() or new <= 0 or idx < 0 or idx >= ax.frozen_count(self):
            return
        host = self.host
        prev = getattr(host, "_applying_sizes", False)
        host._applying_sizes = True
        try:
            was_hidden = ax.hidden(host, idx)
            if was_hidden:
                ax.set_hidden(host, idx, False)
            if ax.size(host, idx) != new:
                ax.set_size(host, idx, new)
            if was_hidden:
                ax.set_hidden(host, idx, True)
            ax.user_sizes(host)[idx] = new
            for v in self._views:
                if ax.size(v, idx) != new:
                    ax.set_size(v, idx, new)
        finally:
            host._applying_sizes = prev
        # 고정 밴드 총 크기 재계산(corner 기준 — host 는 숨김이라 0)
        setattr(self, ax.span_attr,
                sum(ax.size(self.corner, i) for i in range(ax.frozen_count(self))))
        host.updateGeometries()
        if mirror:
            ax.resized_signal(host).emit(idx, new)

    def _apply_frozen_col_width(self, idx, new, mirror=False):
        self._apply_frozen_size(AX_COL, idx, new, mirror=mirror)

    def _apply_frozen_row_height(self, idx, new, mirror=False):
        self._apply_frozen_size(AX_ROW, idx, new, mirror=mirror)


class _BandHeaderView(_TintHeaderView):
    """틀 고정 밴드 크기를 sizeHint 에 포함시키는 헤더.

    QTableView.updateGeometries() 는 ① 헤더 sizeHint 로 뷰포트 여백을 잡고
    ② **그 시점의 뷰포트 크기**로 스크롤바 range 를 계산한다. 고정 밴드 자리를
    뒤늦게 setViewportMargins 로 더하면 range 는 이미 '밴드를 안 뺀 더 넓은
    뷰포트' 기준으로 계산된 뒤라, 넘치는 폭이 밴드 폭보다 작을 때 range 가 0 이
    되어 **스크롤바가 아예 안 뜬다**. 밴드를 sizeHint 에 미리 넣어 Qt 가 처음부터
    올바른 여백·range 를 계산하게 한다(설계 배경은 updateGeometries 주석 참고).
    """

    # 클래스 속성 기본값 — QHeaderView.__init__ 이 끝나기 전에도 Qt 가 sizeHint()를
    # 호출할 수 있어(가상 함수 콜백), 인스턴스 속성만으론 AttributeError 로 죽는다.
    _band = 0   # 세로 헤더=고정 열 폭 합 / 가로 헤더=고정 행 높이 합

    def set_band(self, px: int):
        px = max(0, int(px))
        if px != self._band:
            self._band = px
            self.updateGeometry()

    def natural_size(self) -> QSize:
        """밴드를 뺀 헤더 본래 크기 — 헤더 재배치 계산은 반드시 이 값을 쓴다.
        width()/height() 는 super().updateGeometries() 가 '밴드 포함' 크기로
        세팅해 둔 값이라 그대로 쓰면 밴드가 이중 반영된다."""
        return QHeaderView.sizeHint(self)

    def sizeHint(self) -> QSize:
        sh = QHeaderView.sizeHint(self)
        if self._band <= 0:
            return sh
        if self.orientation() == Qt.Vertical:
            return QSize(sh.width() + self._band, sh.height())
        return QSize(sh.width(), sh.height() + self._band)

class ExcelTableView(QTableView):
    stage_requested   = pyqtSignal(str)   # direction: 'a_to_b' | 'b_to_a'
    unstage_requested = pyqtSignal()
    key_col_changed   = pyqtSignal(int)   # 키 열 변경 요청
    key_row_changed   = pyqtSignal(int)   # 키 행(헤더 행) 변경 요청
    columns_exclude_set = pyqtSignal(list, bool)   # (cols, exclude) — True: 제외 추가, False: 제외 해제
    column_resized    = pyqtSignal(int, int)   # (col, new_width) — 사용자 조작에 의한 변경만
    row_resized       = pyqtSignal(int, int)   # (row, new_height) — 사용자 조작에 의한 변경만
    drag_selection_finished = pyqtSignal()     # 드래그 선택 종료 — A↔B 미러를 1회만 수행

    def __init__(self, side: str, parent=None):
        super().__init__(parent)
        # 틀 고정 — updateGeometries() 오버라이드가 setModel 등에서 조기 호출될 수 있어 먼저 초기화.
        self._freeze = None            # FreezeController(생성 시 자기 자신을 여기 설정)
        self._in_update_geoms = False  # updateGeometries 재진입 가드
        self.side = side
        # 헤더 교체는 setModel 및 아래 헤더 설정(리사이즈 모드/아이콘/시그널/이벤트 필터)
        # 보다 먼저. 나중에 바꾸면 그 설정들이 버려지는 옛 헤더에 적용된다.
        for _hdr, _setter in ((_BandHeaderView(Qt.Horizontal, self), self.setHorizontalHeader),
                              (_BandHeaderView(Qt.Vertical, self), self.setVerticalHeader)):
            # ★ QTableView 는 이 두 속성을 **자기가 만든 기본 헤더에만** 켠다
            #   (QTableViewPrivate::init — setHorizontalHeader/setVerticalHeader 안이 아니다).
            #   헤더를 교체하면서 이걸 빠뜨리면 sectionsClickable 이 False 라 헤더 좌클릭이
            #   sectionPressed 를 아예 안 쏜다 → ① 열/행 헤더를 클릭해도 선택이 안 되고,
            #   ② _header_anchor_col/_row 가 갱신되지 않아 살짝만 끌어도 **낡은 앵커**부터
            #   현재 열까지 통째로 선택된다(F 를 눌렀는데 A~F 가 선택되던 증상).
            _hdr.setSectionsClickable(True)
            _hdr.setHighlightSections(True)
            _setter(_hdr)
        self._model = DiffTableModel(side, self)
        self.setModel(self._model)   # selectionModel은 여기서 1회 생성 — 이후 교체 없음
        # A/B 값이 다른 문자 구간을 셀 안에서 핑크로 강조 (modified 셀 한정)
        self.setItemDelegate(DiffHighlightDelegate(self))
        self._populating = False
        self._key_col: int = 0
        self._key_row: int = 0
        self._excluded_cols: set[int] = set()   # 변경 검사 제외 열 (display 인덱스)
        # 사용자가 직접 조정한 열/행 크기 — 세션 동안만 유지 (재로드/저장/새로고침 후 복원)
        self._user_col_widths: dict[int, int] = {}
        self._user_row_heights: dict[int, int] = {}
        # 외부(다른 패널)에서 크기를 강제 적용 중일 때 sectionResized 재방출 방지
        self._applying_sizes: bool = False
        # 키 열/행 보충(_supplement_key_selection) 중 selectionChanged 재진입 방지
        self._supplementing: bool = False
        # 헤더 다중 선택의 anchor (Shift+방향 확장의 고정점)
        self._header_anchor_col: int | None = None
        self._header_anchor_row: int | None = None
        # 본체에서 시작한 드래그가 왼쪽 고정 열 밴드로 넘어갈 때의 앵커 (row, col) 모델 좌표.
        self._body_drag_anchor = None
        self._body_drag_custom = False   # 이번 드래그에서 경계 넘기 커스텀 선택을 만들었는가
        # 헤더 드래그 선택(경계 넘기): 이벤트 필터를 건 헤더 → 축('col'/'row'), 진행 중 축.
        self._header_axis = {}
        self._header_drag_axis = None
        # 헤더 드래그 스로틀: 마지막으로 선택을 적용한 대상 열/행. 한 열/행 위에서 픽셀이
        # 움직일 때마다 같은 범위를 다시 선택하는 낭비를 없앤다(Qt 는 동일 선택이면
        # selectionChanged 를 안 쏘므로 큰 이득은 아니지만, 파이썬 작업 자체를 줄인다).
        self._header_drag_target = None
        # 드래그(셀/헤더/오버레이) 진행 중 표시. 진행 중에는 A↔B 선택 미러를 보류하고
        # 종료 시 1회만 미러한다 — 미러가 드래그 지연의 지배적 원인이었다(측정: 실제 변화가
        # 생기는 선택 갱신 24.6ms 중 약 22ms 가 미러, 미러 지연만으로 7.6x 개선).
        self._drag_selecting = False
        self.setFont(ui_font(9))
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        # 키/제외 열 헤더의 PNG 아이콘(DecorationRole) 가시성 확보용 크기.
        self.horizontalHeader().setIconSize(QSize(14, 14))
        self.verticalHeader().setDefaultSectionSize(22)
        # 키 행 헤더의 PNG 아이콘(DecorationRole) 가시성 확보용 크기(가로와 동일).
        self.verticalHeader().setIconSize(QSize(14, 14))
        self.setAlternatingRowColors(False)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        # 선택색 통일: 포커스가 없어도(버튼/단축키로 이동 시) 비활성 하이라이트가
        # 회색으로 흐려지지 않고 활성(클릭 선택)과 같은 파랑으로 보이게 한다.
        force_active_highlight(self)
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
        # 본체 헤더 드래그 선택 필터(고정 밴드 경계 넘기). 헤더의 viewport에 필터를 건다.
        # corner 헤더는 FreezeController가 register_frozen_headers 로 추가 등록한다.
        for _h, _ax in ((self.horizontalHeader(), "col"), (self.verticalHeader(), "row")):
            self._header_axis[_h.viewport()] = (_ax, _h)
            _h.viewport().installEventFilter(self)
        # 선택 정규화: '변경 행만 보기'로 숨긴(볼 수 없는) 행을 전폭 밴드 선택에서 제외.
        # (DiffView의 A↔B 선택 미러보다 먼저 연결돼야 미러 전에 정규화가 반영된다.)
        self.selectionModel().selectionChanged.connect(self._normalize_selection)

    # ── QTableWidget 호환 헬퍼 ───────────────────────────────────────────────
    def rowCount(self) -> int:
        return self._model.rowCount()

    def columnCount(self) -> int:
        return self._model.columnCount()

    def _current_cell(self) -> tuple:
        """QTableWidget.currentRow()/currentColumn() 대응 — 무효 시 (-1, -1)."""
        idx = self.currentIndex()
        return (idx.row(), idx.column()) if idx.isValid() else (-1, -1)

    def _scroll_hidden_cursor_into_view(self, r: int, c: int):
        """고정 밴드 열 위에 놓인 커서를 따라 **세로** 스크롤을 맞춘다.

        QTableView.scrollTo 는 isIndexHidden 인 인덱스를 통째로 무시한다. 틀 고정이
        켜지면 키 밴드 열은 본체에서 숨김이므로, 밴드 안에서 ↑/↓ 로 행을 옮겨도
        뷰포트가 따라오지 않아 커서가 화면 밖으로 걸어나간다. 같은 행의 '본체에 보이는'
        열로 대신 스크롤해 세로만 맞추고, 가로 위치는 원래대로 되돌린다.
        """
        if AX_ROW.hidden(self, r) or not AX_COL.hidden(self, c):
            return   # 고정 밴드 '행'은 오버레이에 늘 보이고, 보이는 열은 Qt 가 처리한다
        helper = self.columnAt(0)
        if helper < 0 or self.isColumnHidden(helper):
            helper = next((k for k in range(self.columnCount())
                           if not self.isColumnHidden(k)), -1)
        if helper < 0:
            return
        hbar = self.horizontalScrollBar()
        hv = hbar.value()
        self.scrollTo(self._model.index(r, helper))
        hbar.setValue(hv)

    def _set_current_cell(self, r: int, c: int):
        """QTableWidget.setCurrentCell() 대응.
        setCurrentIndex()는 호출 시점의 키보드 수정자에 따라 선택을 확장/클리어하는
        기존 setCurrentCell과 완전히 같은 경로(selectionCommand)를 탄다."""
        if 0 <= r < self.rowCount() and 0 <= c < self.columnCount():
            self.setCurrentIndex(self._model.index(r, c))
            self._scroll_hidden_cursor_into_view(r, c)

    def _set_current_cell_no_update(self, r: int, c: int):
        """선택을 건드리지 않고 currentIndex만 이동.
        헤더/범위 선택 직후에 사용 — setCurrentIndex()는 눌려 있는 Shift를 보고
        SelectCurrent(앵커 사각형으로 선택 대체)를 적용해 방금 만든 범위
        선택을 붕괴시키므로, NoUpdate로 현재 셀만 옮긴다."""
        sm = self.selectionModel()
        if sm is not None and 0 <= r < self.rowCount() and 0 <= c < self.columnCount():
            sm.setCurrentIndex(self._model.index(r, c), QItemSelectionModel.NoUpdate)
            self._scroll_hidden_cursor_into_view(r, c)

    def _move_current_cell(self, r: int, c: int):
        """현재 셀을 단일 선택으로 이동 — Excel의 Ctrl+점프처럼 기존 선택을 비운다.
        setCurrentIndex()는 Ctrl이 눌린 상태에서 Toggle로 동작해 원래 셀 선택이
        남으므로 ClearAndSelect를 명시한다."""
        sm = self.selectionModel()
        if sm is not None and 0 <= r < self.rowCount() and 0 <= c < self.columnCount():
            sm.setCurrentIndex(self._model.index(r, c),
                               QItemSelectionModel.ClearAndSelect)
            self._scroll_hidden_cursor_into_view(r, c)

    # ── 본체→왼쪽 고정 열 밴드로 넘어가는 드래그 선택 ─────────────────────────
    # 본체에서 시작한 드래그가 왼쪽 고정 열(키 열 및 그 좌측)로 넘어가면, 본체엔 고정 열이
    # 숨겨져 있어 Qt 기본 드래그가 고정 열을 잡지 못한다(예: 본체에서 K→E 드래그 후 D로 못 이어짐).
    # 오버레이가 본체와 selection model 을 공유하므로, 커서가 본체 뷰포트 왼쪽을 벗어나면
    # 앵커→고정열 사각형으로 직접 확장한다. (반대 방향 = 오버레이→본체 는 _FrozenView가 처리.)
    def _end_drag_select(self):
        """드래그 선택 종료 — 진행 중이었다면 A↔B 미러를 1회 수행하도록 신호한다.

        본체/오버레이/헤더의 **모든 release 경로**에서 불려야 한다. 한 곳이라도 빠지면
        플래그가 켜진 채 남아 그 뒤의 선택이 반대 패널에 미러되지 않는다.
        """
        if self._drag_selecting:
            self._drag_selecting = False
            self.drag_selection_finished.emit()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            idx = self.indexAt(event.pos())
            self._body_drag_anchor = (idx.row(), idx.column()) if idx.isValid() else None
            self._body_drag_custom = False
            self._drag_selecting = True   # 드래그 중 미러 억제(종료 시 1회)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        # Ctrl/Shift 드래그는 Qt 기본에 맡겨 다중 영역/확장 시맨틱을 보존한다(_FrozenView와 동일).
        if (self._body_drag_anchor is not None and (event.buttons() & Qt.LeftButton)
                and not (event.modifiers() & (Qt.ControlModifier | Qt.ShiftModifier))):
            fc = getattr(self, "_freeze", None)
            bv = self.viewport().mapFromGlobal(event.globalPos())
            if (bv.x() < 0 and fc is not None and getattr(fc, "active", False)
                    and getattr(fc, "_n_cols", 0) > 0):
                col = self._frozen_col_at(fc, event.globalPos())
                if col is not None:
                    row = self._drag_row_at(bv.y())   # 아래 끝 넘음/빈 영역에서도 마지막 행으로 clamp
                    ar, ac = self._body_drag_anchor
                    self._select_range(ar, ac, row, col)
                    self._set_current_cell_no_update(row, col)
                    self._body_drag_custom = True
                    event.accept()
                    return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        # 커스텀 경계-넘기 드래그로 이동을 소비했으면 Qt는 드래그 상태로 진입하지 못한 채라,
        # super()의 릴리즈가 이를 '클릭'으로 보고 선택을 리셋해 방금 만든 선택을 지운다
        # (_FrozenView와 동일 이유). 커스텀 드래그였다면 super()를 건너뛰고 상태만 정리한다.
        if self._body_drag_custom:
            self._body_drag_custom = False
            self._body_drag_anchor = None
            self.setState(QAbstractItemView.NoState)
            event.accept()
            self._end_drag_select()
            return
        self._body_drag_anchor = None
        super().mouseReleaseEvent(event)
        self._end_drag_select()

    @staticmethod
    def _frozen_axis_at(ax: _Axis, fc, gp):
        """전역좌표 gp 아래의 고정 열/행 인덱스. 밴드 앞(행 헤더 등)이면 첫 고정 섹션(0),
        끝을 넘으면 마지막 고정 섹션으로 clamp.

        이 축의 고정 밴드를 **실제로 그리는** 오버레이(열=left, 행=top) 기준으로 매핑한다 —
        고정 섹션이 숨김인 본체/반대 오버레이로 매핑하면 좌표가 틀어지기 때문."""
        ov = ax.overlay(fc)
        p = ov.viewport().mapFromGlobal(gp)
        v = ax.pos(p)
        s = ax.section_at(ov, v)
        if s < 0:
            s = 0 if v < 0 else ax.frozen_count(fc) - 1
        return s

    @staticmethod
    def _frozen_col_at(fc, gp):
        return ExcelTableView._frozen_axis_at(AX_COL, fc, gp)

    @staticmethod
    def _frozen_row_at(fc, gp):
        return ExcelTableView._frozen_axis_at(AX_ROW, fc, gp)

    def _drag_col_at(self, x):
        """드래그 확장용: 본체 뷰포트 x 아래의 목표 데이터 열. 마지막 열보다 오른쪽(열이 뷰포트를
        안 채운 빈 영역이거나 우측 끝을 넘음)이면 마지막 데이터 열로 clamp(앵커로 붕괴 방지),
        뷰포트 밖(우측)이면 한 칸 오토스크롤. EXTRA 빈 열은 데이터 범위로 clamp한다."""
        dc = self._model.data_cols
        last = max(0, dc - 1)
        c = self.columnAt(x)
        if c >= 0:
            return min(c, last)
        if x >= self.viewport().width():
            hb = self.horizontalScrollBar()
            if hb.value() < hb.maximum():
                hb.setValue(hb.value() + 1)
                c = self.columnAt(max(0, self.viewport().width() - 1))
                if c >= 0:
                    return min(c, last)
        return last

    def _drag_row_at(self, y):
        """_drag_col_at 의 세로 대칭. 위쪽(고정 행/헤더) 넘으면 첫 행, 뷰포트 밖(아래)이면
        한 칸 오토스크롤, 행이 없는 빈 영역이면 마지막 행으로 clamp.

        **커서 위에 실제로 행이 있으면 그 행을 그대로 쓴다** — EXTRA(빈) 행이어도 마찬가지.
        예전에는 데이터 범위(data_rows-1)로 clamp 했는데, '변경점만 보기' ON 이면 보이는 행
        대부분이 EXTRA 행이라(예: 24개 중 20개) 그 위에서 드래그할 때 목표가 항상 마지막
        데이터 행으로 접혀 **선택이 커서를 따라가지 않고, 드래그하지도 않은 위쪽 행이 선택**
        되는 문제가 있었다. 엑셀처럼 드래그한 범위가 그대로 선택되게 한다(EXTRA 행은 diff
        상태가 없어 병합 준비에서 자동 제외되므로 병합 동작에는 영향이 없다).
        """
        last = max(0, self.rowCount() - 1)
        if y < 0:
            return 0
        r = self.rowAt(y)
        if r >= 0:
            return min(r, last)
        if y >= self.viewport().height():
            vb = self.verticalScrollBar()
            if vb.value() < vb.maximum():
                vb.setValue(vb.value() + 1)
                r = self.rowAt(max(0, self.viewport().height() - 1))
                if r >= 0:
                    return min(r, last)
        return last

    def _col_under_global(self, gp):
        """전역 커서 gp 아래의 목표 데이터 열 — 고정 밴드(고정 열 오버레이)·본체 스크롤·끝 clamp를
        통합한다. 헤더 드래그(가로)에서 커서가 헤더 위여도 열은 x만으로 결정되므로 그대로 쓴다."""
        bvx = self.viewport().mapFromGlobal(gp).x()
        if bvx >= 0:
            return self._drag_col_at(bvx)
        fc = getattr(self, "_freeze", None)
        if fc is not None and getattr(fc, "_n_cols", 0) > 0:
            return self._frozen_col_at(fc, gp)
        c = self.columnAt(0)
        return c if c >= 0 else 0

    def _row_under_global(self, gp):
        """_col_under_global 의 세로 대칭 (고정 행 오버레이·본체·끝 clamp 통합)."""
        bvy = self.viewport().mapFromGlobal(gp).y()
        if bvy >= 0:
            return self._drag_row_at(bvy)
        fc = getattr(self, "_freeze", None)
        if fc is not None and getattr(fc, "_n_rows", 0) > 0:
            return self._frozen_row_at(fc, gp)
        return 0

    # ── 헤더 드래그 선택 (고정 밴드 경계를 넘는 열/행 헤더 드래그) ──────────────
    # 고정 열(키 열 이하) 헤더는 corner 오버레이, 스크롤 열 헤더는 본체에 있어 서로 다른 위젯이라
    # Qt 기본 헤더 드래그는 경계를 못 넘는다(K→A/A→K 실패, 키 행 방향에 따라 10→1 실패 등).
    # 4개 헤더(본체·corner의 가로/세로)에 이벤트 필터를 걸어, 드래그 중 전역 커서를 열/행로
    # 매핑(_col_under_global/_row_under_global)해 앵커→목표 범위를 공유 모델에 직접 선택한다.
    def register_frozen_headers(self, corner):
        """FreezeController가 corner 오버레이 헤더를 host의 드래그 선택 필터/anchor에 등록."""
        for h, ax in ((corner.horizontalHeader(), "col"), (corner.verticalHeader(), "row")):
            # 필터는 헤더의 viewport에 건다 — QHeaderView 마우스 이벤트는 viewport로 전달되므로
            # 헤더 객체 자체에 걸면 실제 드래그에서 필터가 호출되지 않는다.
            self._header_axis[h.viewport()] = (ax, h)
            h.viewport().installEventFilter(self)
        # corner 헤더 클릭도 anchor를 세팅하도록(본체 헤더와 동일 배선)
        corner.horizontalHeader().sectionPressed.connect(self._on_h_section_pressed)
        corner.verticalHeader().sectionPressed.connect(self._on_v_section_pressed)

    @staticmethod
    def _near_section_boundary(header, axis, pos):
        """pos(헤더 로컬)가 리사이즈 핸들 위인가 — 그러면 드래그 선택을 시작하지 않고 Qt 기본
        (열/행 크기 조절)에 맡긴다. Qt가 그립 위에서 세팅하는 split 커서를 우선 신뢰하고(그립
        폭이 스타일/DPI마다 달라도 정확), 못 잡으면 섹션 경계 6px 여백으로 보수적으로 판단한다."""
        if header.cursor().shape() in (Qt.SplitHCursor, Qt.SplitVCursor):
            return True
        p = pos.x() if axis == "col" else pos.y()
        idx = header.logicalIndexAt(p)
        if idx < 0:
            return False
        start = header.sectionViewportPosition(idx)
        size = header.sectionSize(idx)
        # 고정 6px 은 얇은 섹션을 통째로 삼켰다: 행 높이 22px 에서 앞뒤 6px = 13px(59%)이
        # '리사이즈'로 판정돼 드래그 선택이 절반 이상 먹지 않았다(실측). Qt 가 실제로 쓰는
        # 그립 폭(PM_HeaderGripMargin, 보통 4px)에 맞추고, 얇은 섹션에서는 그 1/3 이하로
        # 캡해 섹션 전체가 그립이 되는 일을 막는다. 리사이즈 자체는 그대로 동작한다.
        grip = QApplication.style().pixelMetric(QStyle.PM_HeaderGripMargin) or 4
        m = max(1, min(grip, max(1, size // 3)))
        return (p - start) < m or (start + size - p) <= m

    def eventFilter(self, obj, event):
        info = getattr(self, "_header_axis", {}).get(obj)
        if info is not None:
            axis, header = info
            et = event.type()
            if et == QEvent.MouseButtonPress:
                # 좌클릭·리사이즈 아님·modifier 없음일 때만 우리가 드래그 선택을 관리한다.
                # (press는 소비하지 않는다 — Qt가 grab/초기선택/anchor(sectionPressed)를 처리)
                if (event.button() == Qt.LeftButton
                        and not self._near_section_boundary(header, axis, event.pos())
                        and not (event.modifiers() & (Qt.ControlModifier | Qt.ShiftModifier))):
                    self._header_drag_axis = axis
                else:
                    self._header_drag_axis = None
                self._header_drag_target = None    # 새 드래그 — 스로틀 기준 초기화
                if event.button() == Qt.LeftButton:
                    self._drag_selecting = True    # 드래그 중 미러 억제(종료 시 1회)
            elif et == QEvent.MouseMove:
                if self._header_drag_axis == axis and (event.buttons() & Qt.LeftButton):
                    # 스로틀: 대상 열/행이 바뀔 때만 재선택한다.
                    if axis == "col" and self._header_anchor_col is not None:
                        tgt = self._col_under_global(event.globalPos())
                        if (tgt is not None and tgt >= 0
                                and tgt != self._header_drag_target):
                            self._header_drag_target = tgt
                            self._select_column_range(self._header_anchor_col, tgt)
                            self._set_current_cell_no_update(
                                max(0, self._current_cell()[0]), tgt)
                        return True   # Qt 기본 드래그(경계 내 clamp) 대신 우리 선택 사용
                    if axis == "row" and self._header_anchor_row is not None:
                        tgt = self._row_under_global(event.globalPos())
                        if (tgt is not None and tgt >= 0
                                and tgt != self._header_drag_target):
                            self._header_drag_target = tgt
                            self._select_row_range(self._header_anchor_row, tgt)
                            self._set_current_cell_no_update(
                                tgt, max(0, self._current_cell()[1]))
                        return True
            elif et == QEvent.MouseButtonRelease:
                self._header_drag_axis = None
                self._header_drag_target = None
                self._end_drag_select()
        return super().eventFilter(obj, event)

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
        fc = getattr(self, "_freeze", None)
        if fc is not None and getattr(fc, "active", False) and col < fc._n_cols:
            # 고정(키) 열: 본체에서 숨겨져 있어 특수 경로로 동기(미러 재방출은 안 함)
            fc._apply_frozen_col_width(col, width, mirror=False)
            return
        if 0 <= col < self.columnCount() and self.columnWidth(col) != width:
            self._applying_sizes = True
            try:
                self.setColumnWidth(col, width)
            finally:
                self._applying_sizes = False
            self._resync_freeze_sizes()

    def apply_row_height(self, row: int, height: int):
        """반대 패널에서의 행 높이 변경을 동기 적용."""
        self._user_row_heights[row] = height
        fc = getattr(self, "_freeze", None)
        if fc is not None and getattr(fc, "active", False) and row < fc._n_rows:
            fc._apply_frozen_row_height(row, height, mirror=False)
            return
        if 0 <= row < self.rowCount() and self.rowHeight(row) != height:
            self._applying_sizes = True
            try:
                self.setRowHeight(row, height)
            finally:
                self._applying_sizes = False
            self._resync_freeze_sizes()

    def _resync_freeze_sizes(self):
        """틀 고정 헬퍼 뷰의 열폭/행높이를 본체와 다시 맞추고 재배치.
        apply_column_width/apply_row_height 는 _applying_sizes 가드 아래서 크기를 적용하는데,
        이 가드는 FreezeController._on_col_resized/_on_row_resized 까지 억제한다(호스트의
        sectionResized 로 트리거됨). 그 결과 반대 패널의 상단/좌측 고정 밴드가 새 열폭/행높이를
        반영하지 못해 본체와 어긋난다 → 여기서 명시적으로 밴드를 재동기한다."""
        fc = getattr(self, "_freeze", None)
        if fc is not None and getattr(fc, "active", False):
            fc._sync_sizes()
            self.updateGeometries()

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
        # 헤더 아이콘(키 열 열쇠·제외 열 표시) 자리. 예전엔 열 문자 폭만 재서 아이콘이
        # 글자를 밀어냈고, 키 열은 열쇠에 가려 열 문자가 잘려 보였다(🔑A → 🔑!).
        icon_w = self.horizontalHeader().iconSize().width() + 6
        self._applying_sizes = True
        try:
            for c in range(cols):
                # 헤더에 **실제로 보이는 글자**를 잰다. 열을 이름으로 맞춘 비교에서는
                # 패널마다 열 문자가 다르고, 그 파일에 없는 열은 '-' 다.
                txt = self._model.headerData(c, Qt.Horizontal, Qt.DisplayRole)
                txt = str(txt) if txt else get_column_letter(c + 1)
                hdr_min = hdr_fm.horizontalAdvance(txt) + 24
                if self._model.headerData(c, Qt.Horizontal, Qt.DecorationRole) is not None:
                    hdr_min += icon_w
                w = hdr_min
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
                # 상한을 씌우되 헤더가 잘릴 만큼 줄이지는 않는다 — 데이터가 길어도
                # 열 문자와 아이콘은 보여야 한다.
                self.setColumnWidth(c, max(min(w, MAX_AUTO_COL_WIDTH_PX), hdr_min))
        finally:
            self._applying_sizes = False

    def set_key_col(self, col: int):
        self._key_col = col
        self._model.set_key_col(col)

    def set_key_row(self, row: int):
        self._key_row = row
        self._model.set_key_row(row)

    def updateGeometries(self):
        """틀 고정 활성 시 상단/좌측에 고정 헬퍼 자리를 예약하고 본체 데이터 뷰포트를
        그만큼 밀어낸다(가림 방지). Qt가 지오메트리를 재계산할 때마다(리사이즈·헤더폭
        변화·스크롤·스플리터) 예약을 다시 적용하고 헬퍼 위치를 갱신한다.

        ★ 예약은 setViewportMargins 가 아니라 **헤더 sizeHint**(_BandHeaderView.set_band)로
        한다. QTableView.updateGeometries() 는 ① 헤더 sizeHint 로 여백을 잡은 뒤
        ② 그 뷰포트 크기로 스크롤바 range 를 계산하므로, 여백을 super() 뒤에 더하면
        range 가 고정 밴드 폭만큼 과소평가된다. 그 결과 '넘치는 폭 < 밴드 폭'이면
        range 가 0 이 되어 **가로 스크롤바가 아예 안 뜨고 오른쪽 열을 못 본다**
        (세로도 고정 행 높이만큼 동일한 오차). 밴드를 sizeHint 에 미리 넣으면 Qt 가
        여백과 range 를 한 번에 정확히 계산한다.
        """
        if self._in_update_geoms:
            # 재진입(헤더 setGeometry → geometriesChanged 등) — 한 패스에서 두 번
            # 계산하지 않도록 무시한다.
            return
        self._in_update_geoms = True
        try:
            fc = self._freeze
            vh, hh = self.verticalHeader(), self.horizontalHeader()
            # ctor 에서 헤더를 교체하기 전에도 Qt 가 이 함수를 부른다(기본 QHeaderView).
            # 그땐 틀 고정도 없으므로 밴드 처리 없이 super() 만 태운다.
            if not isinstance(vh, _BandHeaderView) or not isinstance(hh, _BandHeaderView):
                super().updateGeometries()
                return
            active = fc is not None and fc.active
            # 벌크(_applying_sizes: 필터의 setRowHidden 폭풍/크기 일괄 적용) 중엔 무거운 freeze
            # 작업(밴드 재계산·헤더 재배치·reposition)을 건너뛴다 — 벌크가 끝난 뒤 refresh가
            # 1회만 재적용. 이때 밴드 값은 건드리지 않아 예약된 여백이 그대로 유지된다.
            do_freeze = active and not self._applying_sizes
            if do_freeze:
                fw, fh = fc.frozen_px()
                vh.set_band(fw)
                hh.set_band(fh)
            elif not active:
                vh.set_band(0)
                hh.set_band(0)
            super().updateGeometries()
            if do_freeze:
                fr = self.frameWidth()
                # 밴드를 뺀 '본래' 헤더 크기. width()/height()는 super()가 밴드 포함
                # 크기로 세팅해 둔 값이라 쓰면 밴드가 이중 반영된다.
                hdr_w = vh.natural_size().width()
                hdr_h = hh.natural_size().height()
                # 본체 헤더를 밀어낸 뷰포트에 맞춰 재배치 — 그래야 고정 눈금(corner) 옆에
                # 스크롤 눈금(본체 헤더)이 정확히 이어진다. QTableView는 헤더를 밴드 포함
                # 크기로 배치하므로 수동 정렬한다.
                vp = self.viewport().geometry()
                hh.setGeometry(fr + hdr_w + fw, fr, vp.width(), hdr_h)
                vh.setGeometry(fr, fr + hdr_h + fh, hdr_w, vp.height())
                fc.reposition()
        finally:
            self._in_update_geoms = False

    def set_excluded_cols(self, cols: set):
        """외부(MainWindow)에서 제외 열 집합을 갱신하고 헤더/셀을 다시 칠한다."""
        self._excluded_cols = set(cols)
        self._model.set_excluded_cols(self._excluded_cols)

    # ── 축 대칭 연산 (열/행 공용 구현 + 축만 넘기는 위임) ──────────────────────
    # 알고리즘은 _axis_* 하나뿐이다. 열/행 중 한쪽만 고쳐져 생기던 축 불일치 버그를
    # 구조적으로 막기 위한 것이며, 기존 이름은 호출부·테스트 호환을 위해 유지한다.
    def _axis_has_changed(self, ax: _Axis, i: int) -> bool:
        """지정 열/행에 changed 셀이 있는가 — 첫 changed에서 조기 종료."""
        m = self._model
        return any(ax.kind(m, i, x) == "changed"
                   for x in range(ax.cross_data_count(m)))

    def _axis_have_staged(self, ax: _Axis, items) -> bool:
        """대상 열/행들 중 하나라도 staged 셀을 포함하는가 — O(#staged)."""
        want = set(items)
        return any(ax.of_staged(coord) in want
                   for coord in self._model.staged_coords())

    def _select_axis(self, ax: _Axis, i: int) -> None:
        """해당 열/행 전체 셀을 선택 상태로 설정."""
        sm = self.selectionModel()
        n = ax.count(self)
        cross = ax.cross_count(self)
        if sm is None or cross == 0 or n == 0 or not (0 <= i < n):
            return
        model = self.model()
        sel = QItemSelection(ax.index(model, i, 0), ax.index(model, i, cross - 1))
        sm.select(sel, QItemSelectionModel.ClearAndSelect)

    def _select_axis_multi(self, ax: _Axis, items) -> None:
        """여러 열/행 전체 셀을 한 번에 선택 (비연속 지원)."""
        sm = self.selectionModel()
        cross = ax.cross_count(self)
        i_max = ax.count(self) - 1
        if sm is None or cross == 0 or i_max < 0:
            return
        model = self.model()
        sel = QItemSelection()
        for i in items:
            if 0 <= i <= i_max:
                sel.append(QItemSelectionRange(
                    ax.index(model, i, 0), ax.index(model, i, cross - 1)))
        sm.select(sel, QItemSelectionModel.ClearAndSelect)

    def _col_has_changed(self, col: int) -> bool:
        return self._axis_has_changed(AX_COL, col)

    def _row_has_changed(self, row: int) -> bool:
        return self._axis_has_changed(AX_ROW, row)

    def _cols_have_staged(self, cols) -> bool:
        return self._axis_have_staged(AX_COL, cols)

    def _rows_have_staged(self, rows) -> bool:
        return self._axis_have_staged(AX_ROW, rows)

    def _select_col(self, col: int):
        self._select_axis(AX_COL, col)

    def _select_cols(self, cols) -> None:
        """여러 열 전체 셀을 한 번에 선택 — 다중 열 병합 준비용."""
        self._select_axis_multi(AX_COL, cols)

    def _select_row(self, row: int):
        self._select_axis(AX_ROW, row)

    def _select_rows(self, rows) -> None:
        self._select_axis_multi(AX_ROW, rows)

    # 전폭 행-밴드에서 '변경 행만 보기'로 숨긴 데이터 행을 한 번에 제거할지 판단하는 상한.
    # 조각(range) 폭주로 페인팅/질의가 느려지는 걸 막는다 — 넘으면 그 밴드는 건드리지 않는다.
    _MAX_PRUNE_ROWS = 2000

    def _normalize_selection(self, *_):
        """선택 정규화 — selectionChanged 마다 1회, _supplementing 가드로 재진입 차단.
        엑셀식 원칙: '터치한 셀만' 선택. 선택에 셀을 '더하지' 않는다(키 열/행 자동 보충 없음).
        유일한 정규화는 '변경 행만 보기'로 숨긴, 사용자가 실제로 볼 수 없는 행을 전폭 행-밴드
        선택에서 걷어내는 것뿐이다(보이지 않는 행은 '터치'로 볼 수 없으므로).
        미러(mirror_selection/mirror_selection_from)는 _populating=True 라 여기서 걸러진다."""
        if self._populating or self._supplementing or self._applying_sizes:
            return
        sm = self.selectionModel()
        if sm is None:
            return
        n_rows, n_cols = self.rowCount(), self.columnCount()
        if n_rows == 0 or n_cols == 0:
            return
        self._supplementing = True
        try:
            self._prune_filtered_rows(sm, n_rows, n_cols)
        finally:
            self._supplementing = False

    def _prune_filtered_rows(self, sm, n_rows: int, n_cols: int):
        """'변경 행만 보기'로 숨긴 데이터 행이 전폭 행-밴드 선택에 끼어들면 선택에서 제거한다.
        - 전폭(모든 열: left==0 && right==col_max) 행-밴드만 대상 — 행 헤더 클릭/Shift 범위가 만든다.
          임의 셀 블록(전폭 아님)·데이터 열 선택은 사용자가 잡은 그대로 둔다.
        - 키 프레임 행(0..key_row)은 틀 고정으로 항상 숨김이지만 '보이는' 것으로 취급(상단 고정
          밴드에 표시되므로) → 건드리지 않는다.
        - 조각(range) 폭주를 막기 위해 한 밴드에서 제거할 숨김 행이 _MAX_PRUNE_ROWS 를 넘으면
          그 밴드는 건너뛴다(전폭 selectAll·초대형 범위 대비)."""
        row_max, col_max = n_rows - 1, n_cols - 1
        kr = self._key_row if (self._key_row and self._key_row > 0) else 0
        model = self.model()
        dead = QItemSelection()
        for rng in sm.selection():
            if rng.left() != 0 or rng.right() != col_max:
                continue   # 전폭 행-밴드만 (셀 블록/열 선택 제외)
            top, bot = rng.top(), rng.bottom()
            if top == 0 and bot == row_max:
                continue   # 전체 높이(selectAll/전 시트)는 그대로 — 조각화 방지, 행 헤더 밴드 아님
            band = [r for r in range(top, bot + 1)
                    if r > kr and self.isRowHidden(r)]
            if not band or len(band) > self._MAX_PRUNE_ROWS:
                continue
            for r in band:
                dead.append(QItemSelectionRange(
                    model.index(r, 0), model.index(r, col_max)))
        if not dead.isEmpty():
            sm.select(dead, QItemSelectionModel.Deselect)

    def _axis_touched(self, ax: _Axis) -> set[int]:
        """선택 range가 닿은 모든 열/행(부분 선택 포함). O(#range × 평균 폭)."""
        sm = self.selectionModel()
        out: set[int] = set()
        if sm is not None:
            for rng in sm.selection():
                out.update(range(ax.lo(rng), ax.hi(rng) + 1))
        return out

    def _axis_selected_header(self, ax: _Axis, anchor: int) -> list[int]:
        """헤더 우클릭 시 대상 열/행 집합 결정.
        - 우클릭한 열/행이 현재 다중 선택에 포함되어 있으면 그 선택 전체.
        - 아니면 우클릭한 단일 열/행만.
        range 기반 — 전체 열/행 선택 우선, 없으면 닿은 집합(selectedIndexes 폴백과 동일 의미)."""
        items = set(ax.full_selected(self)) or ax.touched(self)
        if anchor in items and len(items) > 1:
            return sorted(items)
        return [anchor]

    def _touched_rows(self) -> set[int]:
        return self._axis_touched(AX_ROW)

    def _touched_cols(self) -> set[int]:
        return self._axis_touched(AX_COL)

    def _selected_header_rows(self, anchor_row: int) -> list[int]:
        return self._axis_selected_header(AX_ROW, anchor_row)

    def _selected_header_cols(self, anchor_col: int) -> list[int]:
        return self._axis_selected_header(AX_COL, anchor_col)

    def _show_header_context_menu(self, pos, header=None):
        # header: 클릭된 실제 가로 헤더(기본=본체). 틀 고정 corner 헤더에서도 호출될 수 있다.
        header = header or self.horizontalHeader()
        col = header.logicalIndexAt(pos)
        if col < 0:
            return

        target_cols = self._selected_header_cols(col)
        multi = len(target_cols) > 1
        col_letter = get_column_letter(col + 1)
        cols_label = ", ".join(get_column_letter(c + 1) for c in target_cols)

        # 병합 준비/취소는 단일·다중 모두 지원 — 대상 열들 중 하나라도 변경/스테이징 셀이 있으면 노출.
        # (변경 검사 제외 열도 cell_kind가 실제 status를 보고하므로 여기에 기여한다 — 제외 열만
        #  골라 병합 준비하는 것도 가능. 제외는 표시·집계 전용이고 병합은 명시적 지시다.)
        # has_changed는 열별 조기 종료(_col_has_changed) + any(), has_staged는 staged 집합 기반.
        has_changed = any(self._col_has_changed(c) for c in target_cols)
        has_staged  = self._cols_have_staged(target_cols)

        menu = QMenu(self)
        menu.setStyleSheet(MENU_QSS)

        # ── 병합 준비 항목 (단일·다중 공통) ──
        act_a2b = act_b2a = act_unstage = None
        if has_changed:
            act_a2b = menu.addAction(f"A → B  병합 준비  [{cols_label}열]")
            act_b2a = menu.addAction(f"B → A  병합 준비  [{cols_label}열]")
        if has_staged:
            if has_changed:
                menu.addSeparator()
            act_unstage = menu.addAction(f"병합 준비 취소  [{cols_label}열]")

        if has_changed or has_staged:
            menu.addSeparator()

        # ── 키 열 항목 (단일 선택일 때만) ──
        act_key_clear = act_key_set = None
        if not multi:
            if col == self._key_col:
                act_key_clear = menu.addAction("🔓  키 열 해제 (ROW 순서 기반 비교)")
            else:
                # 키 아이콘을 함께 노출해 키 지정 동작임을 시각적으로 표시.
                act_key_set = menu.addAction(
                    key_header_icon(), f"키 열로 설정  [{col_letter}열]")

        # ── 변경 검사 제외/해제 토글 ──
        # 키 열은 변경 검사에서 제외할 수 없으므로 대상에서 뺀다. 선택에 '제외됨'과 '비제외'가
        # 섞여 있으면 두 항목(제외 / 해제)을 모두 노출한다.
        exclude_cols = [c for c in target_cols if c != self._key_col]
        to_exclude = [c for c in exclude_cols if c not in self._excluded_cols]
        to_unexclude = [c for c in exclude_cols if c in self._excluded_cols]
        act_excl = act_unexcl = None
        if to_exclude or to_unexclude:
            menu.addSeparator()
        if to_exclude:
            lbl = ", ".join(get_column_letter(c + 1) for c in to_exclude)
            act_excl = menu.addAction(
                exclude_header_icon(), f"변경 검사에서 제외  [{lbl}열]")
        if to_unexclude:
            lbl = ", ".join(get_column_letter(c + 1) for c in to_unexclude)
            act_unexcl = menu.addAction(
                reset_header_icon(), f"검사 제외 해제  [{lbl}열]")

        act = menu.exec_(header.mapToGlobal(pos))
        if act is None:
            return

        if act_a2b is not None and act == act_a2b:
            self._select_cols(target_cols)
            self.stage_requested.emit(DIR_A2B)
        elif act_b2a is not None and act == act_b2a:
            self._select_cols(target_cols)
            self.stage_requested.emit(DIR_B2A)
        elif act_unstage is not None and act == act_unstage:
            self._select_cols(target_cols)
            self.unstage_requested.emit()
        elif act_key_clear is not None and act == act_key_clear:
            self.key_col_changed.emit(-1)
        elif act_key_set is not None and act == act_key_set:
            self.key_col_changed.emit(col)
        elif act_excl is not None and act == act_excl:
            self.columns_exclude_set.emit(to_exclude, True)
        elif act_unexcl is not None and act == act_unexcl:
            self.columns_exclude_set.emit(to_unexclude, False)

    def _show_row_header_context_menu(self, pos, header=None):
        # header: 클릭된 실제 세로 헤더(기본=본체). 틀 고정 corner 헤더에서도 호출될 수 있다.
        header = header or self.verticalHeader()
        row = header.logicalIndexAt(pos)
        if row < 0:
            return

        # 다중 선택된 행 헤더 전체를 대상으로 (열 헤더와 대칭)
        target_rows = self._selected_header_rows(row)
        multi = len(target_rows) > 1
        # has_changed는 행별 조기 종료(_row_has_changed) + any(), has_staged는 staged 집합 기반.
        has_changed = any(self._row_has_changed(r) for r in target_rows)
        has_staged  = self._rows_have_staged(target_rows)

        # 키 행(헤더 행) 지정: 단일 행이면서 원본 파일 행으로 매핑 가능한 데이터 행일 때만.
        # display 행은 키 매칭으로 재정렬될 수 있으므로 원본 파일 행 번호로 변환해 emit한다.
        orig_row = self._model.orig_row(row) if not multi else None
        can_key = orig_row is not None
        # 현재 키 행 우클릭 시엔 키 항목을 노출하지 않는다(초기화 기능 제거).
        # 키 행이 아닌 데이터 행에서만 '키 행으로 설정'을 노출.
        show_key_set = can_key and orig_row != self._key_row

        if not has_changed and not has_staged and not show_key_set:
            return

        suffix = f"  [{len(target_rows)}개 행]" if multi else ""
        menu = QMenu(self)
        menu.setStyleSheet(MENU_QSS)

        act_a2b = act_b2a = act_unstage = None
        if has_changed:
            act_a2b = menu.addAction(f"A → B  병합 준비{suffix}")
            act_b2a = menu.addAction(f"B → A  병합 준비{suffix}")
        if has_staged:
            if has_changed:
                menu.addSeparator()
            act_unstage = menu.addAction(f"병합 준비 취소{suffix}")

        # ── 키 행 항목: 키 행이 아닌 단일 데이터 행에만 '키 행으로 설정' 노출 ──
        act_key_row_set = None
        if show_key_set:
            if has_changed or has_staged:
                menu.addSeparator()
            act_key_row_set = menu.addAction(
                key_header_icon(), f"키 행으로 설정  [{orig_row + 1}행]")

        act = menu.exec_(header.mapToGlobal(pos))
        if act is None:
            return

        if act_a2b is not None and act == act_a2b:
            self._select_rows(target_rows)
            self.stage_requested.emit(DIR_A2B)
        elif act_b2a is not None and act == act_b2a:
            self._select_rows(target_rows)
            self.stage_requested.emit(DIR_B2A)
        elif act_unstage is not None and act == act_unstage:
            self._select_rows(target_rows)
            self.unstage_requested.emit()
        elif act_key_row_set is not None and act == act_key_row_set:
            self.key_row_changed.emit(orig_row)

    def _show_context_menu(self, pos):
        """본체 셀 우클릭 — pos 는 본체 위젯 좌표."""
        self._popup_cell_menu(self.viewport().mapToGlobal(pos))

    def _popup_cell_menu(self, global_pos):
        """선택 셀에 대한 병합 준비/취소 메뉴를 global_pos 에 띄운다.

        본체와 **틀 고정 오버레이(키 열/행 밴드)** 가 공유하는 단일 진입점이다. 고정 밴드의
        셀(키 열 및 그 좌측 열, 키 행)은 본체에서 숨겨져 오버레이가 그리므로, 오버레이에
        메뉴를 배선하지 않으면 그 위 우클릭은 아무 메뉴도 못 띄운다(예: 키 열이 D면
        #Description2(C) 셀에서 병합 준비 불가).

        ★ 반드시 popup()(비모달). 오버레이 이벤트 처리 중 모달 메뉴(exec_ 중첩 이벤트루프)를
        쓰면 Qt 상태가 깨져 access violation 이 난다(corner 헤더 메뉴와 동일한 제약).
        """
        # has_changed: any()가 첫 changed에서 조기 종료. has_staged: staged 집합 기반(소수).
        # 과거엔 전 선택 셀을 순회하며 둘 다 찾을 때까지 멈추지 않아, 전체 선택+staged 없음이면
        # 수백만 셀을 훑어 우클릭이 지연됐다.
        has_staged = self._has_staged_selection()
        has_changed = self._has_changed_selection()
        if not has_changed and not has_staged:
            return

        menu = QMenu(self)
        menu.setStyleSheet(MENU_QSS)
        menu.setAttribute(Qt.WA_DeleteOnClose)

        if has_changed:
            menu.addAction("A → B  병합 준비").triggered.connect(
                lambda _=False: self.stage_requested.emit(DIR_A2B))
            menu.addAction("B → A  병합 준비").triggered.connect(
                lambda _=False: self.stage_requested.emit(DIR_B2A))
        if has_staged:
            if has_changed:
                menu.addSeparator()
            menu.addAction("병합 준비 취소").triggered.connect(
                lambda _=False: self.unstage_requested.emit())

        menu.popup(global_pos)

    # ── 엑셀식 키보드 네비/선택/병합 단축키 ──────────────────────────────────
    def _is_empty_cell(self, r: int, c: int) -> bool:
        if r < 0 or r >= self.rowCount() or c < 0 or c >= self.columnCount():
            return True
        return self._model.display_text(r, c) == ""

    def _jump_target(self, r: int, c: int, dr: int, dc: int) -> tuple:
        """엑셀의 Ctrl+방향키 시맨틱으로 점프 대상 (row, col) 반환.
        세로 이동은 보이는 행만 밟는다 — 숨겨진 행에는 착지하지 않는다."""
        max_r = self.rowCount() - 1
        max_c = self.columnCount() - 1
        if max_r < 0 or max_c < 0:
            return (max(0, r), max(0, c))

        def step(rr, cc):
            if dr:
                # _next_visible_row 는 경계(더 갈 보이는 행 없음)에서 start(=rr)를 그대로
                # 돌려준다. 그 값을 그대로 쓰면 in_range 가 계속 참이라 아래 while 루프가
                # 무한 반복된다(격자 맨끝/맨앞에서 Ctrl+↕ → UI 행). 더 나아갈 수 없으면
                # off-grid 센티넬로 바꿔 in_range 가 루프를 종료시키게 한다.
                nn = self._next_visible_row(rr, dr)
                if nn == rr:
                    nn = max_r + 1 if dr > 0 else -1
                return nn, cc
            return rr, cc + dc

        def in_range(rr, cc):
            return 0 <= rr <= max_r and 0 <= cc <= max_c

        nr, nc = step(r, c)
        if not in_range(nr, nc):
            return (max(0, min(r, max_r)), max(0, min(c, max_c)))
        if self._is_empty_cell(r, c) or self._is_empty_cell(nr, nc):
            # 빈 구간 건너 다음 비어있지 않은 셀까지 — 못 찾으면 마지막 보이는 셀
            prev = (r, c)
            while in_range(nr, nc) and self._is_empty_cell(nr, nc):
                prev = (nr, nc)
                nr, nc = step(nr, nc)
            if not in_range(nr, nc):
                return prev
            return (nr, nc)
        # 연속 데이터의 마지막 비어있지 않은 셀까지
        while True:
            r2, c2 = step(nr, nc)
            if not in_range(r2, c2) or self._is_empty_cell(r2, c2):
                return (nr, nc)
            nr, nc = r2, c2

    @staticmethod
    def _header_jump_target(cur: int, d: int, last: int, is_empty) -> int:
        """엑셀 Ctrl+Shift 시맨틱의 열/행 단위 점프 대상 — _jump_target의 1차원 버전.
        매 호출마다 현재 위치(cur) 기준으로 재판정하므로 경계에서 한 번 더 누르면
        다음 값 블록 또는 그리드 끝으로 계속 이동한다."""
        if last < 0:
            return max(0, cur)
        n = cur + d
        if n < 0 or n > last:
            return min(max(cur, 0), last)
        if is_empty(cur) or is_empty(n):
            # 빈 구간 건너 다음 값 블록의 첫 열/행까지 — 없으면 그리드 끝
            while 0 <= n <= last and is_empty(n):
                n += d
            if n < 0 or n > last:
                return 0 if d < 0 else last
            return n
        # 연속 값 블록의 마지막 열/행까지
        while 0 <= n + d <= last and not is_empty(n + d):
            n += d
        return n

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

    def _shift_anchor(self, cur_r: int, cur_c: int) -> tuple:
        """Shift+방향키 확장의 앵커 — 현재 선택 사각형에서 현재 셀의 **반대 모서리**.

        Qt 내부 앵커는 노출되지 않으므로 선택 범위로 추정한다. 그래야 Shift+← 를 연속으로
        눌렀을 때 확장이 누적되고, 한 번 누를 때마다 두 칸짜리 선택으로 붕괴하지 않는다.
        선택이 없으면 현재 셀이 곧 앵커다."""
        sm = self.selectionModel()
        sel = sm.selection() if sm is not None else None
        if not sel:
            return cur_r, cur_c
        top = min(r.top() for r in sel)
        bot = max(r.bottom() for r in sel)
        left = min(r.left() for r in sel)
        right = max(r.right() for r in sel)
        return (bot if cur_r == top else top,
                right if cur_c == left else left)

    def _has_changed_selection(self) -> bool:
        # any()가 첫 changed 셀에서 조기 종료 — 변경이 있는 선택(일반)에선 빠르다.
        m = self._model
        return any(m.cell_kind(r, c) == "changed"
                   for (r, c) in self._iter_selected_cells())

    def _has_staged_selection(self) -> bool:
        """선택 안에 staged 셀이 있는가 — staged 집합(소수)을 선택 range에 대해 조회.
        전 선택 셀을 순회하던 과거 O(선택셀수) 대신 O(#staged × #range)."""
        sm = self.selectionModel()
        if sm is None:
            return False
        staged = self._model.staged_coords()
        if not staged:
            return False
        sel = sm.selection()
        model = self._model
        return any(sel.contains(model.index(r, c)) for (r, c) in staged)

    # ── 헤더 다중 선택 지원 ──────────────────────────────────────────────────
    def _full_columns_selected(self) -> list[int]:
        """모든 행에 걸쳐 선택된 열 목록 = '열 헤더 선택' 상태.
        selectedColumns()는 대량 선택에서 셀을 개별 열거해 O(선택셀수)로 느리므로,
        selection() range를 직접 본다: 행 전체(0..rowCount-1)를 덮는 range의 열들 합집합.
        앱은 전체 열 선택을 그 경계 range로 만들므로 selectedColumns()와 동일 결과 — O(#range).

        ★ 고정 키 밴드(0..key_col)도 **선택돼 있으면 그대로 보고한다.** 예전엔 기본으로
        걸러냈는데, 그건 선택에 키 열을 자동 보충하던 시절의 잔재다(지금은 _normalize_selection
        이 명시하듯 선택에 셀을 더하지 않는다 — 터치한 셀만 선택). 걸러내면 A~G 헤더를 잡아
        병합 준비해도 키 밴드인 A~D 가 조용히 빠진다."""
        sm = self.selectionModel()
        if sm is None:
            return []
        row_max = self.rowCount() - 1
        if row_max < 0:
            return []
        cols: set[int] = set()
        for rng in sm.selection():
            if rng.top() == 0 and rng.bottom() == row_max:
                cols.update(range(rng.left(), rng.right() + 1))
        return sorted(cols)

    def _full_rows_selected(self) -> list[int]:
        """모든 열에 걸쳐 선택된 행 목록 (열 대칭). O(#range).
        키 밴드 처리도 _full_columns_selected 와 동일 — 선택돼 있으면 그대로 보고한다."""
        sm = self.selectionModel()
        if sm is None:
            return []
        col_max = self.columnCount() - 1
        if col_max < 0:
            return []
        rows: set[int] = set()
        for rng in sm.selection():
            if rng.left() == 0 and rng.right() == col_max:
                rows.update(range(rng.top(), rng.bottom() + 1))
        return sorted(rows)

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

    # ── 커서가 갈 수 있는 칸(navigable) 판정 ─────────────────────────────────
    # 틀 고정이 켜지면 본체는 키 밴드를 **실제로 숨긴다**(_apply_col_hidden / 필터의
    # 0..key_row 숨김). 오버레이가 그 자리를 그리므로 사용자에겐 보이지만, Qt 기본
    # moveCursor 와 '숨김 = 못 감' 규칙은 그 칸을 갈 수 없는 곳으로 취급한다.
    # → 밴드는 '보이는 것'으로 쳐서 키보드가 들어갈 수 있게 한다.
    #   (이미 _prune_filtered_rows 가 키 프레임 행을 선택에서 안 걷어내고, 드래그 선택도
    #    밴드를 따로 처리한다 — 키보드 이동에만 이 개념이 빠져 있었다.)
    # 반면 '변경점만 보기'로 숨은 행은 여전히 갈 수 없다(정말로 볼 수 없으므로).
    # 열/행 한쪽만 고쳐지는 일이 없도록 축(_Axis)을 받아 처리한다.

    def _frozen_count(self, ax: _Axis) -> int:
        fc = getattr(self, "_freeze", None)
        return ax.frozen_count(fc) if (fc is not None and fc.active) else 0

    def _navigable(self, ax: _Axis, i: int) -> bool:
        return (not ax.hidden(self, i)) or i < self._frozen_count(ax)

    def _host_visible_cell(self, r: int, c: int) -> bool:
        """본체(host)가 이 칸을 실제로 그리는가 — 행·열 **모두** 숨김이 아니어야 한다.

        Qt 는 출발/도착 어느 한쪽이라도 숨김이면 moveCursor 가 무효 인덱스를 돌려주고
        scrollTo 도 무시한다. 즉 이 판정이 False 인 칸이 끼면 Qt 에 맡길 수 없다.
        """
        return not (AX_ROW.hidden(self, r) or AX_COL.hidden(self, c))

    def _page_rows(self) -> int:
        """PageUp/PageDown 한 걸음 = 화면에 실제로 그려진 행 수 - 1(한 행 겹침).

        Qt 는 이 걸음을 뷰포트 픽셀로 잰다. 그런데 커서가 고정 밴드 열에 있으면 scrollTo
        가 숨긴 인덱스를 무시해 뷰포트가 커서 기준으로 정렬되지 않는다 — 그래서 Qt 는
        '커서에서 한 페이지'가 아니라 '지금 화면에서 한 페이지'를 재고, 같은 칸에서
        눌러도 스크롤 상태에 따라 다른 행에 착지한다.
        (실측, 커서 1087행 고정 + 스크롤값만 10/15/20 으로 통제:
         밴드 A열 → 1120 / 1317 / 1841, 일반 E열 → 1841 / 1841 / 1841.)
        화면 행을 직접 세면 열·스크롤과 무관하게 같은 걸음이 된다.
        '변경점만 보기'로 숨은 행은 볼 수 없으므로 걸음에서 빠진다.
        """
        top = self.rowAt(0)
        if top < 0:
            top = 0
        bot = self.rowAt(self.viewport().height() - 1)
        if bot < 0:
            bot = self.rowCount() - 1
        on = sum(1 for r in range(top, bot + 1) if not AX_ROW.hidden(self, r))
        return max(1, on - 1)

    def _page_target_row(self, cur_r: int, delta: int) -> int:
        """cur_r 에서 delta 방향으로 한 페이지 — 갈 수 있는 행만 세어 옮긴다."""
        r = cur_r
        for _ in range(self._page_rows()):
            nxt = self._next_navigable(AX_ROW, r, delta)
            if nxt == r:
                break      # 격자 끝
            r = nxt
        return r

    def _next_navigable(self, ax: _Axis, start: int, delta: int) -> int:
        """start 에서 delta(+1/-1) 방향의 첫 '갈 수 있는' 칸. 없으면 start(제자리)."""
        last = ax.count(self) - 1
        n = start + delta
        while 0 <= n <= last:
            if self._navigable(ax, n):
                return n
            n += delta
        return start

    def _next_visible_row(self, start: int, delta: int) -> int:
        """start 에서 delta(+1/-1) 방향의 첫 '갈 수 있는' 행. 없으면 start(제자리).

        '변경점만 보기'로 숨은 행은 건너뛰고, **틀 고정 키 행은 건너뛰지 않는다** —
        본체에선 숨겨져 있지만 상단 고정 밴드에 보이기 때문이다. (예전 주석은 키 행이
        본체에서 숨김이 아니라고 적혀 있었지만 사실이 아니었고, 그래서 Ctrl+↑ 와 헤더
        Shift+↑ 가 키 행에 도달하지 못했다.)
        """
        return self._next_navigable(AX_ROW, start, delta)

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
            # 헤더 확장 판정/anchor는 키 열·행을 포함해야 한다 — 키 열만 선택된 상태에서도
            # 확장 경로가 켜져 키 열+데이터 열 다중 선택이 되도록(키 열은 본체 숨김이라
            # Qt 기본 확장이 안 됨). 컨텍스트 메뉴 대상은 여전히 키 제외 세트를 쓴다.
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
                delta = -1 if key == Qt.Key_Left else 1
                if ctrl:
                    # 엑셀처럼 현재 위치 기준 재판정 — 값 블록 끝 → 다음 블록 → 그리드 끝
                    m = self._model
                    target = self._header_jump_target(
                        cur_end, delta, self.columnCount() - 1,
                        lambda c: not m.col_has_values(c))
                else:
                    target = max(0, min(self.columnCount() - 1, cur_end + delta))
                self._select_column_range(self._header_anchor_col, target)
                self._set_current_cell_no_update(max(0, cur_r if cur_r >= 0 else 0), target)
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
                delta = -1 if key == Qt.Key_Up else 1
                if ctrl:
                    # 엑셀처럼 현재 위치 기준 재판정 — 값 블록 끝 → 다음 블록 → 그리드 끝
                    m = self._model
                    target = self._header_jump_target(
                        cur_end, delta, self.rowCount() - 1,
                        lambda r: not m.row_has_values(r))
                    # '변경 행만 보기'로 숨은 행에 착지하면 같은 방향의 보이는 행으로 스냅
                    # (숨은 행은 _normalize_selection이 되잘라내 확장이 무효화되므로).
                    if 0 <= target < self.rowCount() and self.isRowHidden(target):
                        target = self._next_visible_row(target, delta)
                else:
                    # 인접 인덱스(cur_end±1)가 아니라 '다음 보이는 행'으로 확장 — 필터로 숨은
                    # 행을 건너뛴다. 그러지 않으면 숨은 행에 착지→prune→제자리라 확장이 안 됐다.
                    target = self._next_visible_row(cur_end, delta)
                self._select_row_range(self._header_anchor_row, target)
                self._set_current_cell_no_update(target, max(0, cur_c if cur_c >= 0 else 0))
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
                    self.stage_requested.emit(DIR_A2B)
                event.accept(); return
            if key == Qt.Key_Left:
                if self._has_changed_selection():
                    self.stage_requested.emit(DIR_B2A)
                event.accept(); return
            if key in (Qt.Key_Backspace, Qt.Key_Delete):
                if self._has_staged_selection():
                    self.unstage_requested.emit()
                event.accept(); return

        # ── Enter/Return: 엑셀처럼 아래 칸으로 이동 ──
        # '한 행 아래'가 아니라 **갈 수 있는** 다음 행으로 옮긴다 — '변경점만 보기'로
        # 숨은 행에 커서를 놓으면 커서가 화면에서 사라지고, 보이지도 않는 행에 Alt+→ 로
        # 병합 준비가 걸린다. (실측: 필터 ON, 80행에서 Enter 6번 → 81~86 전부 숨은 행.
        # 같은 자리의 ↓ 는 102·123·180… 으로 제대로 건너뛴다.)
        if key in (Qt.Key_Return, Qt.Key_Enter) and not ctrl and not alt:
            if cur_r >= 0 and cur_c >= 0:
                tr = self._next_navigable(AX_ROW, cur_r, 1)
                if tr != cur_r:
                    self._set_current_cell(tr, cur_c)
            event.accept(); return

        # ── Shift+Space: 행 전체, Ctrl+Space: 열 전체 ──
        if key == Qt.Key_Space and shift and not ctrl and not alt and cur_r >= 0:
            self._select_row(cur_r)
            event.accept(); return
        if key == Qt.Key_Space and ctrl and not shift and not alt and cur_c >= 0:
            self._select_col(cur_c)
            event.accept(); return

        # ── 아직 셀을 고르지 않았을 때의 PageUp/PageDown — 화면만 한 페이지 굴린다 ──
        # '고르지 않았다' = 선택이 비어 있다. 현재 셀(커서)만 있는 경우도 여기 든다:
        # 표가 키보드 포커스를 받으면 Qt 가 선택과 무관하게 커서를 첫 칸에 꽂아 두기
        # 때문이다. 그 커서를 기준으로 페이지를 옮기면 **사용자가 고른 적 없는 셀이
        # 선택돼 버린다**(보고받은 증상). 고른 게 없으면 선택·커서를 건드리지 않고
        # 스크롤만 옮긴다 — B 패널은 스크롤 동기화로 따라온다.
        # (Qt 기본에 넘기면 '커서 없음'을 격자 맨 앞으로 쳐서 PageDown 인데도 화면이
        #  위로 튄다: 실측 스크롤 1 → 0.)
        sm_pick = self.selectionModel()
        if (key in (Qt.Key_PageUp, Qt.Key_PageDown) and not ctrl and not alt
                and (cur_r < 0 or cur_c < 0
                     or sm_pick is None or not sm_pick.hasSelection())):
            bar = self.verticalScrollBar()
            step = bar.pageStep() or 1
            bar.setValue(bar.value() + (-step if key == Qt.Key_PageUp else step))
            event.accept(); return

        # ── 방향키/Home/End/PageUp/PageDown: 밴드가 얽히거나 Shift 확장일 때 ──
        # (1) 커서 이동 — 틀 고정은 본체에서 키 밴드를 실제로 숨기고 오버레이가 그 자리를
        # 그린다. 그런데 QTableView.moveCursor 는 맨 끝에서 결과 칸이 숨김이면 **무효
        # 인덱스**를 돌려준다 (Qt 소스의 `if (!isRowHidden(..) && !isColumnHidden(..) &&
        # isIndexEnabled(..)) return result;  return QModelIndex();` 가드). 세로 이동은
        # 열이 그대로라 출발이 밴드면 결과도 밴드다 — 그래서 밴드가 얽힌 이동은 Qt 에
        # 맡길 수 없다. 밴드로 **들어가는** 것도(E 에서 ←), 밴드 **안에서** 위아래로
        # 움직이는 것도(A~D 에서 ↑/↓) 모두 제자리가 된다.
        # PageUp/PageDown 은 이 가드 앞에서 일찍 return 하므로 제자리가 되진 않지만,
        # 걸음을 뷰포트 픽셀로 재는 탓에 밴드 열에서 다른 칸에 착지한다(→ _page_rows).
        # 그래서 페이지 이동은 Shift 유무와 무관하게 우리가 센다.
        #
        # (2) Shift 확장 — 이쪽은 밴드와 무관하게 **항상** 우리가 처리해야 한다.
        # Qt 의 Shift+방향키 앵커는 QAbstractItemViewPrivate::pressedPosition, 즉
        # **뷰포트 픽셀 좌표**다. 확장은 `setSelection(QRect(pressedPosition - offset,
        # visualRect(new).center()), ...)` 로 그 픽셀 사각형을 indexAt 으로 되짚어
        # 만들어진다. 이 좌표는 마우스 press 만 갱신하므로,
        #   · 오버레이 드래그·헤더 선택·찾기 이동·Ctrl+점프처럼 우리가 _select_range 로
        #     만든 선택 뒤에는 값이 낡아 있고, 무효로 판정되면 Qt 는 조용히 **직전 커서
        #     칸으로 재앵커**해 선택을 무너뜨린다(A1:G1 에서 Shift+↓ 두 번 → 14~31행 ×
        #     G열만 남았다).
        #   · 픽셀→인덱스라서 폭 0 으로 숨긴 밴드 열(A~D)은 애초에 표현할 수 없다.
        #     앵커가 맞아도 밴드는 선택에서 빠진다.
        # 픽셀 앵커는 private 이라 손댈 수 없으니, Shift 확장은 선택 사각형에서 앵커를
        # 되찾는 _shift_anchor(이 앱의 헤더/밴드 확장이 이미 쓰는 모델)로 직접 만든다.
        #
        # Shift 가 없고 밴드도 안 걸리는 평범한 이동만 super() 에 넘긴다 — Qt 기본
        # 처리(스크롤·내부 앵커 갱신)를 필요 이상으로 뺏지 않는다.
        if (not ctrl and not alt and cur_r >= 0 and cur_c >= 0
                and key in (Qt.Key_Left, Qt.Key_Right, Qt.Key_Up, Qt.Key_Down,
                            Qt.Key_Home, Qt.Key_End,
                            Qt.Key_PageUp, Qt.Key_PageDown)):
            if key == Qt.Key_Home:
                tr, tc = cur_r, 0
            elif key == Qt.Key_End:
                tc = self.columnCount() - 1
                if tc >= 0 and not self._navigable(AX_COL, tc):
                    tc = self._next_navigable(AX_COL, tc, -1)
                tr, tc = cur_r, max(0, tc)
            elif key in (Qt.Key_PageUp, Qt.Key_PageDown):
                tr, tc = self._page_target_row(
                    cur_r, -1 if key == Qt.Key_PageUp else 1), cur_c
            elif key in (Qt.Key_Left, Qt.Key_Right):
                tr, tc = cur_r, self._next_navigable(
                    AX_COL, cur_c, -1 if key == Qt.Key_Left else 1)
            else:
                tr, tc = self._next_navigable(
                    AX_ROW, cur_r, -1 if key == Qt.Key_Up else 1), cur_c
            # Qt 에 넘겨도 되는 이동인가 —
            #  · Shift 확장은 앵커가 픽셀이라 절대 못 넘긴다(위 (2)).
            #  · PageUp/PageDown 은 걸음을 픽셀로 재 밴드 열에서 다른 칸에 착지하므로
            #    (→ _page_rows) Shift 유무와 관계없이 우리가 세어 밴드 안팎을 같게 만든다.
            #  · 나머지는 출발·도착이 **모두** 본체에 그려질 때만 넘긴다.
            delegate = (not shift
                        and key not in (Qt.Key_PageUp, Qt.Key_PageDown)
                        and self._host_visible_cell(cur_r, cur_c)
                        and self._host_visible_cell(tr, tc))
            if not delegate:
                if (tr, tc) != (cur_r, cur_c):
                    if shift:
                        ar, ac = self._shift_anchor(cur_r, cur_c)
                        self._select_range(ar, ac, tr, tc)
                        self._set_current_cell_no_update(tr, tc)
                    else:
                        self._move_current_cell(tr, tc)
                # 갈 곳이 없어도(격자 끝) 이벤트는 삼킨다 — super() 로 넘기면 Qt 가
                # 숨은 열을 피해 엉뚱한 칸으로 튄다(A 에서 Home → 첫 데이터 열).
                event.accept(); return

        # ── Ctrl(+Shift)+방향키: 데이터 경계 점프 (Excel 시맨틱) ──
        if ctrl and not alt and key in (Qt.Key_Left, Qt.Key_Right, Qt.Key_Up, Qt.Key_Down):
            if cur_r < 0 or cur_c < 0:
                return super().keyPressEvent(event)
            dr = -1 if key == Qt.Key_Up else (1 if key == Qt.Key_Down else 0)
            dc = -1 if key == Qt.Key_Left else (1 if key == Qt.Key_Right else 0)
            tr, tc = self._jump_target(cur_r, cur_c, dr, dc)
            if shift:
                # 앵커는 currentIndex(= 지금 커서)가 아니라 선택 사각형의 반대 모서리다.
                # currentIndex 를 쓰면 Ctrl+Shift+↓ 가 기존 확장을 커서 칸으로 붕괴시킨다.
                ar, ac = self._shift_anchor(cur_r, cur_c)
                self._select_range(ar, ac, tr, tc)
                self._set_current_cell_no_update(tr, tc)
            else:
                self._move_current_cell(tr, tc)
            event.accept(); return

        # ── Ctrl+Home / Ctrl+End ──
        if ctrl and not alt and key == Qt.Key_Home:
            tr, tc = 0, 0
            if shift and cur_r >= 0 and cur_c >= 0:
                ar, ac = self._shift_anchor(cur_r, cur_c)
                self._select_range(ar, ac, tr, tc)
                self._set_current_cell_no_update(tr, tc)
            else:
                self._move_current_cell(tr, tc)
            event.accept(); return
        if ctrl and not alt and key == Qt.Key_End:
            tr, tc = max(0, self.rowCount() - 1), max(0, self.columnCount() - 1)
            if shift and cur_r >= 0 and cur_c >= 0:
                ar, ac = self._shift_anchor(cur_r, cur_c)
                self._select_range(ar, ac, tr, tc)
                self._set_current_cell_no_update(tr, tc)
            else:
                self._move_current_cell(tr, tc)
            event.accept(); return

        super().keyPressEvent(event)

    def populate(self, diff_matrix: list[list], which: str,
                 merged_set: set = None, staged: dict = None,
                 row_meta: list = None, excluded_cols: set = None,
                 col_meta: list = None):
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
            # ※ `staged or {}`처럼 falsy 체크를 쓰면 '빈 dict'일 때 새 객체가
            #   만들어져 모델이 MainWindow 상태와 분리된다(이후 스테이징 색 미반영).
            #   반드시 is None 체크로 원본 참조를 공유해야 한다.
            self._model.set_diff_data(
                diff_matrix, row_meta,
                staged if staged is not None else {},
                merged_set if merged_set is not None else set(),
                self._excluded_cols, col_meta)
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
        return set(self._iter_selected_cells())

    def _iter_selected_cells(self):
        """선택 셀 (r, c)를 range 단위로 순회.
        selectedIndexes()처럼 전체 QModelIndex 리스트를 먼저 만들지 않아
        대량 선택 + 조기 종료 조합에서 훨씬 가볍다."""
        sm = self.selectionModel()
        if sm is None:
            return
        for rng in sm.selection():
            for r in range(rng.top(), rng.bottom() + 1):
                for c in range(rng.left(), rng.right() + 1):
                    yield (r, c)

    def _single_selected_cell(self):
        """정확히 1개 셀이 선택돼 있으면 (r, c), 아니면 None — O(range 수)."""
        sm = self.selectionModel()
        if sm is None:
            return None
        total = 0
        first = None
        for rng in sm.selection():
            if first is None:
                first = (rng.top(), rng.left())
            total += rng.width() * rng.height()
            if total > 1:
                return None
        return first if total == 1 else None

    def mirror_selection_from(self, src: "ExcelTableView"):
        """반대 패널의 선택을 range 단위로 그대로 복제 — O(range 수).
        헤더 클릭(열/행 전체 선택)은 range 1개라 셀 수와 무관하게 즉시 끝난다.
        (셀 집합으로 풀었다 재조립하면 행마다 range가 생겨 이후 모든 선택
        질의/페인팅이 느려진다 — 헤더 클릭 딜레이의 주범이었다.)"""
        src_sm = src.selectionModel()
        sm = self.selectionModel()
        if src_sm is None or sm is None:
            return
        self._populating = True
        prev_updates = self.updatesEnabled()
        self.setUpdatesEnabled(False)
        try:
            row_max = self.rowCount() - 1
            col_max = self.columnCount() - 1
            sel = QItemSelection()
            model = self.model()
            for rng in src_sm.selection():
                if rng.top() > row_max or rng.left() > col_max:
                    continue
                sel.append(QItemSelectionRange(
                    model.index(rng.top(), rng.left()),
                    model.index(min(rng.bottom(), row_max),
                                min(rng.right(), col_max))))
            sm.select(sel, QItemSelectionModel.ClearAndSelect)
        finally:
            self.setUpdatesEnabled(prev_updates)
            self._populating = False

    def mirror_selection(self, cells: set):
        # 셀 집합을 row별 연속 column 구간(span)으로 묶고, 연속 행의 span이
        # 동일하면 직사각형으로 수직 병합해 range 수를 최소화한다.
        # (열 전체 선택 = range 1개. range 수천 개짜리 selection model은
        # 이후 모든 질의·페인팅·선택 병합을 느리게 만든다.)
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

            def _spans(cs: list) -> tuple:
                cs.sort()
                out = []
                start = prev = cs[0]
                for c in cs[1:]:
                    if c == prev + 1:
                        prev = c
                        continue
                    out.append((start, prev))
                    start = prev = c
                out.append((start, prev))
                return tuple(out)

            sel = QItemSelection()
            model = self.model()
            run_start = prev_r = None
            run_spans = None

            def _flush(end_r):
                for (c1, c2) in run_spans:
                    sel.append(QItemSelectionRange(
                        model.index(run_start, c1), model.index(end_r, c2)))

            for r in sorted(by_row):
                sp = _spans(by_row[r])
                if run_spans is not None and sp == run_spans and r == prev_r + 1:
                    prev_r = r
                    continue
                if run_spans is not None:
                    _flush(prev_r)
                run_start = prev_r = r
                run_spans = sp
            if run_spans is not None:
                _flush(prev_r)
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


def _extract_folder_path(mime_data) -> str:
    """드롭된 MIME에서 첫 번째 '폴더' 로컬 경로를 반환(없으면 "")."""
    if mime_data.hasUrls():
        for url in mime_data.urls():
            p = url.toLocalFile()
            if p and os.path.isdir(p):
                return p
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
    """선택 셀의 값 표시란 — 읽기전용 뷰어 (직접 수정 기능 제거됨).
    기본 4줄 높이지만, 상하 스플리터 핸들을 드래그해 높이를 조절할 수 있다.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        line_h = self.fontMetrics().lineSpacing()
        # 기본 4줄이 한눈에 보이도록 — 고정이 아니라 최소/기본 높이만 지정(스플리터로 조절 가능).
        base_h = line_h * 4 + 12
        self.setMinimumHeight(line_h + 10)
        self.resize(self.width(), base_h)
        self._base_height = base_h
        self.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setReadOnly(True)   # 값 확인·복사 전용

    def event(self, ev):
        """고른 글자가 있으면 Ctrl+C 를 이 위젯이 가져간다.

        Qt 는 텍스트 위젯이 **편집 가능할 때만** Ctrl+C 에 ShortcutOverride 를
        돌려준다(QWidgetTextControl 은 Qt::TextEditable 일 때만 받아들인다). 이 위젯은
        읽기전용이라 그 길이 막혔고, 패널에 걸린 Ctrl+C 단축키가 먼저 먹었다 — 값을
        골라 놓고 눌러도 **파일 경로**가 복사됐다. 키가 닿기만 하면 복사 자체는 멀쩡하다.

        고른 게 있을 때만 가로챈다 — 아무것도 안 골랐으면 예전처럼 패널이 경로를 복사한다.
        """
        if (ev.type() == QEvent.ShortcutOverride
                and ev.matches(QKeySequence.Copy)
                and self.textCursor().hasSelection()):
            ev.accept()
            return True
        return super().event(ev)

    def sizeHint(self):
        return QSize(super().sizeHint().width(), getattr(self, "_base_height", 48))

    def text(self):
        return self.toPlainText()

    def _set_plain_default(self, val: str):
        """텍스트를 기본 서식(검정·배경 없음)으로 설정.
        setPlainText는 직전 강조(빨강)의 char format을 새 텍스트에 물려받는 경우가 있어,
        전체를 빈 QTextCharFormat으로 다시 칠해 강조 잔상을 제거한다."""
        self.setPlainText(val if val is not None else "")
        cur = QTextCursor(self.document())
        cur.select(QTextCursor.Document)
        cur.setCharFormat(QTextCharFormat())
        self.setTextCursor(QTextCursor(self.document()))

    def setText(self, val: str):
        self._set_plain_default(val)
        # 텍스트 설정 후 항상 맨 위부터 표시
        self.verticalScrollBar().setValue(0)

    def set_highlighted(self, val: str, ranges, color):
        """val을 표시하되 ranges의 [start, end) 문자 구간 배경을 color로 강조.
        ranges가 비면 일반 텍스트와 동일."""
        self._set_plain_default(val)
        if val and ranges:
            fmt = QTextCharFormat()
            fmt.setBackground(color)
            fmt.setForeground(CELL_DIFF_FG)   # 강조 구간 폰트도 빨강 (배경 위 가독)
            n = len(val)
            for start, end in ranges:
                start = max(0, start)
                end = min(n, end)
                if start >= end:
                    continue
                cur = self.textCursor()
                cur.setPosition(start)
                cur.setPosition(end, QTextCursor.KeepAnchor)
                cur.mergeCharFormat(fmt)
            # 커서/선택을 맨 앞으로 되돌려 강조가 파란 선택색에 가려지지 않게
            self.setTextCursor(QTextCursor(self.document()))
        self.verticalScrollBar().setValue(0)


class SheetTabBar(QTabBar):
    """하단 시트 탭 바 — 변경점이 있는 시트 탭을 노란 배경으로 강조.
    기본 탭을 그린 뒤 변경 탭 위에 반투명 노랑을 덧칠해 텍스트 가독을 유지한다."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._changed_idx: set[int] = set()   # 변경점 있는 탭 인덱스

    def set_changed_indices(self, indices):
        new = set(indices)
        if new != self._changed_idx:
            self._changed_idx = new
            self.update()

    def clear_changed(self):
        self.set_changed_indices(set())

    def paintEvent(self, _event):
        p = QStylePainter(self)
        opt = QStyleOptionTab()
        # 선택 탭을 마지막에 그려 이웃 탭 경계에 가려지지 않게 한다(기본 QTabBar와 동일).
        sel = self.currentIndex()
        order = [i for i in range(self.count()) if i != sel]
        if 0 <= sel < self.count():
            order.append(sel)
        for i in order:
            self.initStyleOption(opt, i)
            p.drawControl(QStyle.CE_TabBarTab, opt)
        # 변경 탭: 노랑을 덧칠하면 기존 글자가 묻히므로, 그 위에 진한 볼드 텍스트를 다시 그린다.
        for i in self._changed_idx:
            if not (0 <= i < self.count()):
                continue
            r = self.tabRect(i)
            p.fillRect(r, SHEET_TAB_CHANGED_BG)
            p.save()
            f = self.font()
            f.setBold(True)
            p.setFont(f)
            p.setPen(QColor(0x20, 0x20, 0x20))   # 진한 먹색 — 노랑 위 가독
            p.drawText(r, Qt.AlignCenter, self.tabText(i))
            p.restore()

