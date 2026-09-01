# -*- coding: utf-8 -*-
"""변경 문자 강조(핑크 배경 + 빨강 글자) 렌더 회귀 테스트.

스크롤 비용 절감을 위해 강조 렌더가 두 경로로 갈렸다:
  · 단일 라인 → QPainter 직접(_paint_inline)  ← 셀당 QTextDocument 제거
  · 멀티라인   → 기존 QTextDocument 경로

두 경로 모두에서 **강조가 실제로 그려지는지**, 그리고 **강조 구간이 올바른 문자 위에
찍히는지**(문자열 앞부분 폭 기준 x 범위 안)를 화면 렌더 픽셀로 검증한다.
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFontMetrics

from excelmerge.theme import CELL_DIFF_HL


def _render(qapp, a_val, b_val, width=420, row_h=None):
    """(status=modified, a_val, b_val) 셀 하나를 가진 표를 렌더해 (이미지, 셀 rect, 폰트) 반환."""
    from excelmerge.widgets import ExcelTableView
    t = ExcelTableView("a")
    dm = [[("modified", a_val, b_val)]]
    t.populate(dm, "a", set(), {}, [(0, 0)], set())
    t.resize(width, 160)
    t.setColumnWidth(0, width - 60)
    if row_h:
        t.setRowHeight(0, row_h)
    t.show()
    qapp.processEvents()
    img = t.viewport().grab().toImage()   # visualRect 와 같은 좌표계(헤더 제외)
    rect = t.visualRect(t.model().index(0, 0))
    return img, rect, t


def _hl_columns(img, rect):
    """셀 안에서 강조색(CELL_DIFF_HL)이 나타나는 x 좌표 집합."""
    target = CELL_DIFF_HL.rgb() & 0xFFFFFF
    xs = set()
    for y in range(max(0, rect.top()), min(img.height(), rect.bottom() + 1)):
        for x in range(max(0, rect.left()), min(img.width(), rect.right() + 1)):
            if (img.pixel(x, y) & 0xFFFFFF) == target:
                xs.add(x)
    return xs


def test_single_line_highlight_is_drawn(qapp):
    """단일 라인(빠른 경로) — 강조 배경이 그려져야 한다."""
    img, rect, t = _render(qapp, "abcdef", "abXYef")
    xs = _hl_columns(img, rect)
    assert xs, "단일 라인 셀에 변경 강조가 전혀 그려지지 않았다"
    t.deleteLater()


def test_single_line_highlight_lands_on_changed_chars(qapp):
    """강조 x 범위가 '바뀐 문자'의 위치와 맞아야 한다(앞 2글자는 동일 → 강조 밖)."""
    a, b = "abcdef", "abXYef"
    img, rect, t = _render(qapp, a, b)
    xs = _hl_columns(img, rect)
    assert xs, "강조 없음"
    ranges = t.model().diff_char_ranges(0, 0)
    assert ranges, "모델이 변경 구간을 보고하지 않음"

    fm = QFontMetrics(t.font())
    lo = min(s for s, _ in ranges)
    hi = max(e for _, e in ranges)
    # 셀 텍스트 시작 x 는 스타일 여백에 따라 다르므로, '동일한 앞부분(lo글자)의 폭'보다
    # 오른쪽에서 강조가 시작되는지 / 전체 폭 안에서 끝나는지로 검증한다.
    pre = fm.horizontalAdvance(a, lo)
    span = fm.horizontalAdvance(a, hi) - pre
    assert max(xs) - min(xs) <= span + 6, (
        f"강조 폭이 변경 구간보다 넓다: 실측 {max(xs) - min(xs)}px, 예상 ~{span}px")
    assert min(xs) >= rect.left() + pre - 2, (
        "강조가 변경되지 않은 앞부분까지 덮었다")
    t.deleteLater()


def _hl_rows(img, rect):
    """셀 안에서 강조색이 나타나는 y 좌표 집합."""
    target = CELL_DIFF_HL.rgb() & 0xFFFFFF
    ys = set()
    for y in range(max(0, rect.top()), min(img.height(), rect.bottom() + 1)):
        for x in range(max(0, rect.left()), min(img.width(), rect.right() + 1)):
            if (img.pixel(x, y) & 0xFFFFFF) == target:
                ys.add(y)
                break
    return ys


def test_multiline_highlight_lands_on_the_changed_line(qapp):
    """멀티라인 셀은 QTextDocument 경로로 가야 하고, 강조는 **바뀐 줄**에 찍혀야 한다.

    회귀: 분기 조건을 "\n 포함 여부"로 두면 멀티라인이 단일 라인 경로로 새어 들어간다 —
    QStyledItemDelegate.initStyleOption 이 DisplayRole 의 \n 을 U+2028 로 바꿔 넣어
    opt.text 에는 \n 이 아예 없기 때문이다. 그러면 두 줄이 한 줄로 뭉개지고
    강조도 엉뚱한 위치에 찍힌다.
    """
    img, rect, t = _render(qapp, "one\ntwo", "one\ntXo", row_h=44)
    ys = _hl_rows(img, rect)
    assert ys, "멀티라인 셀에 변경 강조가 사라졌다"
    mid = rect.top() + rect.height() // 2
    assert min(ys) >= mid - 4, (
        f"변경은 둘째 줄인데 강조가 위쪽에 찍혔다 (y={min(ys)}..{max(ys)}, 중앙={mid}) "
        "— 멀티라인이 단일 라인 경로로 렌더된 것으로 보인다")
    t.deleteLater()


def test_multiline_first_line_change_highlights_upper_half(qapp):
    """대칭 검증 — 첫 줄이 바뀌면 강조는 위쪽 절반에 있어야 한다."""
    img, rect, t = _render(qapp, "one\ntwo", "oXe\ntwo", row_h=44)
    ys = _hl_rows(img, rect)
    assert ys, "강조 없음"
    mid = rect.top() + rect.height() // 2
    assert max(ys) <= mid + 4, (
        f"변경은 첫 줄인데 강조가 아래쪽에 찍혔다 (y={min(ys)}..{max(ys)}, 중앙={mid})")
    t.deleteLater()

def test_no_highlight_when_values_equal(qapp):
    """값이 같으면 강조가 없어야 한다(빠른 경로가 오판하지 않는지)."""
    img, rect, t = _render(qapp, "same", "same")
    assert not _hl_columns(img, rect), "동일 값 셀에 강조가 찍혔다"
    t.deleteLater()


def test_selected_cell_keeps_highlight(qapp):
    """선택된 셀에서도 강조가 보여야 한다(선택 배경에 묻히면 변경을 못 본다)."""
    from excelmerge.widgets import ExcelTableView
    t = ExcelTableView("a")
    t.populate([[("modified", "abcdef", "abXYef")]], "a", set(), {}, [(0, 0)], set())
    t.resize(420, 120)
    t.setColumnWidth(0, 360)
    t.show()
    t._select_range(0, 0, 0, 0)
    qapp.processEvents()
    img = t.viewport().grab().toImage()
    rect = t.visualRect(t.model().index(0, 0))
    assert _hl_columns(img, rect), "선택 상태에서 변경 강조가 사라졌다"
    t.deleteLater()
