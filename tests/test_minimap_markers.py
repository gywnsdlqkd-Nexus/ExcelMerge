# -*- coding: utf-8 -*-
"""미니맵 마커 렌더 회귀 — 겹쳐 그리기를 접으면서 외형은 지켜야 한다.

변경 행이 많으면 마커 수가 스크롤바 트랙 픽셀 수를 크게 넘는다(실측: 2,479개가
고유 615px 에 4배 중복) — 같은 자리에 같은 사각형을 몇 번씩 그렸고, 스크롤 한 스텝의
5ms 가 여기 있었다. 이제 자리마다 한 번만 그린다(paint 2.50ms -> 0.71ms).

지키기 까다로운 두 가지:

  1. **알파**. 마커 색은 alpha=220 이라 겹쳐 그리면 진해진다. 중복만 버리면 변경이
     밀집한 구간이 눈에 띄게 연해진다 → 겹친 개수로 alpha_eff = 1-(1-a)^k 를 계산해
     한 번에 합성한다(_stacked_color).
  2. **해상도**. 접기는 트랙 길이(denom) 기준이다. 미리 접어 두면 창을 키워 트랙이
     길어졌을 때 마커가 실제보다 듬성해진다 → paint 시점의 denom 으로 계산하고
     denom 이 바뀌면 다시 접는다.
"""
import random

import pytest
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPainter
from PyQt5.QtWidgets import QScrollBar, QStyle, QStyleOptionSlider

from excelmerge.theme import MINIMAP_MARKER_COLOR


def _bar(qapp, orientation, ratios, w, h, cls=None):
    from excelmerge.widgets import MinimapScrollBar
    b = (cls or MinimapScrollBar)(orientation)
    b.setRange(0, 100)
    b.resize(w, h)
    b.set_change_ratios(ratios)
    b.show()
    qapp.processEvents()
    return b


def test_offsets_fold_duplicates_and_keep_counts(qapp):
    """접은 결과는 오프셋이 유일하고, 겹친 개수의 합이 원래 마커 수와 같아야 한다."""
    rnd = random.Random(1)
    ratios = sorted(rnd.random() for _ in range(2000))
    b = _bar(qapp, Qt.Vertical, ratios, 18, 616)
    folded = b._pixel_offsets(614)
    offs = [o for o, _ in folded]
    assert len(offs) == len(set(offs)), "같은 픽셀 오프셋이 두 번 남았다"
    assert sum(k for _, k in folded) == len(ratios), "겹친 개수 합이 마커 수와 다르다"
    assert len(offs) < len(ratios), "접히지 않았다(테스트 전제)"
    b.deleteLater()


def test_offsets_recomputed_when_track_grows(qapp):
    """트랙이 길어지면 다시 접어야 한다 — 창 크기 변경 후 마커가 듬성해지면 안 된다."""
    rnd = random.Random(2)
    ratios = sorted(rnd.random() for _ in range(2000))
    b = _bar(qapp, Qt.Vertical, ratios, 18, 616)
    short = len(b._pixel_offsets(614))
    tall = len(b._pixel_offsets(1198))
    assert tall > short, f"트랙을 늘렸는데 마커가 늘지 않았다 ({short} -> {tall})"
    assert len(b._pixel_offsets(614)) == short, "좁혔을 때 원래 밀도로 안 돌아왔다"
    b.deleteLater()


def test_set_change_ratios_invalidates_cache(qapp):
    b = _bar(qapp, Qt.Vertical, [0.1, 0.9], 18, 616)
    assert len(b._pixel_offsets(614)) == 2
    b.set_change_ratios([0.1, 0.2, 0.3, 0.9])
    assert len(b._pixel_offsets(614)) == 4, "비율이 바뀌었는데 낡은 캐시를 돌려줬다"
    b.deleteLater()


def test_stacked_color_alpha_grows_with_overlap(qapp):
    """겹친 개수가 늘면 실효 알파가 커지고, 1회는 원래 색 그대로여야 한다."""
    b = _bar(qapp, Qt.Vertical, [0.5], 18, 616)
    assert b._stacked_color(1).alpha() == MINIMAP_MARKER_COLOR.alpha()
    alphas = [b._stacked_color(k).alpha() for k in (1, 2, 3, 8)]
    assert alphas == sorted(alphas) and len(set(alphas)) > 1, alphas
    assert alphas[-1] == 255, f"많이 겹치면 불투명해져야 한다: {alphas}"
    for k in (1, 2, 8):
        c = b._stacked_color(k)
        assert (c.red(), c.green(), c.blue()) == (
            MINIMAP_MARKER_COLOR.red(), MINIMAP_MARKER_COLOR.green(),
            MINIMAP_MARKER_COLOR.blue()), "색조가 바뀌었다"
    b.deleteLater()


class _NaiveBar:
    """접기 도입 전 paintEvent — 비율마다 fillRect(중복 포함). 렌더 오라클.

    자리(띠 위치·핸들 건너뛰기)는 본 구현의 헬퍼를 그대로 쓴다. 이 오라클이 검증하는
    건 **알파 누적을 한 번의 합성으로 재현했는가**이지 마커를 어디에 그리는가가 아니다.
    """

    @staticmethod
    def make(cls):
        class Naive(cls):
            def paintEvent(self, e):
                QScrollBar.paintEvent(self, e)
                if not self._ratios:
                    return
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
                p = QPainter(self)
                p.setRenderHint(QPainter.Antialiasing, False)
                p.setPen(Qt.NoPen)
                p.setBrush(self._MARKER_COLOR)
                if self.orientation() == Qt.Vertical:
                    top, denom = groove.top(), max(0, groove.height() - 2)
                    for r in self._ratios:
                        y = top + int(r * denom)
                        if self._hits_slider(y, slider, True):
                            continue
                        p.fillRect(band, y, thick, 2, self._MARKER_COLOR)
                else:
                    left, denom = groove.left(), max(0, groove.width() - 2)
                    for r in self._ratios:
                        x = left + int(r * denom)
                        if self._hits_slider(x, slider, False):
                            continue
                        p.fillRect(x, band, 2, thick, self._MARKER_COLOR)
                p.end()
        return Naive


def _worst_channel_diff(i1, i2):
    assert i1.size() == i2.size()
    worst = 0
    for y in range(i1.height()):
        for x in range(i1.width()):
            p1, p2 = i1.pixel(x, y), i2.pixel(x, y)
            if p1 != p2:
                for sh in (0, 8, 16, 24):
                    worst = max(worst, abs(((p1 >> sh) & 255) - ((p2 >> sh) & 255)))
    return worst


@pytest.mark.parametrize("n", [1, 20, 2479])
@pytest.mark.parametrize("orient,w,h", [
    (Qt.Vertical, 18, 616), (Qt.Vertical, 18, 1200), (Qt.Horizontal, 766, 18)])
def test_render_matches_naive_within_rounding(qapp, n, orient, w, h):
    """접기 전/후 렌더가 8비트 반올림(채널당 최대 1) 안에서 같아야 한다.

    중복만 버리면(알파 누적 재현 없이) 밀집 구간에서 채널 오차가 30 이상 벌어진다 —
    이 테스트가 그 회귀를 잡는다.
    """
    from excelmerge.widgets import MinimapScrollBar
    rnd = random.Random(n)
    ratios = sorted(rnd.random() for _ in range(n))
    new = _bar(qapp, orient, ratios, w, h)
    old = _bar(qapp, orient, ratios, w, h, cls=_NaiveBar.make(MinimapScrollBar))
    worst = _worst_channel_diff(old.grab().toImage(), new.grab().toImage())
    assert worst <= 1, f"마커 외형이 달라졌다 — 최대 채널 오차 {worst}"
    new.deleteLater()
    old.deleteLater()


# ── 마커 자리: 바깥쪽 4px 띠 + 핸들 비우기 ──────────────────────────────────
# 예전에는 트랙 폭 14px 중 10px 을 핸들 위에 덮어 그려, 변경이 많은 파일에서 스크롤바가
# 주황 벽이 되고 핸들이 어디 있는지 보이지 않았다.

def _marker_pixels(bar):
    """마커 색이 칠해진 픽셀 좌표 목록 — 주황 계열만 추린다."""
    img = bar.grab().toImage()
    mc = MINIMAP_MARKER_COLOR
    out = []
    for y in range(img.height()):
        for x in range(img.width()):
            c = img.pixelColor(x, y)
            # 알파 합성 후라 정확히 같지는 않다 — '빨강이 크고 파랑이 작은' 주황만 본다.
            if c.red() > 180 and c.green() > 80 and c.blue() < 90 and abs(c.red() - mc.red()) < 80:
                out.append((x, y))
    return out


def _sub_rects(bar):
    opt = QStyleOptionSlider()
    bar.initStyleOption(opt)
    st = bar.style()
    return (st.subControlRect(QStyle.CC_ScrollBar, opt, QStyle.SC_ScrollBarGroove, bar),
            st.subControlRect(QStyle.CC_ScrollBar, opt, QStyle.SC_ScrollBarSlider, bar))


def _scrollable_bar(qapp, orientation, ratios, w, h):
    """핸들이 트랙보다 확실히 짧도록 range/pageStep 을 준 스크롤바."""
    from excelmerge.widgets import MinimapScrollBar
    b = MinimapScrollBar(orientation)
    b.setRange(0, 1000)
    b.setPageStep(40)
    b.setValue(400)
    b.resize(w, h)
    b.set_change_ratios(ratios)
    b.show()
    qapp.processEvents()
    return b


def test_vertical_markers_live_in_the_right_gutter(qapp):
    ratios = [i / 200 for i in range(200)]
    b = _scrollable_bar(qapp, Qt.Vertical, ratios, 14, 400)
    groove, _slider = _sub_rects(b)
    px = _marker_pixels(b)
    assert px, "마커가 하나도 안 그려졌다"
    xs = {x for x, _y in px}
    band_lo = groove.right() - b._GUTTER_PX
    assert min(xs) >= band_lo, f"띠 왼쪽으로 새어 나갔다: x {min(xs)} < {band_lo}"
    assert max(xs) <= groove.right(), "트랙 밖으로 나갔다"
    assert len(xs) <= b._GUTTER_PX, f"띠보다 두껍다: {sorted(xs)}"
    b.deleteLater()


def test_vertical_markers_never_cover_the_handle(qapp):
    ratios = [i / 400 for i in range(400)]
    b = _scrollable_bar(qapp, Qt.Vertical, ratios, 14, 400)
    _groove, slider = _sub_rects(b)
    assert slider.height() < 300, "전제: 핸들이 트랙보다 짧아야 의미 있는 테스트"
    bad = [(x, y) for x, y in _marker_pixels(b) if slider.top() <= y <= slider.bottom()]
    assert not bad, f"핸들 위에 마커가 그려졌다: {bad[:5]}"
    b.deleteLater()


def test_horizontal_markers_live_in_the_bottom_gutter(qapp):
    ratios = [i / 200 for i in range(200)]
    b = _scrollable_bar(qapp, Qt.Horizontal, ratios, 400, 14)
    groove, slider = _sub_rects(b)
    px = _marker_pixels(b)
    assert px
    ys = {y for _x, y in px}
    assert min(ys) >= groove.bottom() - b._GUTTER_PX, f"띠 위로 새어 나갔다: {sorted(ys)}"
    assert not [1 for x, _y in px if slider.left() <= x <= slider.right()], \
        "핸들 위에 마커가 그려졌다"
    b.deleteLater()


def test_markers_still_mark_every_change_outside_the_handle(qapp):
    """자리만 옮겼을 뿐, 핸들 밖 변경 위치는 하나도 빠지지 않아야 한다."""
    from excelmerge.widgets import MinimapScrollBar
    ratios = [0.02, 0.10, 0.30, 0.55, 0.80, 0.95]
    b = _scrollable_bar(qapp, Qt.Vertical, ratios, 14, 400)
    groove, slider = _sub_rects(b)
    rows = {y for _x, y in _marker_pixels(b)}
    expect = 0
    for r in ratios:
        y = groove.top() + int(r * max(0, groove.height() - 2))
        if not MinimapScrollBar._hits_slider(y, slider, True):
            expect += 1
            assert y in rows or (y + 1) in rows, f"{r} 위치 마커가 없다 (y={y})"
    assert expect >= 4, "전제: 핸들 밖 마커가 여러 개 있어야 한다"
    b.deleteLater()
