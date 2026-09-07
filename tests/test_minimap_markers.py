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
    """접기 도입 전 paintEvent — 비율마다 fillRect(중복 포함). 렌더 오라클."""

    @staticmethod
    def make(cls):
        class Naive(cls):
            def paintEvent(self, e):
                QScrollBar.paintEvent(self, e)
                if not self._ratios:
                    return
                opt = QStyleOptionSlider()
                self.initStyleOption(opt)
                groove = self.style().subControlRect(
                    QStyle.CC_ScrollBar, opt, QStyle.SC_ScrollBarGroove, self)
                if groove.width() <= 0 or groove.height() <= 0:
                    return
                p = QPainter(self)
                p.setRenderHint(QPainter.Antialiasing, False)
                p.setPen(Qt.NoPen)
                p.setBrush(self._MARKER_COLOR)
                if self.orientation() == Qt.Vertical:
                    top, x = groove.top(), groove.left() + 2
                    w, denom = max(1, groove.width() - 4), max(0, groove.height() - 2)
                    for r in self._ratios:
                        p.fillRect(x, top + int(r * denom), w, 2, self._MARKER_COLOR)
                else:
                    left, y = groove.left(), groove.top() + 2
                    h, denom = max(1, groove.height() - 4), max(0, groove.width() - 2)
                    for r in self._ratios:
                        p.fillRect(left + int(r * denom), y, 2, h, self._MARKER_COLOR)
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
