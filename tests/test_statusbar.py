# -*- coding: utf-8 -*-
"""StatusBar 진행률 바 — 확정/불확정 전환과 숨김."""
from excelmerge.statusbar import StatusBar


def test_determinate_progress(qapp):
    sb = StatusBar()
    sb.set_progress(3, 10)
    assert sb._bar.maximum() == 10
    assert sb._bar.value() == 3
    assert not sb._bar.isHidden()


def test_value_clamped(qapp):
    sb = StatusBar()
    sb.set_progress(50, 10)          # done > total
    assert sb._bar.value() == 10


def test_busy_when_total_unknown(qapp):
    sb = StatusBar()
    sb.set_progress(5, 0)            # total<=0 → 불확정(maximum 0)
    assert sb._bar.maximum() == 0
    assert not sb._bar.isHidden()
    sb.begin_busy()
    assert sb._bar.maximum() == 0


def test_end_progress_hides(qapp):
    sb = StatusBar()
    sb.set_progress(3, 10)
    sb.end_progress()
    assert sb._bar.isHidden()
