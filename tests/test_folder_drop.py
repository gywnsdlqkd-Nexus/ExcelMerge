# -*- coding: utf-8 -*-
"""폴더 비교 패널 드롭 등록 — 패널 전체(경로칸 밖 트리 영역 포함)가 폴더 드롭을 받는다.

버그: 파일 비교 패널은 패널 전체가 드롭 대상인데, 폴더 비교 패널은 좁은 경로 입력칸만
드롭을 받아 넓은 트리 영역에 폴더를 떨구면 무반응 → '반대 패널 등록 불가'로 보였다.
수정: 각 side 패널을 _FolderDropPanel(드롭 수용)으로 만들고 folder_dropped→_on_folder_chosen 배선.
"""
import os

import pytest

from excelmerge.folder_view import FolderCompareView, _FolderDropPanel
from excelmerge.statusbar import StatusBar


@pytest.fixture
def dead_pool(qapp):
    """생성한 위젯을 테스트 종료 시 확실히 파괴 — 떠도는 top-level 위젯이 다른 테스트의
    창 포커스/활성에 간섭(오염)하지 않도록 한다."""
    widgets = []
    yield widgets
    for w in widgets:
        try:
            w.close()
            w.deleteLater()
        except RuntimeError:
            pass
    qapp.processEvents()


def test_folderdroppanel_accepts_drops(dead_pool):
    p = _FolderDropPanel()
    dead_pool.append(p)
    assert p.acceptDrops() is True


def test_side_panels_are_drop_targets(dead_pool):
    view = FolderCompareView(StatusBar())
    dead_pool.append(view)
    pa = view.path_a.parentWidget()
    pb = view.path_b.parentWidget()
    assert isinstance(pa, _FolderDropPanel) and pa.acceptDrops()
    assert isinstance(pb, _FolderDropPanel) and pb.acceptDrops()


def test_drop_on_opposite_panel_registers_folder(dead_pool, tmp_path, monkeypatch):
    """A만 등록된 상태에서 반대(B) 패널에 폴더를 떨구면 B가 등록된다(버그 회귀 방지).
    비동기 스캔은 test_async_scan_view가 커버하므로 여기선 _rescan을 무력화해 등록만 검증."""
    a = tmp_path / "A"; a.mkdir()
    b = tmp_path / "B"; b.mkdir()
    view = FolderCompareView(StatusBar())
    dead_pool.append(view)
    monkeypatch.setattr(view, "_rescan", lambda after=None: None)

    view._set_side("a", str(a))
    assert view.folder_a() == os.path.abspath(str(a))
    assert view.folder_b() == ""

    # 반대 패널(B)의 드롭 = folder_dropped 방출과 동치 — 등록되어야 한다.
    view.path_b.parentWidget().folder_dropped.emit(str(b))
    assert view.folder_b() == os.path.abspath(str(b))
