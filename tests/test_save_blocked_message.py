# -*- coding: utf-8 -*-
"""쓸 수 없을 때의 안내 문구 — 이유를 단정하지 않는다.

예전 문구는 "변경하고자 하는 파일이 **열려 있으므로** 저장할 수 없습니다 … 파일을 닫은
후 다시 시도하세요." 였다. 그런데 쓸 수 없는 이유는 여럿이고, 실사용에서 가장 흔한 건
**읽기 전용**이다(P4V 체크아웃 전). 사용자는 열려 있지도 않은 파일을 찾아 닫으려다
시간을 버렸다. 단정하는 안내는 틀렸을 때 사람을 엉뚱한 데로 보낸다.

그리고 **경로 전체**를 보여 준다 — 파일 이름만 보여 주면 같은 이름이 빌드마다 있는
이 데이터에서는 어느 쪽인지 알 수 없다.
"""
import os

import pytest
from PyQt5.QtWidgets import QMessageBox

A_DATA = [["TID", "NAME"], ["k1", "칼"], ["k2", "방패"]]
B_DATA = [["TID", "NAME"], ["k1", "검"], ["k2", "방패"]]


@pytest.fixture
def view(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    monkeypatch.setenv("APPDATA", str(tmp_path))
    win = MainWindow()
    win.show()
    try:
        v = win.tabs.currentWidget()
        v.panel_a.set_path(str(tmp_path / "a.xlsx"))
        v.panel_b.set_path(str(tmp_path / "b.xlsx"))
        v._on_loaded(A_DATA, B_DATA)
        w = getattr(v, "_diff_worker", None)
        if w is not None:
            w.wait(8000)
        for _ in range(20):
            qapp.processEvents()
        yield v
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def _try_save(view, monkeypatch, side):
    """그 쪽으로 머지를 준비해 두고 저장을 시도한다. 뜬 경고를 돌려준다."""
    from excelmerge import diff_view as dv_mod
    from excelmerge.constants import DIR_A2B, DIR_B2A
    seen = []
    monkeypatch.setattr(dv_mod.QMessageBox, "warning",
                        staticmethod(lambda parent, title, text, *a, **k:
                                     seen.append((title, text)) or QMessageBox.Ok))
    monkeypatch.setattr(dv_mod, "_is_file_locked", lambda p: True)  # 쓸 수 없는 상태
    view._staged = {(1, 1): DIR_B2A if side == "a" else DIR_A2B}
    view._save_staged(side)
    assert seen, "경고가 뜨지 않았다"
    return seen[-1]


def test_it_does_not_claim_the_file_is_open(view, monkeypatch):
    """쓸 수 없는 이유를 단정하지 않는다 — 읽기 전용일 때 엉뚱한 안내가 된다."""
    title, text = _try_save(view, monkeypatch, "b")
    assert "열려" not in text and "열려" not in title, (title, text)
    assert "닫은" not in text, text


def test_it_says_what_actually_happened(view, monkeypatch):
    _title, text = _try_save(view, monkeypatch, "b")
    assert text.startswith("파일을 저장할 수 없습니다."), text


@pytest.mark.parametrize("side, label", [("a", "A"), ("b", "B")])
def test_it_names_the_side_and_the_full_path(view, monkeypatch, side, label):
    """A/B 중 어느 쪽인지, 그리고 **경로 전체**를 보여 준다.

    파일 이름만으로는 어느 빌드의 파일인지 알 수 없다 — 같은 이름이 빌드마다 있다.
    """
    _title, text = _try_save(view, monkeypatch, side)
    path = view.panel_a.get_path() if side == "a" else view.panel_b.get_path()
    assert f"{label} 파일 : {path}" in text, text
    assert os.sep in text or "/" in text, f"경로가 아니라 이름만 있다: {text}"


def test_the_writer_error_does_not_claim_it_is_open(tmp_path, monkeypatch):
    """저장 중 실제로 막혔을 때의 예외 문구도 같은 규칙을 따른다."""
    import excelmerge.xlsx_writer as W
    monkeypatch.setattr(W, "_is_file_locked", lambda p: True)
    p = tmp_path / "x.xlsx"
    p.write_bytes(b"not a real xlsx")
    with pytest.raises(PermissionError) as e:
        W._write_patches_to_file(str(p), {"A1": "x"})
    msg = str(e.value)
    assert "열려" not in msg, msg
    assert msg.startswith("파일을 저장할 수 없습니다.") and str(p) in msg, msg
