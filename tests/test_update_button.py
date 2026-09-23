# -*- coding: utf-8 -*-
"""업데이트는 '업데이트 확인' 버튼을 눌렀을 때만 일어난다.

예전에는 앱을 켤 때마다 자동으로 조회하고, 새 버전이 있으면 **확인 창 없이** 받아
교체·재시작했다(다운로드를 취소하면 앱이 아예 종료됐다). 강제 업데이트가 불편하다는
피드백을 받아 시작 시 조회를 없애고, 버튼 → 확인 질문 → 설치 순서로 바꿨다.
"""
import inspect

import pytest
from PyQt5.QtCore import Qt, QObject, pyqtSignal
from PyQt5.QtWidgets import QMessageBox

import excelmerge.updater as upd


@pytest.fixture
def win(qapp, monkeypatch, tmp_path):
    from excelmerge.main_window import MainWindow
    monkeypatch.setenv("APPDATA", str(tmp_path))     # 전역 설정 격리
    w = MainWindow()
    w.show()
    qapp.processEvents()
    try:
        yield w
    finally:
        w.close()
        w.deleteLater()
        qapp.processEvents()


# ── 버튼 ─────────────────────────────────────────────────────────────────────

def test_button_sits_in_the_top_right_corner(win):
    btn = win.update_btn
    assert btn.text() == "업데이트 확인"
    assert btn.isVisible()
    corner = win.tabs.cornerWidget(Qt.TopRightCorner)
    assert btn.parent() is corner, "우상단 코너(＋ 새 비교 옆)에 있어야 한다"
    assert btn.focusPolicy() == Qt.NoFocus, "눌러도 키보드 포커스는 표에 남아야 한다"


def test_clicking_the_button_checks_for_updates(win, monkeypatch, qapp):
    calls = []
    monkeypatch.setattr("excelmerge.updater.check_for_updates",
                        lambda w, silent=True: calls.append((w, silent)))
    win.update_btn.click()
    qapp.processEvents()
    assert len(calls) == 1, "버튼을 눌렀는데 조회하지 않았다"
    assert calls[0][0] is win
    assert calls[0][1] is False, "수동 확인이므로 결과를 항상 알려야 한다(silent=False)"


def test_check_failure_does_not_crash(win, monkeypatch, qapp):
    def boom(*_a, **_k):
        raise RuntimeError("네트워크 없음")
    monkeypatch.setattr("excelmerge.updater.check_for_updates", boom)
    monkeypatch.setattr(QMessageBox, "warning", staticmethod(lambda *a, **k: None))
    win.update_btn.click()
    qapp.processEvents()          # 예외가 새어 나오면 여기서 터진다


# ── 시작할 때는 조회하지 않는다 ───────────────────────────────────────────────

def test_startup_does_not_check_for_updates(win, monkeypatch, qapp):
    """창을 띄우는 것만으로는 아무것도 조회하지 않는다."""
    calls = []
    monkeypatch.setattr("excelmerge.updater.check_for_updates",
                        lambda *a, **k: calls.append(a))
    from excelmerge.main_window import MainWindow
    w2 = MainWindow()
    w2.show()
    qapp.processEvents()
    try:
        assert not calls
    finally:
        w2.close()
        w2.deleteLater()
        qapp.processEvents()


def test_entry_point_no_longer_checks_on_launch():
    """진입점 main() 에 시작 시 조회가 남아 있으면 안 된다."""
    import excel_diff_merge
    src = inspect.getsource(excel_diff_merge.main)
    assert "check_for_updates" not in src, "시작 시 업데이트 조회가 남아 있다"


# ── 새 버전이 있어도 물어본 뒤에 받는다 ──────────────────────────────────────

MANIFEST = {"version": "999", "url": "https://x/ExcelMerge_v999.exe",
            "sha256": "", "notes": "새 기능 A\n새 기능 B"}


def _frozen(monkeypatch, yes: bool):
    """frozen(패키지) 실행처럼 보이게 하고, 확인 창 답을 고정한다."""
    monkeypatch.setattr(upd.sys, "frozen", True, raising=False)
    monkeypatch.setattr(
        QMessageBox, "question",
        staticmethod(lambda *a, **k: QMessageBox.Yes if yes else QMessageBox.No))


def test_new_version_asks_before_downloading(win, monkeypatch):
    started = []
    monkeypatch.setattr(upd, "_download_and_apply",
                        lambda w, m: started.append(m))
    _frozen(monkeypatch, yes=False)
    upd._on_manifest(win, dict(MANIFEST), silent=False)
    assert not started, "'아니오'인데 다운로드를 시작했다"

    _frozen(monkeypatch, yes=True)
    upd._on_manifest(win, dict(MANIFEST), silent=False)
    assert len(started) == 1, "'예'인데 다운로드를 시작하지 않았다"


def test_already_latest_says_so(win, monkeypatch):
    seen = []
    monkeypatch.setattr(QMessageBox, "information",
                        staticmethod(lambda _w, _t, msg, *a, **k: seen.append(msg)))
    monkeypatch.setattr(upd, "_download_and_apply",
                        lambda *a: pytest.fail("최신인데 다운로드를 시작했다"))
    upd._on_manifest(win, {"version": upd.__version__, "url": "u", "sha256": ""},
                     silent=False)
    assert seen and "최신" in seen[0]


# ── 다운로드를 취소해도 앱이 죽지 않는다 ─────────────────────────────────────

class _FakeWorker(QObject):
    progress = pyqtSignal(int, int)
    error = pyqtSignal(str)
    done = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, *_a, **_k):
        super().__init__()
        self.canceled = False
        self.started = False

    def start(self):
        self.started = True

    def cancel(self):
        self.canceled = True


def test_canceling_the_download_only_cancels_the_worker(win, monkeypatch, qapp):
    """취소는 받던 것만 버린다 — 앱은 그대로 살아 있어야 한다."""
    made = {}

    def _factory(*a, **k):
        w = _FakeWorker()
        made["w"] = w
        return w

    monkeypatch.setattr(upd, "UpdateDownloadWorker", _factory)
    upd._download_and_apply(win, dict(MANIFEST))
    qapp.processEvents()
    assert made["w"].started and win._update_dl_worker is made["w"]

    dlg = next(c for c in win.children() if c.__class__.__name__ == "QProgressDialog")
    # 사용자가 '취소'를 누른 것과 같다 — QProgressDialog.cancel() 은 신호를 쏘지 않는다.
    dlg.canceled.emit()
    qapp.processEvents()

    assert made["w"].canceled, "취소가 워커에 전달되지 않았다"
    assert win.isVisible(), "취소했는데 창이 닫혔다"

    # 정리 — 가짜 워커를 창에 매달아 둔 채 테스트를 끝내면 종료 시 Qt 가 죽는다.
    dlg.close()
    win._update_dl_worker = None
    qapp.processEvents()


def test_cancel_does_not_quit_the_app():
    """예전 정책(취소 = 앱 종료)이 되살아나지 않게 못 박는다.

    QApplication.quit 은 테스트에서 이벤트 루프가 없어 호출해도 아무 일이 없다 —
    행동으로는 구분되지 않으므로 취소 경로에 종료 호출이 없다는 것으로 고정한다.
    """
    src = inspect.getsource(upd._download_and_apply)
    cancel_part = src[src.index("def _cancel"):]
    assert "quit()" not in cancel_part, "취소 시 앱을 종료하는 코드가 남아 있다"
