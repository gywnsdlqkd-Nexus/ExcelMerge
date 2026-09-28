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
                        lambda w, silent=True, on_state=None: calls.append((w, silent, on_state)))
    win.update_btn.click()
    qapp.processEvents()
    assert len(calls) == 1, "버튼을 눌렀는데 조회하지 않았다"
    assert calls[0][0] is win
    assert calls[0][1] is False, "수동 확인이므로 결과를 항상 알려야 한다(silent=False)"
    assert callable(calls[0][2]), "조회 결과로 버튼 표시도 갱신해야 한다"


def test_check_failure_does_not_crash(win, monkeypatch, qapp):
    def boom(*_a, **_k):
        raise RuntimeError("네트워크 없음")
    monkeypatch.setattr("excelmerge.updater.check_for_updates", boom)
    monkeypatch.setattr(QMessageBox, "warning", staticmethod(lambda *a, **k: None))
    win.update_btn.click()
    qapp.processEvents()          # 예외가 새어 나오면 여기서 터진다


# ── 시작할 때는 조회하지 않는다 ───────────────────────────────────────────────

def test_startup_does_not_install_anything(win, monkeypatch, qapp):
    """창을 띄우는 것만으로는 설치 경로(check_for_updates)를 타지 않는다.

    버튼 강조용 '조용한 조회'(probe_latest)는 진입점에서 따로 부른다 — 그쪽은 창도
    띄우지 않고 설치도 하지 않는다. 여기서 막는 것은 **설치까지 가는 경로**다.
    """
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


def test_entry_point_never_installs_on_launch():
    """진입점 main() 은 설치 경로를 부르지 않는다 — 표시용 조회만 허용한다."""
    import excel_diff_merge
    src = inspect.getsource(excel_diff_merge.main)
    assert "check_for_updates" not in src, "시작 시 설치까지 가는 조회가 남아 있다"
    assert "_probe_updates" in src, "버튼 표시용 조회를 시작 시 걸어야 한다"


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


# ── 버튼 강조/흐림 ───────────────────────────────────────────────────────────
# 받을 게 있으면 눈에 띄게, 없으면 얌전하게. 단 **글자는 늘 '업데이트 확인'** 이다 —
# 라벨이 바뀌면 버튼 폭이 들썩여 옆 버튼까지 밀린다. 색으로 구분하되, 색만으로는
# 구분 못 하는 경우를 위해 툴팁에 상태를 적는다.

def _bg(btn):
    """QSS 가 실제로 칠한 배경색 — 위젯을 그려서 가운데 픽셀을 읽는다."""
    img = btn.grab().toImage()
    return img.pixelColor(btn.width() // 2, btn.height() // 2)


def test_button_starts_muted_before_anything_is_known(win):
    assert win.update_btn.property("hasUpdate") == "false"
    assert win.update_btn.text() == "업데이트 확인"


def test_available_update_highlights_the_button(win, qapp):
    muted = _bg(win.update_btn)
    win._on_update_state(True, {"version": "999"})
    qapp.processEvents()
    btn = win.update_btn
    assert btn.property("hasUpdate") == "true"
    assert btn.text() == "업데이트 확인", f"글자는 늘 고정이어야 한다: {btn.text()}"
    assert "999" in btn.toolTip(), "새 버전 번호는 툴팁에 있어야 한다"
    strong = _bg(btn)
    assert strong != muted, "강조 상태인데 색이 그대로다"
    assert strong.green() > strong.red() and strong.green() > strong.blue(), \
        f"강조색이 초록 계열이 아니다: {strong.name()}"


def test_no_update_mutes_the_button_again(win, qapp):
    win._on_update_state(True, {"version": "999"})
    qapp.processEvents()
    strong = _bg(win.update_btn)
    win._on_update_state(False)
    qapp.processEvents()
    btn = win.update_btn
    assert btn.property("hasUpdate") == "false"
    assert btn.text() == "업데이트 확인"
    assert "최신" in btn.toolTip()
    assert _bg(btn) != strong, "최신인데 강조가 남아 있다"


def test_failed_check_is_muted_not_highlighted(win, qapp):
    """조회 실패(모름)는 '받을 게 있음'이 아니다 — 흐리게 둔다."""
    win._on_update_state(None)
    qapp.processEvents()
    assert win.update_btn.property("hasUpdate") == "false"
    assert "확인합니다" in win.update_btn.toolTip()


def test_manual_check_updates_the_highlight(win, monkeypatch, qapp):
    """버튼을 눌러 '최신'을 확인하면 흐려지고, 새 버전이 나오면 강조된다."""
    monkeypatch.setattr(upd, "_download_and_apply", lambda *a: None)
    _frozen(monkeypatch, yes=False)
    upd._on_manifest(win, dict(MANIFEST), silent=True,
                     on_state=lambda av: win._on_update_state(av, dict(MANIFEST)))
    qapp.processEvents()
    assert win.update_btn.property("hasUpdate") == "true"

    upd._on_manifest(win, {"version": upd.__version__, "url": "u", "sha256": ""},
                     silent=True, on_state=lambda av: win._on_update_state(av, None))
    qapp.processEvents()
    assert win.update_btn.property("hasUpdate") == "false"


def test_probe_reports_only_newer_versions(win, qapp, monkeypatch):
    """probe_latest 는 새 버전일 때만 manifest 를 준다 — 같거나 낮으면 None."""
    seen = []

    class _W(QObject):
        done = pyqtSignal(object)
        finished = pyqtSignal()

        def __init__(self, *_a, **_k):
            super().__init__()

        def start(self):
            pass

    monkeypatch.setattr(upd, "UpdateCheckWorker", _W)
    monkeypatch.setattr(upd, "_source", lambda: ("manifest", "file:///x"))
    upd.probe_latest(win, seen.append)
    w = win._update_probe_worker
    w.done.emit({"version": "999", "url": "u"})
    assert seen[-1] and seen[-1]["version"] == "999"
    w.done.emit({"version": upd.__version__, "url": "u"})
    assert seen[-1] is None, "현재 버전인데 '받을 게 있음'으로 보고했다"
    w.done.emit(None)
    assert seen[-1] is None, "조회 실패인데 '받을 게 있음'으로 보고했다"


def test_probe_never_installs(win, monkeypatch):
    """표시용 조회는 절대 설치로 이어지지 않는다."""
    monkeypatch.setattr(upd, "_download_and_apply",
                        lambda *a: pytest.fail("표시용 조회가 설치를 시작했다"))
    monkeypatch.setattr(upd, "_source", lambda: None)
    got = []
    upd.probe_latest(win, got.append)
    assert got == [None], "소스가 없으면 '모름'(None)으로 보고해야 한다"
    src = inspect.getsource(upd.probe_latest)
    assert "_download_and_apply" not in src and "QMessageBox" not in src


def test_label_never_changes_between_states(win, qapp):
    """어떤 상태에서도 글자는 '업데이트 확인' — 폭이 들썩이지 않아야 한다."""
    widths = []
    for state, man in ((None, None), (True, {"version": "999"}), (False, None)):
        win._on_update_state(state, man)
        qapp.processEvents()
        assert win.update_btn.text() == "업데이트 확인"
        widths.append(win.update_btn.sizeHint().width())
    assert len(set(widths)) == 1, f"상태에 따라 버튼 폭이 변한다: {widths}"
