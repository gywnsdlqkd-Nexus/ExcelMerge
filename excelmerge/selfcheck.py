"""자체 점검 — 빌드된 exe 를 **화면 없이** 몰아 본다(릴리스 게이트용).

    ExcelMerge.exe --selfcheck <A 파일> <B 파일> <결과.json>

실제 키보드·마우스를 쓰던 uicheck.py 는 화면 좌표와 색을 짐작해야 했고, 그 짐작이
일곱 번 틀렸다(DPI · 포그라운드 · 최대화 · 경로칸 · 메뉴 위치 · 첫 실행 지연 · 색 가정).
게다가 도는 동안 사용자의 마우스를 가져갔다. 이 모드는 그 둘을 다 없앤다.

  · Qt 를 offscreen 으로 띄운다 — 창도, 작업 표시줄도, 입력 장치도 쓰지 않는다.
  · 입력은 Qt 안에서 만든다(QTest). 단축키 우선순위(ShortcutOverride)까지 실제 키와
    같은 경로를 탄다 — v219(셀값란 Ctrl+C 를 표가 가로챔)가 바로 그 경로의 버그였다.
  · 판단은 화면이 아니라 **앱 상태**로 한다: 상태바 · 클립보드 · 병합 준비 · 파일 · 예외.
  · 클립보드도 offscreen 플랫폼의 프로세스 안 것이라 사용자 클립보드를 건드리지 않는다.
  · 설정·로그도 임시 폴더로 돌린다(isolate_environment) — 사용자 것을 더럽히지 않는다.

못 잡는 것: OS 입력 · 창 활성화 · DPI 처럼 Qt 바깥의 문제. 그건 손으로 돌리는
uicheck.py 의 몫으로 남긴다.

이 모듈은 **excelmerge 의 다른 모듈을 import 하지 않는다**(맨 위에서). 진단 로그는
import 순간 %APPDATA% 아래로 경로가 정해지므로, 격리가 그보다 먼저 일어나야 한다.
"""
import hashlib
import json
import os
import shutil
import stat
import sys
import tempfile
import time
import traceback
from contextlib import contextmanager

FLAG = "--selfcheck"


def requested(argv=None) -> bool:
    return FLAG in (sys.argv if argv is None else argv)[1:]


def isolate_environment():
    """excelmerge 를 import 하기 **전에** 부른다.

    APPDATA 를 임시 폴더로 — 설정(prefs.json)·진단 로그·크래시 로그가 사용자 것과 섞이지
    않는다. 임시 경로의 비교 기록이 사용자 설정에 쌓이던 것도 막는다(uicheck 는 쌓았다).
    """
    os.environ["APPDATA"] = tempfile.mkdtemp(prefix="excelmerge_selfcheck_")
    os.environ["QT_QPA_PLATFORM"] = "offscreen"


# ── 작은 도구 ───────────────────────────────────────────────────────────────

def _pump(seconds=0.05):
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance()
    end = time.time() + seconds
    while True:
        app.processEvents()
        if time.time() >= end:
            return
        time.sleep(0.01)


def _wait(pred, timeout):
    """pred() 가 참이 될 때까지 이벤트를 돌린다 → 참이 됐는지."""
    end = time.time() + timeout
    while time.time() < end:
        _pump(0.02)
        if pred():
            return True
    return bool(pred())


def _sha(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def _writable_copy(src, dst):
    shutil.copy2(src, dst)
    os.chmod(dst, stat.S_IWRITE | stat.S_IREAD)    # P4V 읽기 전용 속성이 따라온다
    return dst


def _press(widget, key, modifier=None):
    """키 한 번. 수정자 키는 반드시 뗀다 — 눌린 채 남으면 다음 입력이 Ctrl+… 가 된다."""
    from PyQt5.QtCore import Qt
    from PyQt5.QtTest import QTest
    mod = Qt.NoModifier if modifier is None else modifier
    QTest.keyClick(widget, key, mod)
    if modifier == Qt.ControlModifier:
        QTest.keyRelease(widget, Qt.Key_Control, Qt.NoModifier)
    elif modifier == Qt.AltModifier:
        QTest.keyRelease(widget, Qt.Key_Alt, Qt.NoModifier)
    _pump()


@contextmanager
def _recorded_dialogs(seen):
    """알림 창을 띄우는 대신 (종류, 제목, 본문)으로 받아 둔다.

    offscreen 에서도 모달 창은 exec 루프에서 멈춘다 — 누를 사람이 없다. 띄웠다는
    사실과 내용이 곧 검사 대상이다(v221 의 증상이 '오류 창이 뜬다' 였다).
    """
    from PyQt5.QtWidgets import QMessageBox
    saved = {k: QMessageBox.__dict__[k] for k in ("information", "warning", "critical", "question")}

    def rec(kind, answer):
        def f(parent, title, text, *a, **k):
            seen.append((kind, str(title), str(text)))
            return answer
        return staticmethod(f)

    QMessageBox.information = rec("information", QMessageBox.Ok)
    QMessageBox.warning = rec("warning", QMessageBox.Ok)
    QMessageBox.critical = rec("critical", QMessageBox.Ok)
    # 이 점검의 흐름에는 질문 창이 없다. 뜨면 '아니오' — 묻는 일을 대신 승낙하지 않는다.
    QMessageBox.question = rec("question", QMessageBox.No)
    try:
        yield
    finally:
        for k, v in saved.items():
            setattr(QMessageBox, k, v)


@contextmanager
def _recorded_exceptions(errors):
    """Qt 슬롯에서 새어 나온 예외를 받아 둔다. 앱은 죽지 않지만 사용자는 오류 창을 본다."""
    prev = sys.excepthook

    def hook(exc_type, exc, tb):
        errors.append("".join(traceback.format_exception(exc_type, exc, tb)))

    sys.excepthook = hook
    try:
        yield
    finally:
        sys.excepthook = prev


# ── 점검 ────────────────────────────────────────────────────────────────────

def run_checks(a_src, b_src, workdir, timeout=90.0):
    """A/B 를 workdir 에 복사해 열고 사용자 흐름을 몰아 본다 → [{name, ok, detail}].

    원본은 읽기만 한다. QApplication 은 호출하는 쪽이 만들어 둔다.
    """
    from PyQt5.QtWidgets import QApplication

    from excelmerge.main_window import MainWindow

    checks = []

    def ok(name, cond, detail=""):
        checks.append({"name": name, "ok": bool(cond), "detail": str(detail)})
        return bool(cond)

    ext_a, ext_b = os.path.splitext(a_src)[1], os.path.splitext(b_src)[1]
    a = _writable_copy(a_src, os.path.join(workdir, "A" + ext_a))
    b = _writable_copy(b_src, os.path.join(workdir, "B" + ext_b))
    sha_a0, sha_b0 = _sha(a), _sha(b)

    dialogs, errors, status = [], [], []
    with _recorded_dialogs(dialogs), _recorded_exceptions(errors):
        win = MainWindow()
        win.show()
        QApplication.setActiveWindow(win)
        try:
            _drive(win, a, b, sha_a0, sha_b0, dialogs, errors, status, ok, timeout)
        finally:
            win.close()
            win.deleteLater()
            _pump(0.3)
    ok("처리되지 않은 예외 없음", not errors,
       errors[0].strip().splitlines()[-1] if errors else "")
    return checks


def _drive(win, a, b, sha_a0, sha_b0, dialogs, errors, status, ok, timeout):
    from PyQt5.QtCore import QPoint, Qt
    from PyQt5.QtWidgets import QApplication, QMenu

    from excelmerge.constants import DIR_A2B

    view = win.open_file_compare(a, b)
    view.status.messageChanged.connect(status.append)

    def compared():
        return any(m.startswith("비교 완료") for m in status + [view.status.currentMessage()])

    if not ok("비교 완료", _wait(compared, timeout),
              status[-1] if status else "상태바에 아무것도 안 떴다"):
        return

    table = view.panel_a.table
    box = view.panel_a.cell_edit

    # 1) 셀값란 Ctrl+C — 셀값란의 글자가 복사돼야 한다(v219: 표가 가로채 경로·표가 복사됨)
    table.setFocus()
    _press(table, Qt.Key_Down, Qt.AltModifier)           # 첫 변경 셀로
    value = box.toPlainText()
    if ok("셀값란에 값이 뜸", value.strip(), value[:40]):
        cb = QApplication.clipboard()
        cb.setText("<SELFCHECK>")
        box.setFocus()
        box.selectAll()
        _pump()
        _press(box, Qt.Key_C, Qt.ControlModifier)
        got = cb.text()
        ok("셀값란 Ctrl+C 가 그 값을 복사", got == value, repr(got[:40]))

    # 2) 우클릭 메뉴 → A → B 병합 준비 (메뉴는 실제로 띄우고 그 항목을 고른다)
    table.setFocus()
    _press(table, Qt.Key_A, Qt.ControlModifier)          # 전체 선택
    table._popup_cell_menu(table.viewport().mapToGlobal(QPoint(5, 5)))
    _pump(0.2)
    menu = QApplication.activePopupWidget()
    if not ok("우클릭 메뉴가 뜸", isinstance(menu, QMenu)):
        return
    items = [act for act in menu.actions() if not act.isSeparator()]
    ok("첫 항목이 A → B 병합 준비", items and items[0].text().startswith("A → B"),
       [act.text() for act in items])
    pick = next((act for act in items if act.text().startswith("A → B")), None)
    menu.close()
    if pick is None:
        return
    pick.trigger()
    _pump(0.2)
    staged = dict(view._staged)
    if not ok("병합 준비됨 (A → B)", staged and all(v == DIR_A2B for v in staged.values()),
              f"{len(staged)}칸"):
        return

    # 3) Ctrl+S — B 만 바뀌어야 한다
    n0 = len(dialogs)
    _press(table, Qt.Key_S, Qt.ControlModifier)
    _wait(lambda: any(d[1] in ("저장 완료", "오류") for d in dialogs[n0:]), timeout)
    after = dialogs[n0:]
    ok("저장 완료 안내", any(d[1] == "저장 완료" for d in after),
       [f"{d[0]}:{d[1]}" for d in after])
    ok("B 가 저장됨", _sha(b) != sha_b0)
    ok("A 는 그대로", _sha(a) == sha_a0)
    _pump(0.5)                                           # 끝난 저장 워커가 정리될 시간

    # 4) 저장 뒤 Ctrl+S 재입력 — 오류 창·예외가 없어야 한다(v221)
    n1, e1 = len(dialogs), len(errors)
    for _ in range(2):
        _press(table, Qt.Key_S, Qt.ControlModifier)
        _pump(0.3)
    bad = [f"{d[0]}:{d[1]}" for d in dialogs[n1:] if d[0] in ("critical", "warning")]
    ok("저장 뒤 Ctrl+S 에 오류 없음", not bad and len(errors) == e1,
       bad or (errors[-1].strip().splitlines()[-1] if len(errors) > e1 else ""))


# ── 진입점 ──────────────────────────────────────────────────────────────────

def main(argv=None) -> int:
    """exe 진입점에서 부른다. 결과는 JSON 으로 — 창 없는(windowed) exe 라 표준 출력이 없다."""
    argv = sys.argv if argv is None else argv
    i = argv.index(FLAG)
    try:
        a_src, b_src, out = argv[i + 1], argv[i + 2], argv[i + 3]
    except IndexError:
        return 2
    t0 = time.time()
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication(argv[:1])
    app.setStyle("Fusion")
    from excelmerge import __version__
    work = tempfile.mkdtemp(prefix="excelmerge_selfcheck_work_")
    try:
        try:
            checks = run_checks(os.path.abspath(a_src), os.path.abspath(b_src), work)
        except Exception:
            checks = [{"name": "점검 자체가 실패", "ok": False,
                       "detail": traceback.format_exc()}]
        result = {"version": __version__, "seconds": round(time.time() - t0, 1),
                  "ok": all(c["ok"] for c in checks), "checks": checks}
        with open(out, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        return 0 if result["ok"] else 1
    finally:
        shutil.rmtree(work, ignore_errors=True)
        _remove_isolated_appdata()


def _remove_isolated_appdata():
    """isolate_environment 가 만든 임시 APPDATA 를 지운다 — 그 폴더일 때만.

    진단 로그 파일이 열려 있으면 윈도우가 지우기를 거부하므로 핸들러부터 닫는다.
    """
    d = os.environ.get("APPDATA", "")
    if not os.path.basename(d).startswith("excelmerge_selfcheck_"):
        return
    import logging
    lg = logging.getLogger("excelmerge")
    for h in list(lg.handlers):
        try:
            h.close()
        except Exception:
            pass
        lg.removeHandler(h)
    shutil.rmtree(d, ignore_errors=True)
