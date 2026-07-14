"""크래시/미처리 예외 로깅.

재현이 어려운 크래시 원인을 사용자 PC에서 포착하기 위해, 미처리 파이썬 예외와
치명적 오류(세그폴트 등)를 %APPDATA%/ExcelMerge/crash.log 에 기록한다.
콘솔이 있으면 표준 출력에도 찍히고(run.bat), 가능하면 대화상자로도 안내한다.
"""
import datetime
import os
import sys
import traceback
import faulthandler

_fault_file = None   # faulthandler가 세그폴트 덤프에 쓰는 파일 핸들(수명 유지용)

_MAX_BYTES = 1024 * 1024   # crash.log 크기 상한(1MB). 초과 시 .1 로 1세대 로테이트.


def log_path() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "ExcelMerge", "crash.log")


def _rotate_if_large():
    """crash.log 가 상한을 넘으면 crash.log.1 로 밀어내 무한 증가를 막는다(1세대 보관)."""
    try:
        p = log_path()
        if os.path.isfile(p) and os.path.getsize(p) > _MAX_BYTES:
            bak = p + ".1"
            if os.path.exists(bak):
                os.remove(bak)
            os.replace(p, bak)
    except Exception:
        pass


def _append(text: str):
    try:
        p = log_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "a", encoding="utf-8") as f:
            f.write(text)
    except Exception:
        pass


def install():
    """faulthandler + sys.excepthook 설치. main() 시작 시 1회 호출."""
    global _fault_file
    try:
        p = log_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        _rotate_if_large()   # append 핸들을 열기 전에 상한 초과분을 밀어낸다
        _fault_file = open(p, "a", encoding="utf-8")
        _fault_file.write("\n===== 세션 시작 %s =====\n" % datetime.datetime.now())
        _fault_file.flush()
        faulthandler.enable(_fault_file)   # 세그폴트/치명 오류 시 C 스택 덤프
    except Exception:
        try:
            faulthandler.enable()
        except Exception:
            pass

    def hook(exctype, value, tb):
        msg = "".join(traceback.format_exception(exctype, value, tb))
        _append("\n===== 미처리 예외 %s =====\n%s\n" % (datetime.datetime.now(), msg))
        try:
            sys.__excepthook__(exctype, value, tb)   # 콘솔에도 출력
        except Exception:
            pass
        try:
            from PyQt5.QtWidgets import QMessageBox, QApplication
            if QApplication.instance() is not None:
                QMessageBox.critical(
                    None, "오류 발생",
                    "오류가 발생했습니다. 자세한 내용이 아래 파일에 기록됐습니다:\n"
                    + log_path() + "\n\n" + f"{exctype.__name__}: {value}")
        except Exception:
            pass

    sys.excepthook = hook
