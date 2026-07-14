"""ExcelMerge 진입점 — 실제 구현은 excelmerge 패키지에 있다."""
import os
import sys

from PyQt5.QtWidgets import QApplication

from excelmerge.main_window import MainWindow
from excelmerge.theme import load_app_icon


def _parse_args() -> tuple[str, str]:
    """
    P4V diff 호출 형식: -s <원본> -d <수정본>
    위치 인자 2개도 지원: <파일A> <파일B>
    반환: (path_a, path_b)  — 없으면 빈 문자열
    """
    args = sys.argv[1:]
    path_a = path_b = ""
    i = 0
    positional = []
    while i < len(args):
        if args[i] == "-s" and i + 1 < len(args):
            path_a = args[i + 1]; i += 2
        elif args[i] == "-d" and i + 1 < len(args):
            path_b = args[i + 1]; i += 2
        elif not args[i].startswith("-"):
            positional.append(args[i]); i += 1
        else:
            i += 1
    # -s/-d 없이 위치 인자로 넘어온 경우
    if not path_a and len(positional) >= 1:
        path_a = positional[0]
    if not path_b and len(positional) >= 2:
        path_b = positional[1]
    return path_a, path_b


def _set_windows_app_user_model_id():
    """Windows 작업 표시줄이 앱을 python.exe 그룹과 분리하고
    우리가 지정한 아이콘을 사용하도록 AppUserModelID를 등록한다."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "ExcelMerge.App")
    except Exception:
        pass


def main():
    # 크래시/미처리 예외를 %APPDATA%/ExcelMerge/crash.log 에 기록(재현 어려운 버그 진단용).
    try:
        from excelmerge.crashlog import install as _install_crashlog
        _install_crashlog()
    except Exception:
        pass
    _set_windows_app_user_model_id()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setWindowIcon(load_app_icon())
    win = MainWindow()

    # P4V 등 외부 도구가 커맨드라인으로 경로를 넘긴 경우 자동 로드.
    # 폴더가 하나라도 있으면 폴더 비교 탭, 아니면 파일 비교 탭으로 연다.
    path_a, path_b = _parse_args()
    a_is_dir = bool(path_a) and os.path.isdir(path_a)
    b_is_dir = bool(path_b) and os.path.isdir(path_b)
    a_is_file = bool(path_a) and os.path.isfile(path_a)
    b_is_file = bool(path_b) and os.path.isfile(path_b)
    if a_is_dir or b_is_dir:
        win.open_folder_compare(path_a if a_is_dir else "",
                                path_b if b_is_dir else "")
    elif a_is_file or b_is_file:
        win.open_file_compare(path_a if a_is_file else "",
                              path_b if b_is_file else "")

    win.show()

    # 시작 후 백그라운드로 새 버전 확인(비차단). 서버 미설정/네트워크 문제면 조용히 넘어감.
    try:
        from excelmerge.updater import check_for_updates
        check_for_updates(win)
    except Exception:
        pass

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
