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
            "Netmarble.ExcelMerge.App")
    except Exception:
        pass


def main():
    _set_windows_app_user_model_id()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setWindowIcon(load_app_icon())
    win = MainWindow()

    # P4V 등 외부 도구가 커맨드라인으로 파일을 넘긴 경우 자동 로드
    path_a, path_b = _parse_args()
    if path_a and os.path.isfile(path_a):
        win.panel_a.set_path(path_a)
    if path_b and os.path.isfile(path_b):
        win.panel_b.set_path(path_b)
    # 두 파일 모두 있으면 preview 없이 바로 비교 (load_sheet 2회로 단축)
    if path_a and path_b and os.path.isfile(path_a) and os.path.isfile(path_b):
        win._run_compare()
    elif path_a and os.path.isfile(path_a):
        win._run_preview("a", path_a)
    elif path_b and os.path.isfile(path_b):
        win._run_preview("b", path_b)

    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
