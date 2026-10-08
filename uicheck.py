# -*- coding: utf-8 -*-
"""설치본을 실제로 몰아 보고 사람이 쓰는 대로 동작하는지 확인한다 — '실기 확인'의 자동화.

**왜 있나.** pytest 는 오프스크린에서 위젯을 직접 만져 본다. 그걸로는 못 잡는 게 있다.

  · v221 — 저장이 끝난 워커 참조로 isRunning() 을 불러 터졌다. 진짜 저장을 한 번
    돌려야 생기는 상태라, 저장을 가로채던 유닛 테스트는 아무 표시 없이 통과했다.
  · v224 — 고정 열이 본체를 못 따라갔다. 스크롤바 값·숨김 집합·호출 횟수는 모두
    정상이었고 **그려지는 결과만** 틀렸다.

둘 다 설치본을 띄워 눈으로 보고서야 드러났다. 그 확인을 매번 손으로 하느라 하네스를
세 번 다시 만들었고, 그때마다 같은 함정을 다시 밟았다. 그 지식을 여기 둔다.

**함정들** (전부 실제로 밟았다)

  DPI        모니터마다 배율이 다르면 DPI 비인식 프로세스는 가상화된 좌표를 받는다.
             GetWindowRect / SetCursorPos / PrintWindow 가 서로 다른 공간을 가리켜
             클릭이 한 행씩 빗나간다. import 시점에 per-monitor 로 선언한다.
  포그라운드  입력을 보내기 전에 대상 창이 **실제로** 앞에 있는지 확인한다. 확인 없이
             보내면 사용자가 쓰던 다른 앱(언리얼 에디터)에 키가 들어간다. 실제로 두 번
             막혔다. AttachThreadInput 없이는 SetForegroundWindow 가 거절당한다.
  복원        focus() 에서 무조건 ShowWindow(SW_RESTORE) 를 부르면 **최대화가 풀린다.**
             최소화일 때만 되돌린다.
  팝업       메뉴·대화상자는 **별도 최상위 창**이다. 본창만 캡처하면 안 잡혀서
             '아무 반응 없음' 으로 오판한다. 창을 열거해 따로 찍는다.
  클립보드    GetClipboardData 는 restype 을 지정해야 한다. 기본값(c_int)은 64비트
             핸들을 잘라 GlobalLock 이 NULL 을 돌려주고 액세스 위반이 난다.
             powershell Set-Clipboard 는 그 프로세스가 포커스를 가져가니 쓰지 않는다.

**사용**

    python uicheck.py <A.xlsx> <B.xlsx> [--exe 경로] [--expect-version 224]

원본은 **읽기만** 한다 — 임시 폴더로 복사해 거기에만 쓴다(P4V 읽기 전용 속성도 푼다).
하나라도 어긋나면 종료 코드 1.
"""
import argparse
import atexit
import ctypes
import hashlib
import os
import shutil
import stat
import subprocess
import sys
import tempfile
import time
from ctypes import wintypes

for _s in (sys.stdout, sys.stderr):          # cp949 콘솔에서 진행 문구로 죽지 않게
    try:
        _s.reconfigure(errors="replace")
    except (AttributeError, OSError):
        pass

# 모니터별 배율이 다른 환경에서 좌표가 어긋나지 않도록 **import 시점에** 선언한다.
try:
    ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
except Exception:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

u32 = ctypes.WinDLL("user32", use_last_error=True)
g32 = ctypes.WinDLL("gdi32", use_last_error=True)
k32 = ctypes.WinDLL("kernel32", use_last_error=True)

VK = {"ctrl": 0x11, "shift": 0x10, "alt": 0x12, "enter": 0x0D, "esc": 0x1B,
      "a": 0x41, "c": 0x43, "d": 0x44, "s": 0x53,
      "up": 0x26, "down": 0x28, "left": 0x25, "right": 0x27,
      "home": 0x24, "end": 0x23, "pgup": 0x21, "pgdn": 0x22}


# ── 창 ──────────────────────────────────────────────────────────────────────

def windows_of(pid, titled_only=True):
    """그 프로세스의 보이는 최상위 창들. titled_only=False 면 메뉴 같은 무제목 팝업까지."""
    out = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def cb(hw, _):
        p = wintypes.DWORD()
        u32.GetWindowThreadProcessId(hw, ctypes.byref(p))
        if p.value == pid and u32.IsWindowVisible(hw):
            n = u32.GetWindowTextLengthW(hw)
            t = ctypes.create_unicode_buffer(n + 1)
            u32.GetWindowTextW(hw, t, n + 1)
            c = ctypes.create_unicode_buffer(256)
            u32.GetClassNameW(hw, c, 256)
            if t.value or not titled_only:
                out.append((hw, t.value, c.value))
        return True

    u32.EnumWindows(cb, 0)
    return out


def wait_window(pid, timeout=60):
    end = time.time() + timeout
    while time.time() < end:
        w = windows_of(pid)
        if w:
            return w[0]
        time.sleep(0.3)
    raise TimeoutError("창이 뜨지 않았다")


def rect(h):
    r = wintypes.RECT()
    u32.GetWindowRect(h, ctypes.byref(r))
    return r.left, r.top, r.right, r.bottom


def focus(h, tries=15):
    """앞으로 올리고 **확인**한다. 못 올리면 False — 그러면 입력을 보내지 않는다.

    윈도우는 백그라운드 프로세스의 SetForegroundWindow 를 막는다. 입력 큐를 현재
    포그라운드 스레드에 붙이면(AttachThreadInput) 같은 입력 데스크톱으로 취급돼 허용된다.
    """
    if u32.IsIconic(h):              # 최소화일 때만 되돌린다 — 무조건 하면 최대화가 풀린다
        u32.ShowWindow(h, 9)
    for _ in range(tries):
        if u32.GetForegroundWindow() == h:
            return True
        fg = u32.GetForegroundWindow()
        tid_fg = u32.GetWindowThreadProcessId(fg, None)
        tid_me = k32.GetCurrentThreadId()
        attached = u32.AttachThreadInput(tid_me, tid_fg, True)
        try:
            u32.BringWindowToTop(h)
            u32.SetForegroundWindow(h)
            u32.SetActiveWindow(h)
        finally:
            if attached:
                u32.AttachThreadInput(tid_me, tid_fg, False)
        time.sleep(0.3)
        if u32.GetForegroundWindow() == h:
            return True
        u32.SwitchToThisWindow(h, True)
        time.sleep(0.3)
    return u32.GetForegroundWindow() == h


def _assert_front(h):
    cur = u32.GetForegroundWindow()
    if cur != h:
        raise RuntimeError(f"포그라운드가 대상 창이 아니다(hwnd={cur}, 기대={h}) — 입력 보류")


# ── 입력 ────────────────────────────────────────────────────────────────────

def key(h, name, ctrl=False, shift=False, alt=False):
    _assert_front(h)
    if ctrl:
        u32.keybd_event(VK["ctrl"], 0, 0, 0)
    if shift:
        u32.keybd_event(VK["shift"], 0, 0, 0)
    if alt:
        u32.keybd_event(VK["alt"], 0, 0, 0)
    u32.keybd_event(VK[name], 0, 0, 0)
    time.sleep(0.03)
    u32.keybd_event(VK[name], 0, 2, 0)
    if alt:
        u32.keybd_event(VK["alt"], 0, 2, 0)
    if shift:
        u32.keybd_event(VK["shift"], 0, 2, 0)
    if ctrl:
        u32.keybd_event(VK["ctrl"], 0, 2, 0)
    time.sleep(0.25)


def click(h, x, y, right=False):
    """창 기준 (x, y) 클릭. 좌표는 매번 현재 rect 로 다시 푼다(창이 움직일 수 있다)."""
    _assert_front(h)
    L, T, _, _ = rect(h)
    click_screen(L + x, T + y, right=right)


def click_screen(x, y, right=False):
    u32.SetCursorPos(int(x), int(y))
    time.sleep(0.15)
    down, up = (0x0008, 0x0010) if right else (0x0002, 0x0004)
    u32.mouse_event(down, 0, 0, 0, 0)
    time.sleep(0.05)
    u32.mouse_event(up, 0, 0, 0, 0)
    time.sleep(0.35)


def wheel(h, x, y, notches):
    _assert_front(h)
    L, T, _, _ = rect(h)
    u32.SetCursorPos(L + x, T + y)
    time.sleep(0.15)
    for _ in range(abs(notches)):
        u32.mouse_event(0x0800, 0, 0, ctypes.c_int(-120 if notches < 0 else 120), 0)
        time.sleep(0.02)
    time.sleep(0.5)


# ── 클립보드 ────────────────────────────────────────────────────────────────

def _open_clipboard(tries=20):
    for _ in range(tries):
        if u32.OpenClipboard(0):
            return
        time.sleep(0.1)
    raise RuntimeError("클립보드를 열지 못했다")


def clipboard_text():
    u32.GetClipboardData.restype = ctypes.c_void_p      # 기본값은 64비트 핸들을 자른다
    k32.GlobalLock.argtypes = [ctypes.c_void_p]
    k32.GlobalLock.restype = ctypes.c_void_p
    k32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    _open_clipboard()
    try:
        hnd = u32.GetClipboardData(13)                  # CF_UNICODETEXT
        if not hnd:
            return ""
        p = k32.GlobalLock(hnd)
        if not p:
            return ""
        try:
            return ctypes.wstring_at(p)
        finally:
            k32.GlobalUnlock(hnd)
    finally:
        u32.CloseClipboard()


def set_clipboard(text):
    """powershell Set-Clipboard 는 쓰지 않는다 — 그 프로세스가 포커스를 가져간다."""
    k32.GlobalAlloc.restype = ctypes.c_void_p
    k32.GlobalLock.argtypes = [ctypes.c_void_p]
    k32.GlobalLock.restype = ctypes.c_void_p
    k32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    u32.SetClipboardData.argtypes = [wintypes.UINT, ctypes.c_void_p]
    u32.SetClipboardData.restype = ctypes.c_void_p
    buf = ctypes.create_unicode_buffer(text)
    size = ctypes.sizeof(buf)
    _open_clipboard()
    try:
        u32.EmptyClipboard()
        hg = k32.GlobalAlloc(0x0042, size)              # GMEM_MOVEABLE|GMEM_ZEROINIT
        p = k32.GlobalLock(hg)
        ctypes.memmove(p, buf, size)
        k32.GlobalUnlock(hg)
        u32.SetClipboardData(13, hg)
    finally:
        u32.CloseClipboard()


# ── 캡처 ────────────────────────────────────────────────────────────────────

def shot(h, path):
    """PrintWindow(PW_RENDERFULLCONTENT) — 앞에 없어도 내용이 찍힌다."""
    try:
        from PIL import Image
    except ImportError:
        return None
    L, T, R, B = rect(h)
    w, hh = R - L, B - T
    hdc = u32.GetWindowDC(h)
    mdc = g32.CreateCompatibleDC(hdc)
    bmp = g32.CreateCompatibleBitmap(hdc, w, hh)
    g32.SelectObject(mdc, bmp)
    u32.PrintWindow(h, mdc, 2)

    class BMI(ctypes.Structure):
        _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                    ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                    ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                    ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                    ("biClrImportant", wintypes.DWORD)]

    bi = BMI(ctypes.sizeof(BMI), w, -hh, 1, 32, 0, 0, 0, 0, 0, 0)
    buf = ctypes.create_string_buffer(w * hh * 4)
    g32.GetDIBits(mdc, bmp, 0, hh, buf, ctypes.byref(bi), 0)
    img = Image.frombuffer("RGBA", (w, hh), buf, "raw", "BGRA", 0, 1).convert("RGB")
    img.save(path)
    g32.DeleteObject(bmp)
    g32.DeleteDC(mdc)
    u32.ReleaseDC(h, hdc)
    return path


# ── 확인 절차 ───────────────────────────────────────────────────────────────

def installed_exe():
    return os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs",
                        "ExcelMerge", "ExcelMerge.exe")


def _writable_copy(src, dst):
    shutil.copy2(src, dst)
    os.chmod(dst, stat.S_IWRITE | stat.S_IREAD)    # P4V 읽기 전용 속성이 따라온다
    return dst


def _sha(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()


def _crash_exceptions():
    """crash.log 의 **예외** 기록 수. 파일 크기를 보면 안 된다 — 앱은 켤 때마다
    '세션 시작' 줄을 쓰므로 정상 실행에도 크기가 는다(처음에 그렇게 틀렸다)."""
    p = os.path.join(os.environ.get("APPDATA", ""), "ExcelMerge", "crash.log")
    try:
        with open(p, encoding="utf-8", errors="replace") as f:
            return f.read().count("미처리 예외")
    except OSError:
        return 0


def find_cell_box_y(h, x_ratio=0.14):
    """셀값란의 세로 중심을 **화면에서 찾아낸다**.

    창 높이에 대한 비율로 잡으면 안 된다 — 셀값란은 높이가 고정(기본 4줄)이고 표만
    늘어나므로, 최대화하면 비율이 어긋나 클릭이 표에 떨어진다(처음에 그렇게 틀려
    셀값란 대신 표가 복사됐다). 경로칸 아래 첫 흰 블록을 직접 찾는다.
    """
    try:
        from PIL import Image
    except ImportError:
        return None
    tmp = os.path.join(tempfile.gettempdir(), "uicheck_probe.png")
    if not shot(h, tmp):
        return None
    im = Image.open(tmp)
    W, H = im.size
    x = int(W * x_ratio)
    run_start = None
    for y in range(int(H * 0.10), int(H * 0.45)):
        white = im.getpixel((x, y))[:3] == (255, 255, 255)
        if white and run_start is None:
            run_start = y
        elif not white and run_start is not None:
            # 높이로 가린다. 바로 위 경로칸도 흰 한 줄짜리(약 24px)라, 아무 흰
            # 블록이나 집으면 경로칸을 눌러 **파일 경로**가 복사된다(실제로 그렇게
            # 틀렸다). 셀값란은 기본 4줄이라 60px 을 넘는다.
            if y - run_start >= 40:
                return (run_start + y) // 2
            run_start = None
    return None


def _popups(pid, main):
    return [w for w in windows_of(pid, titled_only=False) if w[0] != main]


def run(a_src, b_src, exe, expect_version, outdir):
    fails = []

    def ok(name, cond, detail=""):
        print(f"  {'OK  ' if cond else '실패'} {name}{(' — ' + detail) if detail else ''}")
        if not cond:
            fails.append(name)
        return cond

    a = _writable_copy(a_src, os.path.join(outdir, "A" + os.path.splitext(a_src)[1]))
    b = _writable_copy(b_src, os.path.join(outdir, "B" + os.path.splitext(b_src)[1]))
    sha_a0, sha_b0 = _sha(a), _sha(b)
    crash0 = _crash_exceptions()

    print(f"\n▶ 실행  {exe}")
    proc = subprocess.Popen([exe, a, b])
    h, title, _ = wait_window(proc.pid)
    time.sleep(12)                                  # 로드 + 비교
    print(f"  창 제목: {title}")
    if expect_version:
        ok("버전", title.strip().endswith("v" + str(expect_version)), title)

    if not focus(h):
        print("  !! 포그라운드를 얻지 못했다(다른 앱이 쓰는 중) — 입력 확인은 건너뛴다")
        proc.kill()
        return ["포그라운드 확보 실패"]

    L, T, R, B = rect(h)
    W, H = R - L, B - T

    # 1) 셀값란 Ctrl+C — 고른 글자가 복사돼야 한다(예전엔 파일 경로가 복사됐다)
    print("▶ 셀값란 Ctrl+C")
    set_clipboard("<UICHECK>")
    focus(h)
    key(h, "down", alt=True)                        # 첫 변경 셀로
    time.sleep(0.8)
    box_y = find_cell_box_y(h)
    if ok("셀값란을 찾음", box_y is not None, f"y={box_y}"):
        focus(h)
        click(h, int(W * 0.14), box_y)
        time.sleep(0.4)
        key(h, "a", ctrl=True)
        key(h, "c", ctrl=True)
        got = clipboard_text()
        ok("고른 글자가 복사됨", bool(got.strip()) and "<UICHECK>" not in got, repr(got[:40]))
        ok("경로가 아님", os.path.basename(a) not in got)
        # 표 복사(TSV)와 구분된다 — 셀값란은 한 칸의 값이라 탭이 없다
        ok("표가 아니라 셀값란이 복사됨", "	" not in got, repr(got[:40]))

    # 2) 병합 준비 → Ctrl+S
    print("▶ 병합 준비 → Ctrl+S")
    focus(h)
    click(h, int(W * 0.14), int(H * 0.30))          # A 패널 표
    time.sleep(0.4)
    key(h, "a", ctrl=True)                          # 전체 선택
    time.sleep(0.6)
    menu = None
    for _ in range(4):
        focus(h)
        click(h, int(W * 0.14), int(H * 0.30), right=True)
        time.sleep(1.3)
        pops = [w for w in _popups(proc.pid, h) if "Popup" in w[2]]
        if pops:
            menu = pops[0]
            break
    if ok("우클릭 메뉴가 뜸", menu is not None):
        ml, mt, mr, mb = rect(menu[0])
        click_screen((ml + mr) // 2, mt + (mb - mt) // 4)   # 첫 항목 = A → B 병합 준비
        time.sleep(1.5)
        focus(h)
        key(h, "s", ctrl=True)
        time.sleep(6.0)
        dlgs = _popups(proc.pid, h)
        ok("저장 완료 안내", any("저장" in w[1] for w in dlgs), str([w[1] for w in dlgs]))
        ok("B 가 저장됨", _sha(b) != sha_b0)
        ok("A 는 그대로", _sha(a) == sha_a0)
        for hw, _t, _c in dlgs:
            if focus(hw):
                key(hw, "enter")
            time.sleep(1.0)

        # 3) 저장 뒤 Ctrl+S 재입력 — 오류 창이 뜨면 안 된다(v221 회귀)
        print("▶ 저장 뒤 Ctrl+S 재입력")
        bad = None
        for _ in range(2):
            if not focus(h):
                break
            key(h, "s", ctrl=True)
            time.sleep(2.0)
            extra = _popups(proc.pid, h)
            if extra:
                bad = extra
                break
        ok("오류 창 없음", bad is None, str([w[1] for w in (bad or [])]))

    shot(h, os.path.join(outdir, "final.png"))
    ok("새 예외 기록 없음", _crash_exceptions() == crash0)
    proc.kill()
    return fails


def main():
    ap = argparse.ArgumentParser(description="설치본을 실제로 몰아 확인한다")
    ap.add_argument("a_file", help="A 파일 — 읽기만 한다")
    ap.add_argument("b_file", help="B 파일 — 읽기만 한다")
    ap.add_argument("--exe", default="", help="실행할 exe(기본: 설치된 ExcelMerge)")
    ap.add_argument("--expect-version", default="", help="창 제목에서 확인할 버전(예: 224)")
    args = ap.parse_args()

    exe = args.exe or installed_exe()
    for p in (args.a_file, args.b_file, exe):
        if not os.path.isfile(p):
            print(f"✖ 없는 경로: {p}")
            return 1

    outdir = tempfile.mkdtemp(prefix="uicheck_")
    atexit.register(shutil.rmtree, outdir, True)
    print(f"작업 폴더: {outdir}   (원본은 읽기만 한다)")

    fails = run(args.a_file, args.b_file, exe, args.expect_version, outdir)
    print()
    if fails:
        print(f"✖ 어긋남 {len(fails)}건: {', '.join(fails)}")
        return 1
    print("✔ 모두 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
