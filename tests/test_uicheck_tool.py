# -*- coding: utf-8 -*-
"""실기 확인 도구(uicheck.py)가 지켜야 할 약속 — 파일을 읽어 고정한다.

이 도구는 **사용자의 실제 키보드·마우스**를 쓴다. 그래서 두 가지가 코드에 남아 있어야
한다. 둘 다 실제로 사고를 냈거나 막았다.

  · 입력 전 포그라운드 확인 — 없으면 사용자가 쓰던 다른 앱(언리얼 에디터)에 키가
    들어간다. 실제로 두 번 막혔다.
  · 원본은 읽기만 — 비교 대상은 빌드 폴더의 실데이터다. 거기에 쓰면 안 된다.

그리고 이 도구가 존재하는 이유 자체(손으로 하다 세 번 다시 만들었다)가 사라지지
않도록, 절차 문서에서 가리키고 있는지도 본다. 빌드 없이 파일만 읽으므로 빠르다.
"""
import os
import re

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _read(name):
    with open(os.path.join(ROOT, name), encoding="utf-8") as f:
        return f.read()


@pytest.fixture(scope="module")
def src():
    return _read("uicheck.py")


# ── 사용자의 입력 장치를 쓴다 — 그 안전장치 ────────────────────────────────

def test_input_is_guarded_by_a_foreground_check(src):
    """키/클릭을 보내기 전에 대상 창이 앞에 있는지 **반드시** 확인한다."""
    for fn in ("def key(", "def click(", "def wheel("):
        i = src.index(fn)
        body = src[i:i + 700]
        assert "_assert_front(h)" in body, f"{fn} 가 포그라운드를 확인하지 않는다"


def test_the_foreground_check_raises_rather_than_warns(src):
    i = src.index("def _assert_front(")
    body = src[i:src.index("\n\n\n", i)]
    assert "raise" in body, "앞에 없을 때 그냥 넘어간다 — 다른 앱에 입력이 들어간다"


def test_focus_does_not_unmaximize(src):
    """무조건 SW_RESTORE 를 부르면 최대화가 풀린다 — 최소화일 때만 되돌려야 한다."""
    i = src.index("def focus(")
    body = src[i:src.index("def _assert_front(")]
    assert "IsIconic" in body, "최소화 여부를 보지 않고 ShowWindow 를 부른다"


# ── 원본을 건드리지 않는다 ──────────────────────────────────────────────────

def test_inputs_are_copied_before_use(src):
    """빌드 폴더의 실데이터를 비교에 쓴다 — 거기에 쓰면 안 된다."""
    i = src.index("def run(")
    body = src[i:src.index("def main(")]
    assert "_writable_copy(a_src" in body and "_writable_copy(b_src" in body, \
        "원본 경로를 그대로 앱에 넘긴다"
    assert "outdir" in body, "출력 폴더를 쓰지 않는다"


def test_the_copy_clears_the_read_only_attribute(src):
    """P4V 가 빌드 파일을 읽기 전용으로 둔다 — 안 풀면 저장 검사가 가짜로 실패한다."""
    i = src.index("def _writable_copy(")
    body = src[i:src.index("def _sha(")]
    assert "S_IWRITE" in body


def test_the_work_folder_is_temporary_and_cleaned(src):
    i = src.index("def main(")
    body = src[i:]
    assert "mkdtemp" in body and "atexit" in body, "임시 폴더가 쌓인다"


# ── 좌표·환경 함정 ──────────────────────────────────────────────────────────

def test_dpi_awareness_is_declared_at_import(src):
    """모니터마다 배율이 다르면 비인식 프로세스는 가상화 좌표를 받아 클릭이 빗나간다."""
    head = src[:src.index("u32 = ctypes.WinDLL")]
    assert "SetProcessDpiAwarenessContext" in head or "SetProcessDpiAware" in head


def test_the_cell_box_is_found_by_height_not_ratio(src):
    """비율로 잡으면 최대화 때 표에 떨어지고, 아무 흰 블록이나 집으면 경로칸을 누른다.

    둘 다 실제로 틀렸다 — 처음엔 표가, 다음엔 파일 경로가 복사됐다.
    """
    i = src.index("def find_cell_box_y(")
    body = src[i:src.index("def _popups(")]
    m = re.search(r"if y - run_start >= (\d+)", body)
    assert m, "높이로 가르지 않는다"
    assert int(m.group(1)) >= 40, f"기준 {m.group(1)}px — 경로칸(약 24px)이 걸린다"


def test_the_clipboard_handle_type_is_declared(src):
    """기본 restype(c_int)은 64비트 핸들을 잘라 액세스 위반이 난다."""
    i = src.index("def clipboard_text(")
    body = src[i:src.index("def set_clipboard(")]
    assert "GetClipboardData.restype" in body


def test_it_does_not_shell_out_for_the_clipboard(src):
    """powershell Set-Clipboard 는 그 프로세스가 포커스를 가져가 다음 입력을 깨뜨린다.

    (설명문에는 그 이름이 함정으로 적혀 있으므로, 이름이 아니라 **외부 프로세스를
    띄우는지**를 본다.)
    """
    i = src.index("def set_clipboard(")
    body = src[i:src.index("# ── 캡처")]
    assert "subprocess" not in body, "클립보드를 쓰려고 외부 프로세스를 띄운다"


# ── 절차에서 가리키는가 ─────────────────────────────────────────────────────

def test_release_notes_point_at_the_tool():
    rel = _read("RELEASE.md")
    assert "uicheck.py" in rel, "RELEASE.md 가 실기 확인 도구를 가리키지 않는다"


def test_it_fails_loudly(src):
    """어긋나면 종료 코드 1 — 조용히 0 을 돌려주면 자동화에 쓸 수 없다."""
    i = src.index("def main(")
    body = src[i:]
    assert "return 1" in body and "sys.exit(main())" in src
