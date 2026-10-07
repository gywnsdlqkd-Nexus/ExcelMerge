# -*- coding: utf-8 -*-
"""시작할 때 무엇을 **안** 불러오는가, 그리고 그 대가를 어디서 치르는가.

`from openpyxl.utils import get_column_letter` 한 줄이 openpyxl 패키지 전체를 끌어왔다
(모듈 430개, 1.14초). 쓰는 건 좌표 변환 함수 둘뿐이라 colref 로 떼어 내고, 값 읽기용
openpyxl 은 함수 안으로 내렸다 — 켜는 시간이 그만큼 줄었다.

**그런데 비용은 사라지지 않고 옮겨간다.** openpyxl 은 평소엔 안 쓰지만(값 읽기는
calamine 이 한다) .xls/.xlsb/암호 걸린 .xlsx 의 시트 이름을 얻을 때는 쓴다. 그 호출은
GUI 스레드에서 나서, 거기서 처음 import 하면 창이 **0.9초 멈췄다**(실측 0.868s).

그래서 창이 뜬 뒤 백그라운드에서 미리 올려 둔다. 이 파일은 둘 다 고정한다 —
시작 경로에 없을 것, 그리고 쓸 때쯤엔 이미 올라와 있을 것.
"""
import os
import subprocess
import sys
import textwrap

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def _run(code: str) -> str:
    """새 프로세스에서 돌린다 — sys.modules 는 전역이라 한 프로세스로는 못 본다.

    자식에게 **이 프로세스의 sys.path 를 그대로** 물려준다. APPDATA 를 물려받게 두면
    안 된다 — 여러 테스트 모듈이 모듈 수준에서 APPDATA 를 임시 폴더로 바꿔 버리는데,
    이 집의 PyQt5 는 사용자 site-packages(%APPDATA% 아래)에 깔려 있어서 자식이
    PyQt5 를 못 찾는다. 멀쩡한 코드인데 수집 순서에 따라 테스트가 깨진다.
    """
    env = {**os.environ,
           "QT_QPA_PLATFORM": "offscreen",
           "PYTHONPATH": os.pathsep.join(p for p in sys.path if p)}
    out = subprocess.run([sys.executable, "-c", textwrap.dedent(code)],
                         capture_output=True, text=True, cwd=ROOT, env=env)
    assert out.returncode == 0, out.stdout + out.stderr
    return out.stdout.strip()


# ── 1. 시작 경로에 없어야 하는 것 ───────────────────────────────────────────

@pytest.mark.parametrize("module", ["openpyxl", "numpy", "PIL"])
def test_the_startup_path_does_not_load(module):
    """앱 모듈을 불러오는 것만으로 이것들이 딸려오면 안 된다.

    numpy·PIL 은 번들에서도 뺐다(ExcelMerge.spec). 시작 경로에 다시 들어오면
    번들에 없어서 **import 가 터진다** — 느려지는 정도가 아니라 못 켜진다.
    """
    got = _run(f"""
        import sys
        import excelmerge.main_window          # noqa: F401
        print({module!r} in sys.modules)
    """)
    assert got == "False", f"{module} 이(가) 시작 경로에 딸려온다"


def test_reading_a_normal_file_does_not_need_openpyxl(tmp_path):
    """값 읽기의 주 경로는 calamine 이다 — 보통은 openpyxl 을 한 번도 안 쓴다."""
    import openpyxl
    p = tmp_path / "a.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["TID", "NAME"])
    wb.active.append(["k1", "칼"])
    wb.save(str(p))
    got = _run(f"""
        import sys, os, tempfile
        os.environ['APPDATA'] = tempfile.mkdtemp()
        from excelmerge.loaders import load_values_any
        rows = load_values_any({str(p)!r})
        print(len(rows), 'openpyxl' in sys.modules)
    """)
    assert got == "2 False", got


# ── 2. 옮겨간 비용은 백그라운드에서 치른다 ─────────────────────────────────

def test_warm_returns_without_waiting(monkeypatch):
    """선로드는 **기다리지 않는다** — 기다리면 시작 경로로 되돌아간 것과 같다."""
    from excelmerge import loaders
    made = {}

    class FakeThread:
        def __init__(self, target=None, name=None, daemon=None):
            made["target"] = target
            made["daemon"] = daemon

        def start(self):
            made["started"] = True

        def join(self, *a):                       # 불리면 안 된다
            made["joined"] = True

    monkeypatch.setattr(loaders, "_warm_started", False)
    monkeypatch.setattr(loaders.threading, "Thread", FakeThread)
    loaders.warm_fallback_readers()
    assert made.get("started") is True, "스레드를 시작하지 않았다"
    assert made.get("joined") is None, "선로드를 기다렸다 — 시작이 다시 느려진다"
    assert made.get("daemon") is True, "데몬이 아니면 종료를 붙잡는다"


def test_warm_runs_only_once(monkeypatch):
    from excelmerge import loaders
    count = {"n": 0}

    class FakeThread:
        def __init__(self, target=None, name=None, daemon=None):
            count["n"] += 1

        def start(self):
            pass

    monkeypatch.setattr(loaders, "_warm_started", False)
    monkeypatch.setattr(loaders.threading, "Thread", FakeThread)
    loaders.warm_fallback_readers()
    loaders.warm_fallback_readers()
    loaders.warm_fallback_readers()
    assert count["n"] == 1, f"선로드 스레드를 {count['n']}번 만들었다"


def test_warm_signals_completion_even_when_it_fails(monkeypatch):
    """실패해도 신호는 올린다 — 안 그러면 기다리는 쪽이 영영 막힌다."""
    from excelmerge import loaders
    import threading
    monkeypatch.setattr(loaders, "_warm_started", False)
    monkeypatch.setattr(loaders, "warm_done", threading.Event())
    real_import = __builtins__["__import__"] if isinstance(__builtins__, dict) \
        else __builtins__.__import__

    def boom(name, *a, **k):
        if name.startswith("openpyxl"):
            raise ImportError("선로드 실패 재현")
        return real_import(name, *a, **k)

    monkeypatch.setattr("builtins.__import__", boom)
    loaders.warm_fallback_readers()
    assert loaders.warm_done.wait(10), "실패했는데 신호를 안 올렸다"


def test_the_warm_up_removes_the_stall(tmp_path):
    """선로드가 끝난 뒤엔 .xls 를 열어도 멈추지 않는다.

    수정 전에는 이 호출이 GUI 스레드에서 0.868초를 먹었다. 기준을 넉넉히 0.3초로
    둔다 — 그래도 선로드가 빠지면 바로 걸린다(한 자릿수 배 차이다).
    """
    p = tmp_path / "fake.xls"
    p.write_bytes(bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 2048)
    got = _run(f"""
        import os, sys, tempfile, time
        os.environ['APPDATA'] = tempfile.mkdtemp()
        from excelmerge import loaders
        loaders.warm_fallback_readers()
        assert loaders.warm_done.wait(60), '선로드가 끝나지 않았다'
        t = time.perf_counter()
        loaders.list_sheet_names({str(p)!r})
        print(f'{{time.perf_counter() - t:.4f}}')
    """)
    assert float(got) < 0.3, f".xls 첫 열기가 {got}초 — 선로드가 안 먹었다"


def test_the_entry_point_warms_after_showing_the_window():
    """진입점이 **창을 띄운 뒤** 선로드를 건다 — 앞에 두면 창이 그만큼 늦게 뜬다."""
    src = open(os.path.join(ROOT, "excel_diff_merge.py"), encoding="utf-8").read()
    assert "warm_fallback_readers()" in src, "진입점이 선로드를 걸지 않는다"
    assert src.index("win.show()") < src.index("warm_fallback_readers()"), \
        "선로드를 창 띄우기 **전에** 걸었다 — 시작이 다시 느려진다"
