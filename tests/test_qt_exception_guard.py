# -*- coding: utf-8 -*-
"""Qt 슬롯에서 샌 예외가 테스트를 **통과시키지 못하게** 한다.

C++ 가 부른 슬롯(closeEvent·타이머·시그널) 안에서 예외가 처리되지 않으면 PyQt5 는
sys.excepthook 을 부르고 그대로 지나간다. 트레이스백은 pytest 가 가로챈 출력에
담겼다가 테스트가 통과하면 버려지고, 그 테스트는 **초록색으로 끝난다.**

실제로 그 구멍으로 v220 의 회귀가 빠져나갔다 — 저장이 끝난 워커 참조로 isRunning()
을 불러 RuntimeError 가 났는데, 유닛 테스트는 아무 표시 없이 통과했고 실기에서야
오류 창으로 드러났다.

conftest 의 pytest_configure + no_unhandled_qt_exceptions 가 그 구멍을 막는다.
여기서는 그 장치가 실제로 도는지 **자식 pytest 를 돌려** 확인한다 — 같은 프로세스
안에서 흉내 내면 정작 Qt 경계를 넘는 경로를 안 밟는다.
"""
import os
import subprocess
import sys
import textwrap

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

CHILD = '''# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import QWidget


class Boom(QWidget):
    def closeEvent(self, ev):
        raise RuntimeError("표식_슬롯예외")


def test_leaks_from_a_slot(qapp):
    w = Boom()
    w.show()
    qapp.processEvents()
    w.close()                 # closeEvent 안에서 터진다
    qapp.processEvents()


def test_clean(qapp):
    w = QWidget()
    w.show()
    qapp.processEvents()
    w.close()
'''


@pytest.fixture(scope="module")
def child_run(tmp_path_factory):
    """conftest 를 그대로 복사해 자식 pytest 를 한 번만 돌린다(느리므로 재사용)."""
    d = tmp_path_factory.mktemp("guard")
    (d / "conftest.py").write_text(
        open(os.path.join(ROOT, "tests", "conftest.py"), encoding="utf-8").read(),
        encoding="utf-8")
    (d / "test_child.py").write_text(textwrap.dedent(CHILD), encoding="utf-8")
    # PyQt5 는 사용자 site-packages(%APPDATA% 아래)에 있다. 이 테스트는 APPDATA 가
    # 격리된 상태로 도므로, 부모의 sys.path 를 넘겨 줘야 자식이 PyQt5 를 찾는다.
    env = dict(os.environ, PYTHONPATH=os.pathsep.join(p for p in sys.path if p),
               PYTHONIOENCODING="utf-8")
    # 자식은 UTF-8 로 쓴다(PYTHONIOENCODING). text=True 는 로케일(cp949)로 읽어
    # 한글 트레이스백에서 UnicodeDecodeError 가 난다 — 인코딩을 명시한다.
    return subprocess.run([sys.executable, "-m", "pytest", str(d), "-q"],
                          capture_output=True, env=env, timeout=600,
                          encoding="utf-8", errors="replace")


def test_the_guard_reports_the_leaking_test(child_run):
    """샌 테스트가 이름으로 지목돼야 한다 — 어디서 터졌는지 모르면 소용이 없다."""
    assert "test_leaks_from_a_slot" in child_run.stdout, child_run.stdout[-2000:]


def test_the_guard_shows_the_original_traceback(child_run):
    assert "표식_슬롯예외" in child_run.stdout, child_run.stdout[-2000:]


def test_the_run_does_not_come_back_all_green(child_run):
    """이게 핵심이다 — 예전엔 이런 실행이 '전부 통과' 로 끝났다."""
    assert child_run.returncode != 0, child_run.stdout[-2000:]


def test_a_clean_test_is_not_blamed(child_run):
    """앞 테스트가 흘린 것까지 뒤집어쓰면 범인이 엉뚱해진다.

    (예외는 teardown 에서 판정하므로 call 단계는 통과로 찍히고 ERROR 로 뜬다.)
    """
    blamed = [ln for ln in child_run.stdout.splitlines()
              if ln.startswith(("ERROR ", "FAILED "))]
    assert blamed, child_run.stdout[-2000:]
    assert any("test_leaks_from_a_slot" in ln for ln in blamed), blamed
    assert not any("test_clean" in ln for ln in blamed), blamed
