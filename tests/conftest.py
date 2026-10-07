# -*- coding: utf-8 -*-
"""pytest 공통 설정 — 경로/Qt 오프스크린/QApplication 싱글턴을 한 곳에서 세팅.

기존 테스트 파일은 각자 `sys.path.insert`·`QT_QPA_PLATFORM=offscreen`·
`QApplication.instance() or QApplication([])` 를 반복했다. pytest 실행 시엔
이 conftest 가 그 역할을 대신하며, 개별 파일을 `python tests/xxx.py` 로 직접 실행하는
모드도 그대로 동작하도록 기존 부트스트랩은 남겨 두었다(중복이나 무해).
"""
import os
import sys

# Qt 는 import 시점에 플랫폼 플러그인을 고르므로 PyQt5 import 전에 설정해야 한다.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

import pytest


@pytest.fixture(scope="session")
def qapp():
    """세션 1회 QApplication. Qt 위젯/모델을 다루는 테스트에서 인자로 받아 쓴다."""
    from PyQt5.QtWidgets import QApplication
    app = QApplication.instance() or QApplication([])
    yield app

@pytest.fixture(autouse=True)
def isolated_appdata(tmp_path, monkeypatch):
    """모든 테스트에 **자기 APPDATA** 를 준다 — 전역 자동 적용.

    앱은 키 위치·검사 제외 열·로그를 %APPDATA%/ExcelMerge 에 쓴다. 격리하지 않으면
    테스트가 실제 사용자 설정을 덮어쓴다.

    예전엔 각 테스트 파일이 알아서 막았는데, 다섯 파일은 **모듈 수준에서**
    `os.environ["APPDATA"] = tempfile.mkdtemp(...)` 를 했다. 그게 두 가지 문제를 만들었다.

      · **되돌아오지 않는다.** 한 번 바뀌면 그 뒤로 수집되는 모든 테스트가 그 값을
        쓴다. 이 집은 PyQt5 가 사용자 site-packages(%APPDATA% 아래)에 깔려 있어서,
        자식 프로세스를 띄우는 테스트가 PyQt5 를 못 찾고 깨졌다 — 코드는 멀쩡한데
        **수집 순서에 따라** 결과가 달라졌다.
      · **치우지 않는다.** mkdtemp 로 만들기만 하고 지우지 않아 임시 폴더가 쌓였다
        (실측: 1549개, 357MB).

    tmp_path 는 테스트마다 새로 나오고 pytest 가 오래된 것을 알아서 지운다.
    monkeypatch 는 테스트가 끝나면 환경변수를 되돌린다. 둘 다 해결된다.

    개별 파일이 자기 APPDATA 를 또 지정해도 문제없다 — 나중 것이 이긴다.
    """
    d = tmp_path / "appdata"
    d.mkdir(exist_ok=True)
    monkeypatch.setenv("APPDATA", str(d))
    return d

