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
