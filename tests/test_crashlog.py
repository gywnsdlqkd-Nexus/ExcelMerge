# -*- coding: utf-8 -*-
"""crashlog 로테이션 테스트 — crash.log 가 상한을 넘으면 .1 로 밀려 무한 증가를 막는다."""
from excelmerge import crashlog


def test_rotate_when_over_limit(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    p = crashlog.log_path()
    import os
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        f.write("x" * (crashlog._MAX_BYTES + 1))     # 상한 초과

    crashlog._rotate_if_large()

    assert not os.path.exists(p), "상한 초과인데 로테이트 안 됨"
    assert os.path.exists(p + ".1"), ".1 백업 미생성"


def test_no_rotate_when_small(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    p = crashlog.log_path()
    import os
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        f.write("small")

    crashlog._rotate_if_large()

    assert os.path.exists(p), "작은 로그를 잘못 로테이트함"
    assert not os.path.exists(p + ".1")
