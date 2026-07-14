# -*- coding: utf-8 -*-
"""전역 키 위치 저장(prefs) 라운드트립/견고성 테스트.

APPDATA 를 임시 폴더로 돌려 실제 파일 경로에 영향 없이 검증한다.
실행: python tests/test_prefs.py
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excelmerge import prefs


def _with_tmp_appdata(fn):
    old = os.environ.get("APPDATA")
    with tempfile.TemporaryDirectory() as d:
        os.environ["APPDATA"] = d
        try:
            fn(d)
        finally:
            if old is None:
                os.environ.pop("APPDATA", None)
            else:
                os.environ["APPDATA"] = old


def test_roundtrip():
    def body(d):
        # 저장 전 기본값
        assert prefs.load_key_prefs() == (0, 0)
        prefs.save_key_prefs(1, 3)
        assert os.path.isfile(os.path.join(d, "ExcelMerge", "prefs.json"))
        assert prefs.load_key_prefs() == (1, 3)
        # ROW 순서(키 열 -1)도 보존
        prefs.save_key_prefs(0, -1)
        assert prefs.load_key_prefs() == (0, -1)
    _with_tmp_appdata(body)
    print("PASS test_roundtrip")


def test_defaults_on_missing_and_corrupt():
    def body(d):
        # 파일 없음 → 기본값
        assert prefs.load_key_prefs() == (0, 0)
        p = os.path.join(d, "ExcelMerge", "prefs.json")
        os.makedirs(os.path.dirname(p), exist_ok=True)
        # 깨진 JSON → 기본값
        with open(p, "w", encoding="utf-8") as f:
            f.write("{ not json ")
        assert prefs.load_key_prefs() == (0, 0)
        # 이상값(음수 행, dict 아님) → 기본값
        import json
        with open(p, "w", encoding="utf-8") as f:
            json.dump({"key_row": -5, "key_col": 2}, f)
        assert prefs.load_key_prefs() == (0, 0)
        with open(p, "w", encoding="utf-8") as f:
            json.dump([1, 2], f)
        assert prefs.load_key_prefs() == (0, 0)
    _with_tmp_appdata(body)
    print("PASS test_defaults_on_missing_and_corrupt")


if __name__ == "__main__":
    test_roundtrip()
    test_defaults_on_missing_and_corrupt()
    print("ALL PREFS TESTS PASS")
