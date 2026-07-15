# -*- coding: utf-8 -*-
"""코드 서명 헬퍼(sign.py) — 인증서 환경변수 기반 명령 구성/건너뜀 로직."""
import sign


def test_skip_when_no_cert(monkeypatch):
    monkeypatch.delenv("EXCELMERGE_SIGN_THUMBPRINT", raising=False)
    monkeypatch.delenv("EXCELMERGE_SIGN_PFX", raising=False)
    assert sign.build_sign_cmd("signtool.exe", "x.exe") is None


def test_thumbprint_cmd(monkeypatch):
    monkeypatch.delenv("EXCELMERGE_SIGN_PFX", raising=False)
    monkeypatch.setenv("EXCELMERGE_SIGN_THUMBPRINT", "ABCD1234")
    cmd = sign.build_sign_cmd("signtool.exe", "dist/app.exe")
    assert cmd[:2] == ["signtool.exe", "sign"]
    assert "/sha1" in cmd and cmd[cmd.index("/sha1") + 1] == "ABCD1234"
    assert "/fd" in cmd and cmd[cmd.index("/fd") + 1] == "sha256"
    assert "/tr" in cmd and "/td" in cmd          # 타임스탬프
    assert cmd[-1] == "dist/app.exe"


def test_pfx_cmd_with_password(monkeypatch):
    monkeypatch.delenv("EXCELMERGE_SIGN_THUMBPRINT", raising=False)
    monkeypatch.setenv("EXCELMERGE_SIGN_PFX", r"C:\certs\cs.pfx")
    monkeypatch.setenv("EXCELMERGE_SIGN_PFX_PASSWORD", "secret")
    cmd = sign.build_sign_cmd("signtool.exe", "app.exe")
    assert "/f" in cmd and cmd[cmd.index("/f") + 1] == r"C:\certs\cs.pfx"
    assert "/p" in cmd and cmd[cmd.index("/p") + 1] == "secret"


def test_thumbprint_precedence_and_timestamp_override(monkeypatch):
    # 지문이 pfx보다 우선하고, 타임스탬프 override가 반영된다.
    monkeypatch.setenv("EXCELMERGE_SIGN_THUMBPRINT", "TP")
    monkeypatch.setenv("EXCELMERGE_SIGN_PFX", r"C:\x.pfx")
    monkeypatch.setenv("EXCELMERGE_SIGN_TIMESTAMP", "http://ts.example/rfc3161")
    cmd = sign.build_sign_cmd("signtool.exe", "app.exe")
    assert "/sha1" in cmd and "/f" not in cmd
    assert cmd[cmd.index("/tr") + 1] == "http://ts.example/rfc3161"
