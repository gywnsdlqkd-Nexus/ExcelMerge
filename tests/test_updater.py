# -*- coding: utf-8 -*-
"""자동 업데이트 순수 로직 + 매니페스트 조회 테스트.

실행: python tests/test_updater.py  (또는 pytest)
"""
import os
import sys
import json
import tempfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import excelmerge.updater as upd
from excelmerge.updater import (
    _parse_manifest, is_newer, build_update_bat, UpdateCheckWorker,
    gdrive_id_from_url, gdrive_direct_url, _parse_github_release, _source,
)


def test_parse_github_release():
    j = json.dumps({
        "tag_name": "v175",
        "body": "변경점 A\n변경점 B",
        "assets": [
            {"name": "notes.txt", "browser_download_url": "https://x/notes.txt"},
            {"name": "ExcelMerge_v175.exe",
             "browser_download_url": "https://github.com/o/r/releases/download/v175/ExcelMerge_v175.exe",
             "digest": "sha256:DEAD"},
        ],
    }).encode("utf-8")
    m = _parse_github_release(j)
    assert m["version"] == "175", m
    assert m["url"].endswith("ExcelMerge_v175.exe")
    assert m["sha256"] == "dead"
    assert "변경점 A" in m["notes"]
    print("PASS test_parse_github_release")


def test_source_selection(monkeypatch=None):
    # 설정 파일 간섭 배제
    upd._config = lambda: {}
    orig_repo, orig_url = upd.GITHUB_REPO, upd.MANIFEST_URL
    try:
        upd.GITHUB_REPO, upd.MANIFEST_URL = "o/r", ""
        kind, url = _source()
        assert kind == "github" and url == "https://api.github.com/repos/o/r/releases/latest", (kind, url)
        upd.GITHUB_REPO, upd.MANIFEST_URL = "", "https://x/latest.json"
        assert _source() == ("manifest", "https://x/latest.json")
        upd.GITHUB_REPO, upd.MANIFEST_URL = "", ""
        assert _source() is None
    finally:
        upd.GITHUB_REPO, upd.MANIFEST_URL = orig_repo, orig_url
    print("PASS test_source_selection")


def test_gdrive_url_helpers():
    fid = "1AbC_dEf-123"
    assert gdrive_id_from_url(f"https://drive.google.com/file/d/{fid}/view?usp=sharing") == fid
    assert gdrive_id_from_url(f"https://drive.google.com/uc?export=download&id={fid}") == fid
    assert gdrive_id_from_url("https://example.com/x.exe") == ""
    assert gdrive_direct_url(fid) == f"https://drive.google.com/uc?export=download&id={fid}"
    print("PASS test_gdrive_url_helpers")


def test_is_newer():
    assert is_newer("174", "173") is True
    assert is_newer("173", "173") is False
    assert is_newer("173", "174") is False
    # 비숫자 폴백(문자열 비교)
    assert is_newer("1.2.1", "1.2.0") is True
    print("PASS test_is_newer")


def test_parse_manifest():
    data = json.dumps({"version": "174", "url": "https://x/e.exe",
                       "sha256": "ABC", "notes": "n"}).encode("utf-8")
    m = _parse_manifest(data)
    assert m["version"] == "174" and m["url"] == "https://x/e.exe"
    assert m["sha256"] == "abc"          # 소문자 정규화
    assert m["notes"] == "n"
    # version 없으면 오류
    bad = False
    try:
        _parse_manifest(b'{"url":"x"}')
    except ValueError:
        bad = True
    assert bad, "version 누락인데 통과"
    print("PASS test_parse_manifest")


def test_build_update_bat_quotes_paths():
    bat = build_update_bat(r"C:\Temp\new one.exe", r"C:\Program Files\App\ExcelMerge.exe")
    assert '"C:\\Temp\\new one.exe"' in bat
    assert '"C:\\Program Files\\App\\ExcelMerge.exe"' in bat
    assert "move /y" in bat and "start " in bat
    print("PASS test_build_update_bat_quotes_paths")


def test_build_update_bat_clears_bootloader_env():
    """재실행 전 PyInstaller 부트로더 변수를 비워야 한다(자기 교체 후 DLL 로드 실패 방지)."""
    bat = build_update_bat(r"C:\t\new.exe", r"C:\app\ExcelMerge.exe")
    for var in ("_MEIPASS2", "_PYI_ARCHIVE_FILE", "_PYI_APPLICATION_HOME_DIR"):
        assert f'set "{var}="' in bat, f"{var} 미제거"
    # 변수 제거는 start 이전에 있어야 효력이 있다
    assert bat.index('set "_MEIPASS2="') < bat.index('start ""')
    print("PASS test_build_update_bat_clears_bootloader_env")


def test_apply_update_spawns_with_clean_env(monkeypatch):
    """apply_update가 부트로더 변수를 제거한 env로 배치를 띄우는지."""
    monkeypatch.setattr(upd.sys, "frozen", True, raising=False)
    monkeypatch.setattr(upd.sys, "executable", os.path.join(tempfile.gettempdir(), "ExcelMerge.exe"))
    monkeypatch.setenv("_MEIPASS2", r"C:\Temp\_MEI12345")
    monkeypatch.setenv("_PYI_ARCHIVE_FILE", r"C:\Temp\app.exe")
    monkeypatch.setenv("KEEP_ME", "1")

    captured = {}

    def fake_popen(args, **kw):
        captured.update(kw)
        return object()

    monkeypatch.setattr(upd.subprocess, "Popen", fake_popen)
    ok = upd.apply_update(os.path.join(tempfile.gettempdir(), "new.exe"))
    assert ok is True
    env = captured["env"]
    assert env is not None, "env를 명시적으로 전달하지 않음(상속 위험)"
    assert "_MEIPASS2" not in env and "_PYI_ARCHIVE_FILE" not in env
    assert env.get("KEEP_ME") == "1", "일반 환경변수는 유지돼야 함"
    # 콘솔 창 번쩍임 방지: CREATE_NO_WINDOW 만(상호배타 DETACHED_PROCESS 제외) + 숨김 STARTUPINFO
    flags = captured.get("creationflags", 0)
    assert flags & 0x08000000, "CREATE_NO_WINDOW 미설정"
    assert not (flags & 0x00000008), "DETACHED_PROCESS 는 CREATE_NO_WINDOW 와 상호배타 — 제거 필요"
    si = captured.get("startupinfo")
    assert si is not None and si.wShowWindow == upd.subprocess.SW_HIDE
    print("PASS test_apply_update_spawns_with_clean_env")


def test_check_worker_file_url():
    """file:// 매니페스트를 UpdateCheckWorker가 읽어 파싱 결과를 방출."""
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtCore import QEventLoop, QTimer, QUrl
    app = QApplication.instance() or QApplication([])

    d = tempfile.mkdtemp()
    p = os.path.join(d, "latest.json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump({"version": "999", "url": "https://x/e.exe"}, f)
    url = QUrl.fromLocalFile(p).toString()   # file:///...

    got = {}
    loop = QEventLoop()
    w = UpdateCheckWorker(url)
    w.done.connect(lambda m: (got.update(m or {}), loop.quit()))
    w.start()
    QTimer.singleShot(5000, loop.quit)
    loop.exec_()
    assert got.get("version") == "999", got
    print("PASS test_check_worker_file_url")


def main():
    test_is_newer()
    test_parse_manifest()
    test_parse_github_release()
    test_source_selection()
    test_gdrive_url_helpers()
    test_build_update_bat_quotes_paths()
    test_build_update_bat_clears_bootloader_env()
    test_check_worker_file_url()
    print("ALL UPDATER TESTS PASS")


if __name__ == "__main__":
    main()
