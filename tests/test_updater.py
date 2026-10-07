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
    """경로에 공백이 있어도 안전해야 한다 — Program Files 밑에 깔릴 수도 있다."""
    bat = build_update_bat(r"C:\Temp\setup one.exe",
                           r"C:\Program Files\App\ExcelMerge.exe")
    assert '"C:\\Temp\\setup one.exe"' in bat
    assert '"C:\\Program Files\\App\\ExcelMerge.exe"' in bat
    print("PASS test_build_update_bat_quotes_paths")


def test_build_update_bat_runs_the_installer_silently():
    """v217~ 배포물은 설치 파일이다 — 조용히 돌려야 창이 튀지 않는다.

    예전엔 단일 exe 를 `move /y` 로 자기 자신에 덮어썼다. onedir 은 파일이 아니라
    폴더라 그 방법을 쓸 수 없고, 실행 중인 앱을 닫고 파일을 바꾸는 일은 설치 파일이
    대신한다(installer.iss 의 CloseApplications=force).
    """
    bat = build_update_bat(r"C:\t\ExcelMerge_Setup_v217.exe", r"C:\app\ExcelMerge.exe")
    for flag in ("/SILENT", "/NORESTART", "/SUPPRESSMSGBOXES"):
        assert flag in bat, f"{flag} 없음 — 설치 창이 사용자에게 튄다"
    assert "move /y" not in bat, "파일 덮어쓰기 방식이 남아 있다"


def test_build_update_bat_restarts_the_app_even_if_setup_fails():
    """설치가 실패해도 앱은 다시 띄운다 — 쓰던 사람이 빈손으로 남으면 안 된다.

    실패하면 이전 버전이 그대로 남아 있으므로 띄우는 게 맞다. 그래서 배치에 조건
    분기가 없어야 한다 — `if errorlevel` 로 감싸면 실패했을 때 아무것도 안 뜬다.
    """
    bat = build_update_bat(r"C:\t\setup.exe", r"C:\app\ExcelMerge.exe")
    lines = [ln.strip() for ln in bat.splitlines() if ln.strip()]
    run_i = next(i for i, ln in enumerate(lines) if "/SILENT" in ln)
    start_i = next(i for i, ln in enumerate(lines) if ln.startswith('start ""'))
    assert run_i < start_i, "설치보다 먼저 앱을 띄우려 한다"
    assert not any(ln.lower().startswith(("if ", "goto ")) for ln in lines), \
        "조건 분기가 있으면 설치 실패 시 앱이 안 뜰 수 있다"


def test_build_update_bat_cleans_up_after_itself():
    """받아 둔 설치 파일과 배치 자신을 지운다 — %TEMP% 에 쌓이면 안 된다."""
    bat = build_update_bat(r"C:\t\setup.exe", r"C:\app\ExcelMerge.exe")
    assert 'del "C:\\t\\setup.exe"' in bat
    assert 'del "%~f0"' in bat


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
