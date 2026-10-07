# -*- coding: utf-8 -*-
"""설치 파일 배포 구조가 서로 어긋나지 않는가 — 파일을 실제로 읽어 확인한다.

v217 부터 배포물이 단일 exe 가 아니라 **설치 파일**이다(PyInstaller onedir + Inno Setup).
켤 때마다 번들을 %TEMP% 에 푸는 비용을 없애려는 것이고(실측: onefile 4.6초 vs
onedir 2.7초), 폴더 배포의 "압축 안 풀고 exe 만 실행" 사고는 설치 파일이 막는다.

이 전환은 **네 곳이 같은 이름을 보고 있어야** 동작한다 — spec·iss·release·make_release.
하나만 어긋나도 "빌드는 됐는데 올릴 자산이 없다"가 되고, 그건 릴리스 도중에야 드러난다.
여기서 그 합의를 고정한다. 빌드 없이 파일만 읽으므로 빠르다.
"""
import os
import re

import pytest

from excelmerge import __version__

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _read(name: str) -> str:
    with open(os.path.join(ROOT, name), encoding="utf-8") as f:
        return f.read()


# ── 1. spec 이 onedir 인가 ──────────────────────────────────────────────────

def test_spec_builds_a_folder_not_a_single_file():
    """onefile 로 되돌아가면 시작이 다시 2초 느려진다."""
    spec = _read("ExcelMerge.spec")
    assert "COLLECT(" in spec, "COLLECT 가 없다 — onefile 로 돌아갔다"
    assert "exclude_binaries=True" in spec, "EXE 가 바이너리를 안고 있다(onefile)"


def test_the_app_exe_name_has_no_version():
    """설치 경로가 버전마다 바뀌면 P4V 에 등록한 diff 툴 경로가 매번 깨진다."""
    spec = _read("ExcelMerge.spec")
    m = re.search(r"COLLECT\(.*?name='([^']+)'", spec, re.S)
    assert m and m.group(1) == "ExcelMerge", f"COLLECT 이름: {m and m.group(1)}"
    assert "name=f'ExcelMerge_v{_VERSION}'" not in spec, "exe 이름에 버전이 붙어 있다"


# ── 2. 설치 파일 설정 ───────────────────────────────────────────────────────

def test_installer_installs_without_admin():
    """관리자 권한을 요구하면 동료가 받아서 바로 못 쓰고, 조용한 업데이트도 막힌다."""
    iss = _read("installer.iss")
    assert "PrivilegesRequired=lowest" in iss
    assert "{autopf}" in iss, "설치 경로가 사용자 영역이 아니다"


def test_installer_closes_a_running_app():
    """업데이트는 실행 중인 앱을 닫고 덮어써야 한다 — 이게 폴더 교체 문제의 해답이다."""
    iss = _read("installer.iss")
    assert "CloseApplications=force" in iss


def test_installer_app_id_is_fixed():
    """AppId 가 바뀌면 기존 설치를 못 찾아 두 벌이 깔린다."""
    iss = _read("installer.iss")
    m = re.search(r"^AppId=\{\{([0-9A-Fa-f-]{36})", iss, re.M)
    assert m, "AppId 가 고정 GUID 가 아니다"
    assert "{#MyVersion}" not in m.group(0), "AppId 에 버전이 섞였다"


def test_installer_packs_the_whole_folder():
    iss = _read("installer.iss")
    assert "recursesubdirs" in iss, "_internal 하위가 안 들어간다"
    assert r"dist\{#MyAppName}\*" in iss


def test_installer_output_name_matches_what_release_publishes():
    """이름이 어긋나면 '빌드는 됐는데 올릴 자산이 없다'가 된다."""
    iss = _read("installer.iss")
    m = re.search(r"^OutputBaseFilename=(.+)$", iss, re.M)
    assert m, "OutputBaseFilename 이 없다"
    assert m.group(1).strip() == "ExcelMerge_Setup", \
        f"설치 파일 이름에 버전이 붙었다: {m.group(1)!r}"
    assert "OutputDir=dist" in iss
    mk = _read("make_release.py")
    assert '"ExcelMerge_Setup.exe"' in mk


# ── 3. 릴리스 절차가 그 둘을 이어 주는가 ───────────────────────────────────

def test_release_compiles_the_installer_after_building():
    rel = _read("release.py")
    assert "installer.iss" in rel, "릴리스가 설치 파일을 만들지 않는다"
    assert rel.index("PyInstaller") < rel.index("installer.iss"), \
        "빌드 전에 설치 파일을 만들려 한다"


def test_release_signs_the_app_before_packing_it():
    """설치 파일로 묶은 뒤 앱 exe 를 서명하면 묶인 사본은 미서명으로 남는다."""
    rel = _read("release.py")
    sign_app = rel.index('"ExcelMerge.exe"')
    pack = rel.index("installer.iss")
    assert sign_app < pack, "앱 서명이 설치 파일 생성보다 뒤에 있다"


def test_release_smoke_runs_the_installed_layout():
    rel = _read("release.py")
    assert '"dist", "ExcelMerge", "ExcelMerge.exe"' in rel, \
        "스모크가 아직 onefile 경로를 본다"


def test_iscc_is_found_or_release_stops():
    """설치 파일 없이 릴리스가 진행되면 안 된다 — 조용히 넘어가면 자산이 빈다."""
    rel = _read("release.py")
    assert "def iscc_path(" in rel
    body = rel[rel.index("def iscc_path("):rel.index("def exe_smoke(")]
    assert "fail(" in body, "ISCC 를 못 찾아도 그냥 진행한다"


# ── 4. 자동 업데이트가 설치 파일을 받는가 ──────────────────────────────────

def test_updater_prefers_the_setup_asset():
    """릴리스에 exe 가 여럿 올라가도 설치 파일을 골라야 한다."""
    from excelmerge.updater import _parse_github_release
    import json
    data = json.dumps({
        "tag_name": "v217",
        "body": "노트",
        "assets": [
            {"name": "ExcelMerge_debug.exe",
             "browser_download_url": "https://x/ExcelMerge_debug.exe"},
            {"name": "ExcelMerge_Setup.exe",
             "browser_download_url": "https://x/ExcelMerge_Setup.exe",
             "digest": "sha256:" + "a" * 64},
        ],
    }).encode("utf-8")
    m = _parse_github_release(data)
    assert m["url"].endswith("ExcelMerge_Setup.exe"), m["url"]
    assert m["sha256"] == "a" * 64
    assert m["version"] == "217"


def test_apply_update_reports_why_it_failed(monkeypatch, caplog):
    """조용히 False 만 돌려주면 안 된다 — 실제로 이 자리에서 오타를 놓쳤다.

    인자 이름을 바꾸며 생긴 NameError 가 '업데이트가 그냥 안 되네' 로만 보였다.
    """
    import excelmerge.updater as upd
    monkeypatch.setattr(upd.sys, "frozen", True, raising=False)

    def boom(*a, **k):
        raise OSError("디스크 가득 참")

    monkeypatch.setattr("builtins.open", boom)
    with caplog.at_level("WARNING"):
        assert upd.apply_update("setup.exe") is False
    assert any("업데이트 적용 실패" in r.message for r in caplog.records), \
        "실패 이유를 남기지 않았다"


# ── 5. 버전이 한 곳에서만 온다 ─────────────────────────────────────────────

def test_version_comes_from_one_place():
    """iss 는 버전을 직접 적지 않고 ISCC /DMyVersion 으로 받는다."""
    iss = _read("installer.iss")
    assert "#ifndef MyVersion" in iss, "버전을 안 넘겨도 조용히 빌드된다"
    assert f"AppVersion={__version__}" not in iss, "iss 에 버전이 하드코딩됐다"
    rel = _read("release.py")
    assert "/DMyVersion=" in rel
