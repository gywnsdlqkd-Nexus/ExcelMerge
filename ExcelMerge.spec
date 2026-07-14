# -*- mode: python ; coding: utf-8 -*-
import glob
import os
import re
from PyInstaller.utils.hooks import collect_all

# 버전은 excelmerge/__init__.py 의 __version__ 단일 출처에서 읽는다.
# (버전마다 .spec 을 복제하던 관행 제거 — v182~ 이 파일 하나로 빌드)
# SPECPATH 는 PyInstaller 가 주입하는 이 spec 파일이 있는 디렉터리.
with open(os.path.join(SPECPATH, "excelmerge", "__init__.py"), encoding="utf-8") as _f:  # noqa: F821
    _VERSION = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', _f.read()).group(1)

# python-calamine(Rust 확장)은 컴파일된 _python_calamine 모듈을 포함하므로
# collect_all로 바이너리/서브모듈을 모두 수집해야 프리즈 후에도 동작한다.
_cal_datas, _cal_binaries, _cal_hidden = collect_all('python_calamine')
# orjson(Rust 확장, v178~ 빠른 JSON 파싱)도 컴파일된 확장 모듈이라 동일하게 수집한다.
_orj_datas, _orj_binaries, _orj_hidden = collect_all('orjson')

# ── UCRT(Universal C Runtime) 번들 (v179~ 필수) ────────────────────────────────
# python314.dll은 api-ms-win-crt-*.dll(UCRT)에 의존한다. PyInstaller 6.x는 UCRT가 OS에
# 있다고 가정해 번들에서 제외하는데, UCRT가 없는 사용자 PC에서는
# 'Failed to load Python DLL … 지정된 모듈을 찾을 수 없습니다' 오류가 난다(v178 실사례).
# → Windows SDK의 UCRT 재배포 DLL(ucrtbase.dll + api-ms-win-crt-*.dll)을 번들 루트에 포함.
def _find_ucrt_dir():
    base = r"C:\Program Files (x86)\Windows Kits\10\Redist"
    cands = sorted(glob.glob(os.path.join(base, "*", "ucrt", "DLLs", "x64")), reverse=True)
    return cands[0] if cands else None

_ucrt_dir = _find_ucrt_dir()
if not _ucrt_dir:
    raise SystemExit(
        "[빌드 중단] UCRT 재배포 DLL 폴더를 찾지 못했습니다.\n"
        "Windows SDK(UCRT 재배포 구성요소)를 설치하세요 — "
        r"'C:\Program Files (x86)\Windows Kits\10\Redist\<ver>\ucrt\DLLs\x64' 경로가 필요합니다.")
_ucrt_binaries = [(f, '.') for f in glob.glob(os.path.join(_ucrt_dir, '*.dll'))]

a = Analysis(
    ['excel_diff_merge.py'],
    pathex=[],
    binaries=_cal_binaries + _orj_binaries + _ucrt_binaries,
    # 헤더/탭 아이콘(images/)을 번들의 images/ 폴더로 포함(_MEIPASS/images).
    datas=[('images/app_icon.ico', 'images'),
           ('images/Key.png', 'images'), ('images/Exception.png', 'images'),
           ('images/Reset.png', 'images'),
           ('images/Excel.png', 'images'), ('images/JSON.png', 'images')]
          + _cal_datas + _orj_datas,
    # 값 전용(v163~) — 수식 평가(formulas/schedula/numpy) 경로 없음.
    hiddenimports=['python_calamine', 'python_calamine._python_calamine', 'orjson']
                  + _cal_hidden + _orj_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=f'ExcelMerge_v{_VERSION}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['images/app_icon.ico'],
)
