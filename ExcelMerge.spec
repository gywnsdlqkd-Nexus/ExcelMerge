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
    # 번들에 안 들어가야 할 것들 — 프로그램을 켤 때마다 이걸 다 푸느라 느렸다.
    # (실측: 번들 157MB 중 77MB 가 쓰지 않는 것이었고, 빼니 exe 60.9→30.2MB,
    #  켜는 시간 6.4→5.2초. numpy·PIL 을 막고 전체 테스트 820개 통과를 확인했다.)
    #   numpy  — openpyxl 이 '있으면 쓰는' 선택적 의존. 우리 코드는 쓰지 않는다(29MB).
    #   PIL    — openpyxl 의 이미지 지원. 우리는 값만 읽고 XML 만 고친다(12MB).
    #   QtQuick/QtQml 등 — 순수 QtWidgets 앱이라 쓰지 않는다.
    excludes=[
        'numpy', 'PIL', 'pandas', 'scipy', 'matplotlib',
        'PyQt5.QtQuick', 'PyQt5.QtQml', 'PyQt5.Qt3DCore',
        'PyQt5.QtWebEngineWidgets', 'PyQt5.QtMultimedia',
        'tkinter', 'pydoc_data',
    ],
    noarchive=False,
    optimize=0,
)
# 쓰지 않는 Qt 바이너리 — hiddenimport 분석으로는 안 빠지고 바이너리로 딸려온다.
# opengl32sw.dll 하나가 20MB 다. 순수 QtWidgets 앱은 래스터 엔진으로 그리므로
# OpenGL 경로를 타지 않는다. (GPU 드라이버가 부실한 PC 를 만나면 이 줄을 되돌릴 것.)
_DROP_BIN = {n.lower() for n in (
    'opengl32sw.dll',        # 소프트웨어 OpenGL 폴백   20.0 MB
    'd3dcompiler_47.dll',    # Direct3D 셰이더 컴파일러  4.0 MB
    'libGLESv2.dll', 'libEGL.dll',   # ANGLE              3.2 MB
    'Qt5Quick.dll', 'Qt5Qml.dll', 'Qt5QmlModels.dll',   # QML    7.4 MB
    'Qt5Network.dll',        # 업데이트 확인은 urllib 를 쓴다
    'Qt5WebSockets.dll', 'Qt5DBus.dll',
)}
a.binaries = [b for b in a.binaries
              if os.path.basename(b[0]).lower() not in _DROP_BIN]

pyz = PYZ(a.pure)

# ── onefile → onedir (v217~) ────────────────────────────────────────────────
# onefile 은 **켤 때마다** 번들 전체를 %TEMP% 에 푼다. 그게 시작 시간의 가장 큰 몫이었다
# (실측: 같은 코드로 onefile 4.6초 vs onedir 2.7초). 배포는 Inno Setup 설치 파일로 한다
# (installer.iss) — 폴더를 그대로 두면 "압축 안 풀고 exe 만 실행" 사고가 나는데,
# 설치 파일이 그 문제를 없앤다.
#
# exe 이름에 버전을 붙이지 않는다. 설치 경로가 고정돼야 P4V 에 등록한 diff 툴 경로가
# 버전마다 바뀌지 않는다. 버전은 설치 파일 이름(ExcelMerge_Setup_v<N>.exe)에만 붙인다.
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ExcelMerge',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['images/app_icon.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,          # UPX 는 PATH 에 없어 어차피 적용되지 않았다(실측: 켜나 끄나 동일)
    upx_exclude=[],
    name='ExcelMerge',  # → dist/ExcelMerge/ExcelMerge.exe
)
