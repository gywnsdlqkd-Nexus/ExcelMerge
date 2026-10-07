@echo off
cd /d "%~dp0"
REM 버전 무관 단일 빌드 스크립트 — 버전은 ExcelMerge.spec 이 excelmerge/__init__.py 에서 읽는다.
REM --noconfirm: dist/ExcelMerge 가 남아 있어도 묻지 않고 덮어쓴다(onedir).
python -m PyInstaller --noconfirm ExcelMerge.spec
set BUILD_RC=%ERRORLEVEL%
echo BUILD_EXIT_CODE=%BUILD_RC%
if not "%BUILD_RC%"=="0" (
    pause
    exit /b %BUILD_RC%
)
REM 코드 서명(인증서가 환경변수로 설정돼 있으면 서명, 없으면 건너뜀 — 빌드는 막지 않음).
python sign.py dist\ExcelMerge\ExcelMerge.exe
REM 설치 파일은 release.py 가 만든다(ISCC + installer.iss). 여기서는 앱 폴더까지만.
