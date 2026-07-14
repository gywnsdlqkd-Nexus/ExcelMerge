@echo off
cd /d "%~dp0"
REM 버전 무관 단일 빌드 스크립트 — 버전은 ExcelMerge.spec 이 excelmerge/__init__.py 에서 읽는다.
python -m PyInstaller ExcelMerge.spec
echo BUILD_EXIT_CODE=%ERRORLEVEL%
if errorlevel 1 pause
