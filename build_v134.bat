@echo off
cd /d "%~dp0"
python -m PyInstaller ExcelMerge_v134.spec
echo BUILD_EXIT_CODE=%ERRORLEVEL%
