@echo off
cd /d "%~dp0"
python -m PyInstaller ExcelMerge_v129.spec
echo BUILD_EXIT_CODE=%ERRORLEVEL%
