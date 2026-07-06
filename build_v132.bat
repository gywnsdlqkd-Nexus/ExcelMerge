@echo off
cd /d "%~dp0"
python -m PyInstaller ExcelMerge_v132.spec
echo BUILD_EXIT_CODE=%ERRORLEVEL%
