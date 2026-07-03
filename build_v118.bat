@echo off
cd /d "%~dp0"
python -m PyInstaller ExcelDiffMerge_v118.spec
echo BUILD_EXIT_CODE=%ERRORLEVEL%
