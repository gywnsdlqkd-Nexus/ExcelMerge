@echo off
cd /d "%~dp0"
pyinstaller ExcelDiffMerge_v116.spec
echo BUILD_EXIT_CODE=%ERRORLEVEL%
