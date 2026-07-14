@echo off
cd /d "%~dp0"
where python >nul 2>&1
if %errorlevel%==0 (
    python excel_diff_merge.py
) else (
    py excel_diff_merge.py
)
if errorlevel 1 pause
