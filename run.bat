@echo off
cd /d "%~dp0"
where pythonw >nul 2>&1
if %errorlevel%==0 (
    start "" pythonw excel_diff_merge.py %*
) else (
    start "" pyw excel_diff_merge.py %*
)
