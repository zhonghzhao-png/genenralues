@echo off
chcp 65001 >nul
cd /d "%~dp0"
where pyw >nul 2>nul
if %errorlevel%==0 (
    start "" pyw -3 "图形界面.py"
) else (
    start "" pythonw "图形界面.py"
)
