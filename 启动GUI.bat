@echo off
cd /d "%~dp0"

REM 优先使用 Python Launcher；如果不存在，再尝试直接调用 pythonw。
pyw -3 "%~dp0图形界面.py"
if not errorlevel 1 goto end

pythonw "%~dp0图形界面.py"

:end
