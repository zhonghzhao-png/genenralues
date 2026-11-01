@echo off
setlocal
chcp 65001
cd /d "%~dp0"

echo ==========================================
echo 正在启动成绩分析 GUI
echo 当前目录：%CD%
echo ==========================================
echo.

if not exist "%~dp0图形界面.py" (
    echo [错误] 当前文件夹中找不到 图形界面.py
    echo 请确认启动GUI.bat与图形界面.py放在同一个文件夹。
    pause
    exit /b 1
)

where py
if errorlevel 1 (
    echo 未找到 Python Launcher，尝试使用 python...
    python "%~dp0图形界面.py"
) else (
    py -3 "%~dp0图形界面.py"
)

if errorlevel 1 (
    echo.
    echo [启动失败] 请查看上面的错误信息。
) else (
    echo.
    echo 程序已退出。
)

pause
endlocal
