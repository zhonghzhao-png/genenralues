@echo off
chcp 65001
cd /d "%~dp0"

REM 检查 Python 窗口启动器是否存在；不使用 >nul，避免部分 Windows 安全设置拦截 nul 设备访问。
where pyw
if errorlevel 1 (
    start "" pythonw "图形界面.py"
) else (
    start "" pyw -3 "图形界面.py"
)
