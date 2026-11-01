@echo off
chcp 65001
cd /d "%~dp0"
python "图形界面_PySide6.py"
if errorlevel 1 (
  echo.
  echo 启动失败。请确认已安装 PySide6：python -m pip install PySide6
  pause
)
