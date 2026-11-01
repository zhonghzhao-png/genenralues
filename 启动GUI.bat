@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo Starting the score analysis GUI...
echo Working directory: %CD%
echo ==========================================
echo.

where py
if errorlevel 1 (
    echo Python Launcher was not found. Trying python...
    python -c "import pathlib,runpy; files=[p for p in pathlib.Path('.').glob('*.py') if 'class ScoreAnalysisGUI' in p.read_text(encoding='utf-8-sig')]; assert files, 'GUI script not found'; runpy.run_path(str(files[0]), run_name='__main__')"
) else (
    py -3 -c "import pathlib,runpy; files=[p for p in pathlib.Path('.').glob('*.py') if 'class ScoreAnalysisGUI' in p.read_text(encoding='utf-8-sig')]; assert files, 'GUI script not found'; runpy.run_path(str(files[0]), run_name='__main__')"
)

if errorlevel 1 (
    echo.
    echo [ERROR] GUI startup failed. Check the messages above.
) else (
    echo.
    echo GUI process has exited.
)

pause
endlocal
