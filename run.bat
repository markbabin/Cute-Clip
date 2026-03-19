@echo off
echo === Clip Cutter ===
echo.

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python not found! Install from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)

pip install openpyxl >nul 2>&1

python "%~dp0clip_cutter.py" %*

echo.
pause
