@echo off
setlocal enabledelayedexpansion
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

set "INPUT_FOLDER="
set /p INPUT_FOLDER="Path to MP4 folder (Enter for default 'input' folder): "

if defined INPUT_FOLDER (
    python "%~dp0clip_cutter.py" --input "!INPUT_FOLDER!"
) else (
    python "%~dp0clip_cutter.py"
)

echo.
pause
