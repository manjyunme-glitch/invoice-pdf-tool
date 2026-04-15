@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

set "ENTRY_SCRIPT="
for %%F in (*v5.py) do (
    set "ENTRY_SCRIPT=%%F"
)

echo ========================================
echo   Invoice PDF Tool v5.2.1 - Install deps
echo ========================================
echo.

echo [1/2] Install core dependencies...
python -m pip install pandas openpyxl
echo.

echo [2/2] Install optional UI dependencies...
python -m pip install ttkbootstrap
python -m pip install tkinterdnd2

echo.
echo ========================================
echo   Installation finished
if defined ENTRY_SCRIPT (
    echo   Run: python !ENTRY_SCRIPT!
) else (
    echo   Run: python your-entry-script.py
)
echo ========================================
pause
