@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

set "ENTRY_SCRIPT=发票处理工具v6.py"

echo ========================================
echo   Invoice PDF Tool v6.1.1 - Install deps
echo ========================================
echo.

echo [1/2] Install core dependencies...
python -m pip install -r requirements.txt
echo.

echo [2/2] Verify dependencies...
python -m pip check

echo.
echo ========================================
echo   Installation finished
if exist "%ENTRY_SCRIPT%" (
    echo   Run: python !ENTRY_SCRIPT!
) else (
    echo   Run: python your-entry-script.py
)
echo ========================================
pause
