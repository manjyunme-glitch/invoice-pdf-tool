@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

set "ENTRY_SCRIPT="
for %%F in (*v5.py) do (
    set "ENTRY_SCRIPT=%%~fF"
)

if not defined ENTRY_SCRIPT (
    echo [ERROR] Cannot find the GUI entry script.
    pause
    exit /b 1
)

set "OUTPUT_NAME=invoice-pdf-tool-v5.2.1"

echo ========================================
echo   Invoice PDF Tool v5.2.1 - Build EXE
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python 3.10+ is required.
    pause
    exit /b 1
)

echo [1/5] Install dependencies...
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [WARN] Dependency install reported issues. Continue building...
)

echo.
echo [2/5] Clean previous outputs...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
del /q "*.spec" 2>nul
for /d /r %%D in (__pycache__) do @if exist "%%D" rmdir /s /q "%%D"

echo.
echo [3/5] Build GUI EXE...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "%OUTPUT_NAME%" ^
    --noconfirm ^
    --clean ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.cell._writer ^
    --collect-all openpyxl ^
    "%ENTRY_SCRIPT%"
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. Please review the logs above.
    pause
    exit /b 1
)

if not exist "dist\%OUTPUT_NAME%.exe" (
    echo.
    echo [ERROR] Missing output file: dist\%OUTPUT_NAME%.exe
    pause
    exit /b 1
)

echo.
echo [4/5] Clean intermediate files...
if exist "build" rmdir /s /q "build"
del /q "*.spec" 2>nul
for /d /r %%D in (__pycache__) do @if exist "%%D" rmdir /s /q "%%D"

echo.
echo [5/5] Build completed
echo.
echo ========================================
echo   Output:
echo   dist\%OUTPUT_NAME%.exe
echo ========================================
echo.

explorer dist
pause
