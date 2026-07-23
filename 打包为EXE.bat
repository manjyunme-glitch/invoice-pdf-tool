@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "APP_VERSION=6.0.0"
set "SPEC_FILE=%~dp0invoice-pdf-tool-v6.spec"
set "OUTPUT_NAME=invoice-pdf-tool-v6.0.0-windows-x64"
set "OUTPUT_FILE=%~dp0dist\%OUTPUT_NAME%.exe"
set "HASH_FILE=%~dp0dist\%OUTPUT_NAME%.sha256.txt"
set "PYTHONUTF8=1"

echo ========================================
echo   Invoice PDF Tool v%APP_VERSION% - Release Build
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 goto :python_error

if not exist "%SPEC_FILE%" goto :spec_error
if not exist "%~dp0assets\invoice-pdf-tool-icon.ico" goto :icon_error

echo [1/4] Check release dependencies...
python -c "import PyInstaller, pandas, openpyxl, xlrd, xlwt, ttkbootstrap, tkinterdnd2"
if errorlevel 1 goto :dependency_error
python -m pip check
if errorlevel 1 goto :dependency_error

echo.
echo [2/4] Run complete regression suite...
python -X utf8 -m unittest discover -s tests -p "test_*.py"
if errorlevel 1 goto :test_error

echo.
echo [3/4] Build one-file Windows executable...
python -m PyInstaller --noconfirm --clean --distpath "%~dp0dist" --workpath "%~dp0build\v6.0.0" "%SPEC_FILE%"
if errorlevel 1 goto :build_error

if not exist "%OUTPUT_FILE%" goto :output_error

echo.
echo [4/4] Generate SHA-256 checksum...
python -c "import hashlib,sys; from pathlib import Path; source=Path(sys.argv[1]); target=Path(sys.argv[2]); digest=hashlib.file_digest(source.open('rb'), 'sha256').hexdigest(); target.write_text(digest + '  ' + source.name + '\n', encoding='ascii')" "%OUTPUT_FILE%" "%HASH_FILE%"
if errorlevel 1 goto :hash_error

python -c "import hashlib,sys; from pathlib import Path; source=Path(sys.argv[1]); expected=Path(sys.argv[2]).read_text(encoding='ascii').split()[0]; actual=hashlib.file_digest(source.open('rb'), 'sha256').hexdigest(); raise SystemExit(0 if actual == expected else 1)" "%OUTPUT_FILE%" "%HASH_FILE%"
if errorlevel 1 goto :hash_error

echo.
echo ========================================
echo   Build completed
echo   EXE:  %OUTPUT_FILE%
echo   SHA:  %HASH_FILE%
echo ========================================
exit /b 0

:python_error
echo [ERROR] Python 3.10 or newer is required.
goto :failed

:spec_error
echo [ERROR] Missing spec file: %SPEC_FILE%
goto :failed

:icon_error
echo [ERROR] Missing application icon.
goto :failed

:dependency_error
echo [ERROR] Release dependencies are incomplete.
echo         Run: python -m pip install -r requirements-build.txt
goto :failed

:test_error
echo [ERROR] Regression tests failed. EXE was not built.
goto :failed

:build_error
echo [ERROR] PyInstaller build failed.
goto :failed

:output_error
echo [ERROR] Expected EXE was not generated: %OUTPUT_FILE%
goto :failed

:hash_error
echo [ERROR] EXE exists, but checksum generation failed.
goto :failed

:failed
if not defined CI pause
exit /b 1
