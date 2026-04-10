@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   发票处理工具箱 v5.1 - 一键打包
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.10+
    pause
    exit /b 1
)

echo [1/5] 安装依赖...
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [警告] 依赖安装存在异常，将继续尝试打包
)

echo.
echo [2/5] 清理旧产物...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
del /q "*.spec" 2>nul
for /d /r %%D in (__pycache__) do @if exist "%%D" rmdir /s /q "%%D"

echo.
echo [3/5] 打包 GUI EXE...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "发票处理工具箱v5.1" ^
    --noconfirm ^
    --clean ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.cell._writer ^
    --collect-all openpyxl ^
    "发票处理工具v5.py"
if errorlevel 1 (
    echo.
    echo [错误] 打包失败，请检查上方日志
    pause
    exit /b 1
)

if not exist "dist\发票处理工具箱v5.1.exe" (
    echo.
    echo [错误] 未找到打包产物 dist\发票处理工具箱v5.1.exe
    pause
    exit /b 1
)

echo.
echo [4/5] 清理中间文件...
if exist "build" rmdir /s /q "build"
del /q "*.spec" 2>nul
for /d /r %%D in (__pycache__) do @if exist "%%D" rmdir /s /q "%%D"

echo.
echo [5/5] 打包完成
echo.
echo ========================================
echo   成功生成：
echo   dist\发票处理工具箱v5.1.exe
echo ========================================
echo.

explorer dist
pause
