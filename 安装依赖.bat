@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   发票处理工具箱 v5.1 - 安装依赖
echo ========================================
echo.

echo [1/2] 安装核心依赖...
python -m pip install pandas openpyxl
echo.

echo [2/2] 安装可选依赖...
python -m pip install ttkbootstrap
python -m pip install tkinterdnd2

echo.
echo ========================================
echo   安装完成
echo   现在可以运行：
echo   python 发票处理工具v5.py
echo ========================================
pause
