@echo off
chcp 65001 >nul
title 中国水果产区数据大屏 - Flask 服务

echo ============================================
echo   中国水果产区数据大屏
echo   Flask 后端服务启动中...
echo ============================================
echo.

cd /d "%~dp0"

REM 检查 Python 是否可用
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    echo.
    pause
    exit /b 1
)

REM 检查依赖
python -c "import flask, openpyxl, zhdate" >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] 正在安装依赖...
    pip install flask openpyxl zhdate -q
)

echo [信息] Excel 数据文件: fruit_data.xlsx
echo [信息] HTML 大屏文件: 中国水果产区数据大屏.html
echo.
echo 启动成功后，请访问:
echo   http://localhost:5000/
echo.
echo 按 Ctrl+C 停止服务
echo ============================================
echo.

REM 启动 Flask 并自动打开浏览器
start "" http://localhost:5000/
python server.py

pause
