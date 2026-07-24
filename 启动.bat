@echo off
chcp 65001 >nul
title 中国水果产区数据大屏 - Flask 服务

echo ============================================
echo   中国水果产区数据大屏
echo   Flask 后端服务启动中...
echo ============================================
echo.

cd /d "%~dp0"

:: ── 清理端口 5000 残留进程 ──
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5000.*LISTENING" 2^>nul') do (
    echo [清理] 终止端口 5000 残留进程...
    taskkill /F /PID %%a >nul 2>&1
)

:: ── 检查 Python ──
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)

:: ── 检查依赖（静默） ──
python -c "import flask, openpyxl" 2>nul
if %errorlevel% neq 0 (
    echo [提示] 正在安装依赖...
    pip install flask openpyxl -q
)

echo [信息] Excel 数据: fruit_data.xlsx
echo [信息] 大屏文件: 中国水果产区数据大屏.html
echo ============================================
echo.

:: 等待服务器就绪后自动打开浏览器
start "" http://localhost:5000/
python server.py

pause
