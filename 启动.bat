@echo off
chcp 65001 >nul
title 中国水果产区数据大屏

echo.
echo ╔════════════════════════════════════╗
echo ║   中国水果产区数据大屏            ║
echo ║   正在启动服务...                  ║
echo ╚════════════════════════════════════╝
echo.

cd /d "%~dp0"

:: 检查 Python 是否可用
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 检查依赖
python -c "import flask; import openpyxl; import zhdate" >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] 正在安装依赖...
    pip install flask openpyxl zhdate -q
    if %errorlevel% neq 0 (
        echo [错误] 依赖安装失败，请手动运行：pip install flask openpyxl zhdate
        pause
        exit /b 1
    )
)

:: 在后台启动 Flask 服务器
echo [启动] 正在启动服务器 http://localhost:5000/
start "Flask-Server" /MIN python server.py

:: 等待服务器就绪（最多等 15 秒）
echo [等待] 等待服务器就绪...
set /a count=0
:waitloop
timeout /t 1 /nobreak >nul
set /a count+=1
python -c "import urllib.request; urllib.request.urlopen('http://localhost:5000/api/health')" >nul 2>&1
if %errorlevel% equ 0 goto ready
if %count% lss 15 goto waitloop

echo [警告] 服务器启动超时，请检查是否端口被占用
pause
exit /b 1

:ready
echo [就绪] 服务器已启动（耗时 %count% 秒）
echo [启动] 正在打开浏览器...
start "" http://localhost:5000/
echo.
echo ══════════════════════════════════════
echo   服务运行中，请勿关闭 Flask 窗口
echo   在浏览器中编辑数据后点击「更新数据」
echo   按浏览器中的按钮操作即可
echo ══════════════════════════════════════
echo.
pause
