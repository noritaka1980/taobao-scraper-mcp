@echo off
chcp 65001 >nul 2>&1
title TaobaoScraper API Server

echo ============================================
echo   TaobaoScraper API 启动脚本
echo ============================================
echo.

rem ── 1. 加载 .env 环境变量 ──
if exist "%~dp0.env" (
    echo [INFO] 正在加载 .env 配置...
    for /f "usebackq tokens=1,* delims==" %%A in ("%~dp0.env") do (
        rem 跳过注释行和空行
        echo %%A | findstr /b "#" >nul 2>&1
        if errorlevel 1 (
            if not "%%A"=="" set "%%A=%%B"
        )
    )
    echo [OK] .env 配置已加载
) else (
    echo [WARN] 未找到 .env 文件，将使用默认配置
    echo [WARN] 请复制 .env.example 为 .env 并修改 API Key！
    echo.
)

rem ── 2. 检查 Python ──
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] 未找到 Python，请先安装 Python 3.10+
    echo         下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

rem ── 3. 安装依赖 ──
if not exist "%~dp0.deps_api_installed" (
    echo [INFO] 首次运行，正在安装 API 依赖...
    pip install -r "%~dp0requirements_api.txt" -q --disable-pip-version-check
    if %errorlevel% neq 0 (
        echo [ERROR] 依赖安装失败，请检查网络连接
        pause
        exit /b 1
    )
    echo. > "%~dp0.deps_api_installed"
    echo [OK] 依赖安装完成
)

rem ── 4. 检查并启动 Chrome ──
echo [INFO] 检查 Chrome 状态...

set "CHROME="
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" set "CHROME=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not defined CHROME if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" set "CHROME=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not defined CHROME if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"

if not defined CHROME (
    echo [ERROR] 未找到 Chrome，请先安装 Google Chrome
    echo         下载地址: https://www.google.com/chrome/
    pause
    exit /b 1
)

rem 检查调试端口是否已开启
curl -s --connect-timeout 2 http://127.0.0.1:9222/json/version >nul 2>&1
if %errorlevel% == 0 (
    echo [OK] Chrome 调试端口已就绪
    goto :start_api
)

rem Chrome 未以调试模式运行，启动它
echo [INFO] 正在启动 Chrome（调试端口 9222）...
taskkill /f /t /im chrome.exe >nul 2>&1
taskkill /f /im GoogleCrashHandler.exe >nul 2>&1
taskkill /f /im GoogleCrashHandler64.exe >nul 2>&1
timeout /t 3 /nobreak >nul

del /f /q "%LOCALAPPDATA%\Google\Chrome\User Data\SingletonLock" >nul 2>&1
del /f /q "%LOCALAPPDATA%\Google\Chrome\User Data\SingletonSocket" >nul 2>&1
del /f /q "%LOCALAPPDATA%\Google\Chrome\User Data\SingletonCookie" >nul 2>&1

start "" "%CHROME%" ^
  --remote-debugging-port=9222 ^
  --user-data-dir="%LOCALAPPDATA%\Google\Chrome\User Data" ^
  --disable-blink-features=AutomationControlled ^
  --no-first-run ^
  --no-default-browser-check ^
  --no-restore-last-session
timeout /t 8 /nobreak >nul

curl -s --connect-timeout 3 http://127.0.0.1:9222/json/version >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARN] Chrome 调试端口未就绪，等待中...
    timeout /t 5 /nobreak >nul
)

:start_api
echo.
echo ============================================
echo   启动 API 服务...
echo   API 文档: http://localhost:8000/docs
echo   健康检查: http://localhost:8000/api/v1/system/health
echo ============================================
echo.
echo   提示: 首次使用前请在 Chrome 中登录淘宝/京东账号
echo   按 Ctrl+C 可停止服务
echo.

python "%~dp0run_api.py" --host 0.0.0.0 --port 8000
pause
