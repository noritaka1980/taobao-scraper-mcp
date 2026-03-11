@echo off
chcp 65001 >nul 2>&1

set "CHROME="
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" set "CHROME=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not defined CHROME if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" set "CHROME=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not defined CHROME if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"

if not defined CHROME goto :start_gui

rem Check if Chrome is running with debug port by checking the process AND port together
tasklist /fi "imagename eq chrome.exe" 2>nul | findstr /i "chrome.exe" >nul 2>&1
if %errorlevel% neq 0 goto :launch_chrome

rem Chrome is running - check if debug port is actually responsive
curl -s --connect-timeout 2 http://127.0.0.1:9222/json/version >nul 2>&1
if %errorlevel% == 0 goto :start_gui

rem Chrome running but debug port not responsive - restart it
taskkill /f /im chromedriver.exe >nul 2>&1
taskkill /f /t /im chrome.exe >nul 2>&1
taskkill /f /im GoogleCrashHandler.exe >nul 2>&1
taskkill /f /im GoogleCrashHandler64.exe >nul 2>&1
timeout /t 5 /nobreak >nul

rem Clean lock files that prevent debug port from opening
del /f /q "%LOCALAPPDATA%\Google\Chrome\Debug Profile\SingletonLock" >nul 2>&1
del /f /q "%LOCALAPPDATA%\Google\Chrome\Debug Profile\SingletonSocket" >nul 2>&1
del /f /q "%LOCALAPPDATA%\Google\Chrome\Debug Profile\SingletonCookie" >nul 2>&1

:launch_chrome
rem Kill any remaining chrome helpers and stale chromedriver
taskkill /f /im chromedriver.exe >nul 2>&1
taskkill /f /im GoogleCrashHandler.exe >nul 2>&1
taskkill /f /im GoogleCrashHandler64.exe >nul 2>&1
del /f /q "%LOCALAPPDATA%\Google\Chrome\Debug Profile\SingletonLock" >nul 2>&1
del /f /q "%LOCALAPPDATA%\Google\Chrome\Debug Profile\SingletonSocket" >nul 2>&1
del /f /q "%LOCALAPPDATA%\Google\Chrome\Debug Profile\SingletonCookie" >nul 2>&1
timeout /t 2 /nobreak >nul
start "" "%CHROME%" ^
  --remote-debugging-port=9222 ^
  --user-data-dir="%LOCALAPPDATA%\Google\Chrome\Debug Profile" ^
  --disable-blink-features=AutomationControlled ^
  --no-first-run ^
  --no-default-browser-check ^
  --no-restore-last-session
timeout /t 8 /nobreak >nul

rem Verify debug port is actually working
curl -s --connect-timeout 3 http://127.0.0.1:9222/json/version >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARN] Debug port 9222 not ready, waiting 5 more seconds...
    timeout /t 5 /nobreak >nul
)

:start_gui
if not exist "%~dp0.deps_gui_installed" (
    pip install selenium pandas openpyxl webdriver-manager pillow pywin32 -q --disable-pip-version-check >nul 2>&1
    echo. > "%~dp0.deps_gui_installed"
)

wscript //nologo "%~dp0launch.vbs"
