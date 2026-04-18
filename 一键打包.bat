@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\build_windows.ps1"
if errorlevel 1 (
  echo 打包失败。
  pause
  exit /b 1
)
echo.
pause
