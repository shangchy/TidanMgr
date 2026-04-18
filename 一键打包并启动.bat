@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\build_windows.ps1"
if errorlevel 1 (
  echo 打包失败，未启动。
  pause
  exit /b 1
)
set "EXE=%~dp0app\dist\TidanMgr\TidanMgr.exe"
if exist "%EXE%" (
  start "" "%EXE%"
) else (
  echo 未找到 exe：%EXE%
  pause
  exit /b 1
)
exit /b 0
