@echo off
chcp 65001 >nul
set "ROOT=%~dp0"
set "EXE=%ROOT%app\dist\TidanMgr\TidanMgr.exe"
if exist "%EXE%" (
  start "" "%EXE%"
  exit /b 0
)
set "PYW=%ROOT%app\.venv\Scripts\pythonw.exe"
if exist "%PYW%" (
  pushd "%ROOT%app"
  start "" "%PYW%" "%ROOT%app\bill_app.py"
  popd
  exit /b 0
)
set "PY=%ROOT%app\.venv\Scripts\python.exe"
if exist "%PY%" (
  pushd "%ROOT%app"
  start "" "%PY%" "%ROOT%app\bill_app.py"
  popd
  exit /b 0
)
echo 未找到已打包程序：%EXE%
echo 也未找到开发环境：%ROOT%app\.venv
echo 请先双击运行「一键打包.bat」生成 exe，或在本机 app 目录执行 python -m venv .venv 并安装依赖。
pause
exit /b 1
