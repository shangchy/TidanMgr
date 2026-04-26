param(
  [string]$AppName = "TidanMgr"
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
  Write-Error "Python 3.10+ not found in PATH"
}

if (-not (Test-Path ".venv")) {
  python -m venv .venv
}

& ".\.venv\Scripts\Activate.ps1"
python -m pip install --upgrade pip
pip install -r requirements.txt

$addData = @()
if (Test-Path "template.xlsx") {
  $addData += @("--add-data", "template.xlsx;.")
  Write-Host "Bundled: template.xlsx"
} else {
  Write-Host "Warning: template.xlsx missing; export may need template in app folder"
}
if (Test-Path "sum-template.xlsx") {
  $addData += @("--add-data", "sum-template.xlsx;.")
  Write-Host "Bundled: sum-template.xlsx"
} else {
  Write-Host "Warning: sum-template.xlsx missing; merge stats sheet may lose template style"
}

Write-Host "Running PyInstaller (onedir)..."
$pyiArgs = @(
  "--noconfirm",
  "--windowed",
  "--clean",
  "--name", $AppName,
  "--noupx",
  "--hidden-import", "bill_theme"
) + $addData + @("bill_app.py")

pyinstaller @pyiArgs

$distDir = Join-Path $PSScriptRoot "dist\$AppName"
$launcher = Join-Path $distDir "Start.bat"
# Avoid here-string line starting with @ (PowerShell parse issue)
$bat = (@(
  '@echo off',
  'chcp 65001 >nul',
  'cd /d "%~dp0"',
  ('start "" "%~dp0{0}.exe"' -f $AppName)
) -join "`r`n")
Set-Content -LiteralPath $launcher -Value $bat -Encoding utf8

Write-Host ""
Write-Host "Done. Run:"
Write-Host "  1) $distDir\$AppName.exe"
Write-Host "  2) $launcher"
Write-Host "  3) One-click bat at repo root (dev only)"

$zipPath = Join-Path $PSScriptRoot "dist\TidanMgr-Windows-x64-portable.zip"
if (Test-Path $zipPath) {
  Remove-Item -LiteralPath $zipPath -Force
}
Compress-Archive -LiteralPath $distDir -DestinationPath $zipPath -Force
Write-Host ""
Write-Host "Zip for distribution: $zipPath"
