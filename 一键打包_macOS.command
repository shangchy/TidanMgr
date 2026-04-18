#!/bin/bash
# 在 macOS 上双击本文件（若提示未授权，可在终端执行：chmod +x 一键打包_macOS.command）
cd "$(dirname "$0")" || exit 1
exec bash "./app/build_macos.sh"
