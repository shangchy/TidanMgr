#!/usr/bin/env bash
set -euo pipefail

# 无参数静默更新脚本（可双击执行）：
# - 默认读取脚本同目录下的 license.new.json（新授权）
# - 写入脚本同目录（应与 TidanMgr.app 同级）的 license.json
#
# 失败时仅返回非 0，不弹交互提示。

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DST_FILE="$SCRIPT_DIR/license.json"
SRC="$SCRIPT_DIR/license.new.json"
if [[ ! -f "$SRC" ]]; then
  exit 1
fi

cp -f "$SRC" "$DST_FILE" >/dev/null 2>&1
