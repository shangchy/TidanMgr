"""应用数据目录、模板路径与授权校验（与 UI 分离，便于打包路径一致）。"""
import os
import sys
from datetime import datetime
from pathlib import Path


def _app_dir() -> Path:
    """可写数据目录：开发为脚本目录；Windows 打包为 exe 同级；macOS 打包为应用支持库（或便携目录）。"""
    if not getattr(sys, "frozen", False):
        return Path(__file__).resolve().parent
    if sys.platform == "darwin":
        portable = os.environ.get("TIDANMGR_PORTABLE", "").strip().lower() in ("1", "true", "yes")
        if portable:
            macos = Path(sys.executable).resolve().parent
            bundle = macos.parent.parent
            data = bundle.parent / "TidanMgrData"
            data.mkdir(parents=True, exist_ok=True)
            return data
        support = Path.home() / "Library" / "Application Support" / "TidanMgr"
        support.mkdir(parents=True, exist_ok=True)
        return support
    return Path(sys.executable).resolve().parent


APP_DIR = _app_dir()
DATA_FILE = APP_DIR / "data.json"
HISTORY_FILE = APP_DIR / "history_data.json"
THEME_FILE = APP_DIR / "theme.json"
PICKER_RECENT_FILE = APP_DIR / "picker_recent.json"
PRINT_RECORDS_FILE = APP_DIR / "print_records.json"
# 打包时模板随 --add-data 放入 _MEIPASS；未打包时放在 app 目录
if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    _bundled_tpl = Path(sys._MEIPASS) / "template.xlsx"
    TEMPLATE_FILE = _bundled_tpl if _bundled_tpl.exists() else (APP_DIR / "template.xlsx")
else:
    TEMPLATE_FILE = APP_DIR / "template.xlsx"

# 本地时间：此时间之后授权失效（2026-04-30 23:00:00 及之前可用）
_LICENSE_EXPIRE_AT = datetime(2026, 4, 30, 23, 0, 0)
LICENSE_EXPIRED_MSG = "授权已到期，请联系管理员。\n程序将自动退出。"


def is_license_valid() -> bool:
    return datetime.now() <= _LICENSE_EXPIRE_AT
