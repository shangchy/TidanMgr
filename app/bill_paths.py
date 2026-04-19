"""应用数据目录、模板路径与授权校验（与 UI 分离，便于打包路径一致）。"""
import base64
import hashlib
import hmac
import json
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


def _license_dir() -> Path:
    """授权文件目录：与可执行程序同路径（Win: exe 同级；macOS: .app 同级目录）。"""
    if not getattr(sys, "frozen", False):
        return Path(__file__).resolve().parent
    exe = Path(sys.executable).resolve()
    if sys.platform == "darwin":
        # .../<Name>.app/Contents/MacOS/<exe> -> 取 <Name>.app 的上级目录
        app_bundle = exe.parent.parent.parent
        return app_bundle.parent
    return exe.parent


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

# 授权文件（独立于部署包）：到期后只需替换该文件即可续期
# 加密结构示例：{"v":1,"n":"...","c":"...","s":"..."}（明文不落盘）
LICENSE_FILE = _license_dir() / "license.json"
LICENSE_EXPIRED_MSG = "授权已到期，请联系管理员。"

# 轻量对称加密 + HMAC 完整性校验（避免明文授权和简单篡改）
_LICENSE_ENC_KEY = b"TidanMgr-Lic-EncKey-v1-ChangeMe"
_LICENSE_SIG_KEY = b"TidanMgr-Lic-SigKey-v1-ChangeMe"


def _license_keystream(nonce: bytes, length: int) -> bytes:
    out = bytearray()
    counter = 0
    while len(out) < length:
        block = hashlib.sha256(_LICENSE_ENC_KEY + nonce + counter.to_bytes(4, "big")).digest()
        out.extend(block)
        counter += 1
    return bytes(out[:length])


def _license_decrypt(ciphertext: bytes, nonce: bytes) -> bytes:
    ks = _license_keystream(nonce, len(ciphertext))
    return bytes(c ^ k for c, k in zip(ciphertext, ks))


def _license_encrypt(plaintext: bytes, nonce: bytes) -> bytes:
    ks = _license_keystream(nonce, len(plaintext))
    return bytes(p ^ k for p, k in zip(plaintext, ks))


def build_encrypted_license(expire_at: str) -> dict[str, str | int]:
    """构造加密授权对象，expire_at 支持 YYYY-MM-DD / YYYY-MM-DD HH:MM:SS。"""
    s = str(expire_at or "").strip()
    if not s:
        raise ValueError("expire_at is required")
    if len(s) <= 10:
        datetime.strptime(s, "%Y-%m-%d")
    else:
        datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
    payload = json.dumps({"expire_at": s}, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    nonce = os.urandom(16)
    ciphertext = _license_encrypt(payload, nonce)
    n = base64.urlsafe_b64encode(nonce).decode("utf-8")
    c = base64.urlsafe_b64encode(ciphertext).decode("utf-8")
    msg = f"{n}.{c}".encode("utf-8")
    sig = hmac.new(_LICENSE_SIG_KEY, msg, hashlib.sha256).hexdigest()
    return {"v": 1, "n": n, "c": c, "s": sig}


def is_license_valid() -> bool:
    try:
        if not LICENSE_FILE.exists():
            return False
        obj = json.loads(LICENSE_FILE.read_text(encoding="utf-8"))
        if int(obj.get("v", 0)) != 1:
            return False
        n = str(obj.get("n", "")).strip()
        c = str(obj.get("c", "")).strip()
        sgn = str(obj.get("s", "")).strip()
        if not n or not c or not sgn:
            return False
        msg = f"{n}.{c}".encode("utf-8")
        expected = hmac.new(_LICENSE_SIG_KEY, msg, hashlib.sha256).hexdigest()
        if not hmac.compare_digest(expected, sgn):
            return False
        nonce = base64.urlsafe_b64decode(n.encode("utf-8"))
        ciphertext = base64.urlsafe_b64decode(c.encode("utf-8"))
        plain = _license_decrypt(ciphertext, nonce).decode("utf-8")
        payload = json.loads(plain)
        s = str(payload.get("expire_at", "")).strip()
        if not s:
            return False
        # 支持两种格式：YYYY-MM-DD 或 YYYY-MM-DD HH:MM:SS
        if len(s) <= 10:
            expire_at = datetime.strptime(s, "%Y-%m-%d")
        else:
            expire_at = datetime.strptime(s, "%Y-%m-%d %H:%M:%S")
        return datetime.now() <= expire_at
    except Exception:
        # 授权文件损坏/格式错误视为未授权
        return False
