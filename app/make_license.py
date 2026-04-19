"""生成加密授权文件（license.json）。

用法：
  python make_license.py --expire "2026-12-31 23:59:59"
  python make_license.py --expire "2026-12-31" --out "D:/path/license.json"
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from bill_paths import LICENSE_FILE, build_encrypted_license


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate encrypted license.json")
    parser.add_argument(
        "--expire",
        required=True,
        help="Expire datetime, format: YYYY-MM-DD or YYYY-MM-DD HH:MM:SS",
    )
    parser.add_argument(
        "--out",
        default=str(LICENSE_FILE),
        help="Output path of license.json (default: APP_DIR/license.json)",
    )
    args = parser.parse_args()

    out = Path(args.out).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    lic = build_encrypted_license(args.expire)
    out.write_text(json.dumps(lic, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"license generated: {out}")
    print(f"expire_at: {args.expire}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
