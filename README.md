# TidanMgr（提单管理）

单机桌面应用，用于在本机维护提单列表、按规则编辑**任务名**与业务字段、筛选查询，并基于 Excel **模板**导出勾选记录。支持 **Windows** 与 **macOS**，数据以 JSON 形式保存在本机。

**仓库**：<https://github.com/shangchy/TidanMgr>

---

## 功能概览

| 能力 | 说明 |
|------|------|
| **提单主表** | 表格内联编辑；任务名、数量、时长、年龄、pv、行业编码等；运营商/类型下拉；首列勾选与表头全选 |
| **任务名规则** | 与类型、运营商、行业编码、省份/地市等字段**双向联动**；支持按「地域词 + 6 位码」解析并校验 |
| **地域** | 省/市/排除省/排除地市；多选弹窗；支持手输与列表校验；选省后「排除地市」仅展示该省下辖地市（数据范围内） |
| **URL** | 多行多地址；列表摘要显示首条；导出可选是否带 URL |
| **提单时间** | 字段展示为提单时间；从历史恢复时刷新为当前时间 |
| **筛选与搜索** | 任务名模糊搜索；表头多选筛选；排序指示 |
| **历史提单** | 删除/归档进入历史；支持勾选恢复到主表 |
| **主题** | 浅色 / 深色，可持久化 |
| **导出** | 勾选行按 `app/template.xlsx` 写入（openpyxl），文件名含客户名与日期 |
| **打包运行** | PyInstaller 生成 Windows 文件夹版 / macOS `.app`（详见 `docs/`） |

更细的规则与字段说明见 **`docs/需求规格说明书.md`**、**`docs/详细设计文档.md`**。

---

## 技术栈

| 类别 | 选型 |
|------|------|
| 语言 | **Python 3.10+** |
| GUI | **PySide6**（Qt for Python） |
| 表格与 Excel | **openpyxl** 读写模板 |
| 打包 | **PyInstaller**（`*.spec` + 脚本） |
| 数据 | **JSON**（主表、历史、主题、部分选择器记忆） |

核心界面与业务逻辑集中在 **`app/bill_app.py`**，主题样式在 **`app/bill_theme.py`**。

---

## 仓库结构（摘要）

```text
app/
  bill_app.py          # 主程序入口与界面逻辑
  bill_theme.py        # 深浅色 QSS
  requirements.txt     # 运行时与打包依赖
  template.xlsx        # 导出模板（勿随意改列结构）
  TidanMgr.spec        # Windows 等打包配置示例
  build_windows.ps1    # Windows 打包脚本
  build_macos.sh       # macOS 打包脚本
docs/                  # 需求、设计、Windows/macOS 部署说明
一键启动.bat            # 开发机快速启动（见脚本内说明）
```

运行时会在可写目录生成 `data.json`、`history_data.json`、`theme.json` 等（开发模式多为 `app/` 或 exe 同级，详见代码内 `_app_dir()` 与部署文档）。

---

## 开发环境运行

1. 安装 **Python 3.10+**，进入 `app` 目录。
2. 建议使用虚拟环境：

   ```bash
   python -m venv .venv
   .venv\Scripts\activate          # Windows
   # source .venv/bin/activate   # macOS / Linux
   ```

3. 安装依赖：

   ```bash
   pip install -r requirements.txt
   ```

4. 启动（任选）：

   ```bash
   python bill_app.py
   ```

   Windows 无控制台调试可用 `pythonw bill_app.py`。

根目录 **`一键启动.bat`** 会尝试使用 `app\.venv` 并启动程序（需已按上文创建 venv 并安装依赖）。

---

## 打包与试用分发

- **Windows**：`docs/Windows部署试用说明.md`，以及 `app/build_windows.ps1`。
- **macOS**：`docs/MACOS_APP使用说明.md`，以及 `app/build_macos.sh`、`.github/workflows/build-macos-app.yml`。

打包产物体积较大，**不要**将 `app/build/`、`app/dist/`、`app/.venv/` 提交到 Git（本仓库 `.gitignore` 已排除）。

---

## 文档索引

| 文档 | 内容 |
|------|------|
| `docs/README.md` | 文档目录说明 |
| `docs/需求规格说明书.md` | 需求与范围 |
| `docs/详细设计文档.md` | 架构、数据模型、校验与导出设计 |
| `docs/Windows部署试用说明.md` | Windows 试用包说明 |
| `docs/MACOS_APP使用说明.md` | macOS 构建与使用 |
| `app/模板放置说明.txt` | 模板文件放置约定 |

---

## 许可证

见仓库根目录 **`LICENSE`**（MIT）。
