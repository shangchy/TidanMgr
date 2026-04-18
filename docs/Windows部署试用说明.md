# 提单管理（TidanMgr）Windows 试用部署说明

本文说明如何将 **Windows 可执行包** 分发给试用用户，以及用户侧的安装与使用方式。

---

## 一、您要发给用户的文件

打包完成后，在 **`app/dist/`** 目录下会得到：

| 内容 | 说明 |
|------|------|
| **`TidanMgr/`** 整个文件夹 | 内含 `TidanMgr.exe`、依赖库、`Start.bat` 等，**必须整夹一起分发**，不能只拷贝单个 exe。 |
| **`TidanMgr-Windows-x64-portable.zip`**（打包脚本生成） | 对上述文件夹的压缩包，文件名无中文；用户解压后得到 `TidanMgr` 文件夹即可。 |

**请勿**只发送 `TidanMgr.exe`：程序依赖同目录下的多个 `.dll` 与 `_internal` 等文件，缺少会导致无法启动。

---

## 二、用户侧操作步骤

1. **解压**（若收到的是 zip）到任意路径，例如 `D:\Tools\TidanMgr\`。
2. 进入 **`TidanMgr`** 文件夹。
3. 双击 **`TidanMgr.exe`** 启动；也可双击 **`Start.bat`**（已设置 UTF-8 控制台编码，效果与直接运行 exe 相同）。
4. 窗口标题为 **「提单管理」**，界面为提单列表、历史提单、设置等。

### 系统要求

- **Windows 10 64 位** 或更高版本（与当前打包环境一致即可）。
- 首次运行若出现 **Windows 已保护你的电脑 / SmartScreen**，请选择 **「仍要运行」**（未做代码签名时常见）。
- 个别杀毒软件可能误报 PyInstaller 打包程序，可将整个 **`TidanMgr`** 文件夹加入信任区。

---

## 三、数据与模板存放位置（试用版）

程序在 **打包模式** 下，数据文件写在 **`TidanMgr.exe` 所在目录**（与 exe 同级），主要包括：

| 文件 | 用途 |
|------|------|
| `data.json` | 当前提单数据 |
| `history_data.json` | 历史提单数据 |
| `theme.json` | 主题（浅色/深色） |
| `picker_recent.json` | 部分选择器最近使用记录 |

**导出 Excel** 使用的模板在打包时已内置（若打包时 `app` 目录下存在 **`template.xlsx`**）。若您分发的包中导出异常，可将 `template.xlsx` 放在 **与 `TidanMgr.exe` 同一目录** 后重试。

**备份建议**：试用或迁移时，复制整个 `TidanMgr` 文件夹即可带走数据与配置。

---

## 四、如何重新打包（开发人员）

在已安装 **Python 3.10+** 的 Windows 电脑上：

1. 打开 PowerShell，进入 **`app`** 目录：  
   `cd <项目路径>\app`
2. 确保 **`template.xlsx`** 位于 `app` 目录（可选，但建议保留以便导出）。
3. 执行：  
   `powershell -ExecutionPolicy Bypass -File .\build_windows.ps1`
4. 产物在 **`app\dist\TidanMgr\`**；脚本会同时生成 **`app\dist\TidanMgr-Windows-x64-portable.zip`**（纯英文文件名），可直接发给用户。

---

## 五、常见问题

**Q：双击 exe 没反应？**  
A：检查是否在同一文件夹内保留了 PyInstaller 生成的全部文件；可在该目录下用命令行运行 `TidanMgr.exe` 查看是否有报错信息。

**Q：杀毒软件删除文件？**  
A：恢复被删文件或将文件夹加入白名单后重新解压/复制一份完整目录。

**Q：与开发版「一键启动.bat」区别？**  
A：一键启动用于本仓库开发路径运行源码；**试用用户只需使用 `dist\TidanMgr` 内的 exe 或 `Start.bat`**。

---

## 六、macOS 用户说明

macOS 安装包与分发 zip 的说明见 **`docs/MACOS_APP使用说明.md`**（制品名 **`TidanMgr-macos-portable.zip`**，需在 Mac 上打包）。

---

## 七、反馈与版本

试用阶段请将 **系统版本、复现步骤、截图或报错原文** 一并反馈给提供方，便于排查。

（文档随仓库更新；应用窗口内功能以实际界面为准。）
