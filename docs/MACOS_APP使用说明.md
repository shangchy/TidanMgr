# 提单管理（TidanMgr）macOS 部署与试用说明

本文说明如何在 **Mac** 上打包、向用户分发，以及用户侧的安装与使用。**必须在 macOS 上执行打包**（Windows 无法直接生成 `.app`）。

---

## 一、开发人员：一键打包

### 方式 A：仓库根目录

双击 **`一键打包_macOS.command`**（若无法执行，在终端执行：`chmod +x 一键打包_macOS.command`）。

### 方式 B：终端

```bash
cd app
chmod +x build_macos.sh
./build_macos.sh
```

依赖：已安装 **Python 3.10+**、`python3` 在 PATH 中；脚本会创建/使用 **`app/.venv`** 并安装 `requirements.txt`。

---

## 二、打包产物（均在 `app/dist/`）

| 文件 / 目录 | 说明 |
|-------------|------|
| **`TidanMgr.app`** | 主程序；窗口标题仍为「提单管理」。 |
| **`PortableStart.command`** | **便携模式**：设置 `TIDANMGR_PORTABLE=1`，数据写在同级的 **`TidanMgrData/`**（可与 `.app` 一起拷贝、U 盘使用）。文件名无中文。 |
| **`run_TidanMgr.sh`** | 在终端启动（便携数据），便于查看日志。 |
| **`TidanMgr-macos-portable.zip`** | 本地 `build_macos.sh` 默认打出的便携压缩包（**文件名仅 ASCII**）。 |
| **`TidanMgr-macos-arm64-portable.zip`** | CI（Apple Silicon 跑机）产物，适用于 **M 系列芯片** Mac。 |
| **`TidanMgr-macos-intel-portable.zip`** | CI（Intel 跑机）产物，适用于 **Intel 芯片** Mac。 |

**标准使用**：将 **`TidanMgr.app`** 拖入「应用程序」，数据默认在  
`~/Library/Application Support/TidanMgr/`。

**便携使用**：解压 zip 后，双击 **`PortableStart.command`**（勿用 `open` 打开 `.app` 来走便携逻辑；请按脚本说明执行）。

---

## 三、用户侧：收到压缩包后怎么做

1. 解压收到的便携压缩包（CI：**`TidanMgr-macos-arm64-portable.zip`** 或 **`TidanMgr-macos-intel-portable.zip`**；本地打包：**`TidanMgr-macos-portable.zip`**），应看到 **`TidanMgr.app`**、**`PortableStart.command`**、**`run_TidanMgr.sh`**。
2. **常用**：把 **`TidanMgr.app`** 拖到「应用程序」，从启动台或访达打开。
3. **便携（数据跟 U 盘走）**：在同一文件夹内双击 **`PortableStart.command`**；数据目录为 **`TidanMgrData/`**（与说明中 bill_app 逻辑一致）。

### 系统与权限

- **最低系统**：本仓库 **`app/requirements.txt`** 将 **PySide6 限制在 6.6.x**（`<6.7`），以便 **macOS 12（Monterey）** 仍可运行；若自行去掉版本上限并安装 **PySide6 6.7+**，则终端可能报错：**「Qt requires macOS 13.0.0 or later」**，此时需 **升级到 macOS 13（Ventura）或更高**，或恢复依赖中的版本上限后重新打包。
- 未做 Apple 公证时，首次可能提示「无法打开」：**系统设置 → 隐私与安全性** 中允许；或对 **`.app` 右键 → 打开**。经 **微信等 IM 传输** 的文件若仍异常，可对 **`.app` 或 zip 所在目录** 在终端执行：`xattr -dr com.apple.quarantine "/path/to/TidanMgr.app"` 后再打开。

---

## 四、数据与模板

| 场景 | 数据位置 |
|------|----------|
| 标准（直接运行 `.app`） | `~/Library/Application Support/TidanMgr/`（`data.json`、`history_data.json`、`theme.json` 等） |
| 便携（`PortableStart.command`） | 与 `dist` 同级的 **`TidanMgrData/`**（由程序创建） |

**导出模板**：打包前将 **`template.xlsx`** 放在 **`app/`** 目录，会打入应用包；运行时亦可按程序说明在可写目录放置模板。

---

## 五、试用授权（与 Windows 版一致）

程序内含 **授权截止时间**（见 `app/bill_app.py` 中 `_LICENSE_EXPIRE_AT`）。过期后启动会提示 **「授权已过期，请联系管理员。」** 并退出；运行中过期会定时检测并退出。

续期需由开发人员更新该日期后 **重新打包**。

---

## 六、CI 自动构建

推送至默认分支等触发 **`.github/workflows/build-macos-app.yml`** 时，会在 **两种架构** 的 macOS Runner 上各执行一次 **`app/build_macos.sh`**，并上传制品（请按用户 Mac 的 CPU 选用对应压缩包）：

| Artifact 名称 | 压缩包内文件名 | 适用 Mac |
|----------------|----------------|----------|
| `TidanMgr-macos-arm64` | `TidanMgr-macos-arm64-portable.zip` | Apple Silicon（M1 / M2 / M3…，`uname -m` 为 `arm64`） |
| `TidanMgr-macos-intel` | `TidanMgr-macos-intel-portable.zip` | Intel（`uname -m` 为 `x86_64`） |

此前仅使用 `macos-latest`（arm64）打单一 zip 时，在 **Intel Mac** 上双击 `.app` 易出现系统提示：**「您无法打开应用程序…因为这台 Mac 不支持此应用程序」**——实为 **CPU 架构不匹配**，并非 zip 损坏。请改用 **Intel** 对应制品，或在目标 Mac 上本地执行 **`./build_macos.sh`** 自行打包。

---

## 七、常见问题

**Q：系统提示「您无法打开应用程序…因为这台 Mac 不支持此应用程序」？**  
A：多数是 **应用二进制架构与当前 Mac 不一致**（例如在 Intel Mac 上运行了仅在 Apple Silicon 上构建的 `.app`）。

- 在本机终端执行 **`uname -m`**：`arm64` 为 Apple Silicon，`x86_64` 为 Intel。
- 对已收到的 `.app` 可执行（路径按你放置位置调整）：
  ```bash
  file "/path/to/TidanMgr.app/Contents/MacOS/TidanMgr"
  ```
  输出中含 **`arm64`** 则需在 Apple Silicon 上运行（Intel Mac 需 Rosetta 无法“转译”这种原生 GUI 单架构包，应换用 **Intel 版** 或在 Intel Mac 上重新打包）；含 **`x86_64`** 则适用于 Intel（Apple Silicon 上通常可通过 Rosetta 打开，具体以系统为准）。
- **处理**：从 CI 下载与 `uname -m` 一致的 **`TidanMgr-macos-arm64-portable.zip`** 或 **`TidanMgr-macos-intel-portable.zip`**；或在当前这台 Mac 上进入 `app` 执行 **`bash build_macos.sh`** 生成本地匹配的 `TidanMgr.app`。

**Q：终端或 `PortableStart.command` 提示「Qt requires macOS 13.0.0 or later, you have macOS 12.x」？**  
A：当前运行的 **`.app` 是用 PySide6 6.7+ 打的包**，其自带 Qt 二进制 **最低只支持 macOS 13**。处理方式二选一：**(1)** 将系统升级到 **macOS 13+**；**(2)** 使用本仓库已限制 **`PySide6<6.7`** 的 **`requirements.txt`** 后 **重新执行 `build_macos.sh` / CI 再打 zip**，再在 macOS 12 上安装运行。

**Q：双击 `.app` 闪退？**  
A：用终端执行 `bash app/dist/run_TidanMgr.sh` 查看报错；或检查是否被安全软件隔离。

**Q：只想发一个文件给用户？**  
A：需确认对方 Mac 架构后发送对应 zip：**`TidanMgr-macos-arm64-portable.zip`** 或 **`TidanMgr-macos-intel-portable.zip`**（CI）；本地默认 **`TidanMgr-macos-portable.zip`** 与当前打包机架构一致。用户解压后需保留 **`.app` 与 `PortableStart.command` / `run_TidanMgr.sh` 的相对位置**。

**Q：与 Windows 包区别？**  
A：平台不同，不可混用；Windows 说明见 **`docs/Windows部署试用说明.md`**。

---

更多打包要点见 **`app/PACKAGING_MACOS.txt`**。

---

## 八、与 Windows 包对应关系

| 平台 | 分发用压缩包（ASCII 文件名） |
|------|------------------------------|
| Windows | `app/dist/TidanMgr-Windows-x64-portable.zip` |
| macOS（Apple Silicon） | `app/dist/TidanMgr-macos-arm64-portable.zip`（CI） |
| macOS（Intel） | `app/dist/TidanMgr-macos-intel-portable.zip`（CI） |
| macOS（本机打包默认名） | `app/dist/TidanMgr-macos-portable.zip` |

Windows 部署说明见 **`docs/Windows部署试用说明.md`**。
