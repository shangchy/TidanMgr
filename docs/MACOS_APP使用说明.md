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
| **`TidanMgr-macos-portable.zip`** | 对上述 **三项** 的压缩包，**文件名仅 ASCII**，可直接发给试用用户。 |

**标准使用**：将 **`TidanMgr.app`** 拖入「应用程序」，数据默认在  
`~/Library/Application Support/TidanMgr/`。

**便携使用**：解压 zip 后，双击 **`PortableStart.command`**（勿用 `open` 打开 `.app` 来走便携逻辑；请按脚本说明执行）。

---

## 三、用户侧：收到压缩包后怎么做

1. 解压 **`TidanMgr-macos-portable.zip`**，应看到 **`TidanMgr.app`**、**`PortableStart.command`**、**`run_TidanMgr.sh`**。
2. **常用**：把 **`TidanMgr.app`** 拖到「应用程序」，从启动台或访达打开。
3. **便携（数据跟 U 盘走）**：在同一文件夹内双击 **`PortableStart.command`**；数据目录为 **`TidanMgrData/`**（与说明中 bill_app 逻辑一致）。

### 系统与权限

- 建议 **macOS 11** 或更高（与 PySide6 当前环境一致即可）。
- 未做 Apple 公证时，首次可能提示「无法打开」：**系统设置 → 隐私与安全性** 中允许；或对 **`.app` 右键 → 打开**。

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

推送至默认分支等触发 **`.github/workflows/build-macos-app.yml`** 时，会在 **macOS Runner** 上执行 **`app/build_macos.sh`**，并上传制品 **`TidanMgr-macos-portable.zip`**（Artifact 名称：`TidanMgr-macos-portable`）。

---

## 七、常见问题

**Q：双击 `.app` 闪退？**  
A：用终端执行 `bash app/dist/run_TidanMgr.sh` 查看报错；或检查是否被安全软件隔离。

**Q：只想发一个文件给用户？**  
A：发送 **`TidanMgr-macos-portable.zip`** 即可；用户解压后需保留 **`.app` 与 `PortableStart.command` / `run_TidanMgr.sh` 的相对位置**（zip 已按此结构打包）。

**Q：与 Windows 包区别？**  
A：平台不同，不可混用；Windows 说明见 **`docs/Windows部署试用说明.md`**。

---

更多打包要点见 **`app/PACKAGING_MACOS.txt`**。

---

## 八、与 Windows 包对应关系

| 平台 | 分发用压缩包（ASCII 文件名） |
|------|------------------------------|
| Windows | `app/dist/TidanMgr-Windows-x64-portable.zip` |
| macOS | `app/dist/TidanMgr-macos-portable.zip` |

Windows 部署说明见 **`docs/Windows部署试用说明.md`**。
