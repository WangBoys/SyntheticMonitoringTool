# 网址拨测可视化工具

一个基于 PySide6 + Playwright 的 Windows 桌面工具，用于批量读取 Excel 中的网址并执行可访问性拨测，最终回写结果到新的 Excel 文件。

## 1. 功能概览

- Excel 文件输入支持手动填写、拖拽、弹窗选择。
- 自动读取工作簿结构，支持选择 Sheet、网址列、首行是否表头。
- 支持加密 Excel（优先 `msoffcrypto`，可选 Office/WPS COM 兜底）。
- 并发拨测（可运行中动态调整并发）、超时、失败重试。
- 成功规则支持：
  - 仅 `2xx/3xx`
  - `2xx/3xx + 无状态码`
  - 任意 `HTTP` 状态码（`2xx~5xx`）
  - 自定义规则（如 `200,204,3xx,5xx`）
- 结果回写：
  - 原 Sheet 末尾新增 `是否可访问`、`访问详情`
  - 自动重建 `拨测统计` Sheet（总数/成功数/失败数/成功率/生成时间）
- UI 支持开始/暂停/继续/中止，实时展示进度、失败数、成功率、线程日志。

## 2. 快速开始

### 2.1 环境准备

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

> Windows 下若需使用 Office/WPS 解密兜底，请确保本机已安装可被 COM 调用的 Office 或 WPS。

### 2.2 启动

```bash
python run_tool.py
```

### 2.3 输出规则

- 建议输入 `.xlsx` 文件。
- 默认输出路径为：`原文件名_monitoring_result.xlsx`。
- 工具不会覆盖原始输入文件，结果写入新文件。

## 3. Ubuntu / CLI 模式（无 GUI）

适用于服务器或 cron 定时任务，通过 JSON 配置文件驱动拨测，无需安装 PySide6。

### 3.1 一键部署

```bash
chmod +x setup_ubuntu.sh
./setup_ubuntu.sh
```

脚本会创建 `.venv`、安装 `requirements-linux.txt`、同步 Playwright Chromium，并从 `config.example.json` 生成 `config.json`。

### 3.2 配置文件

复制并编辑 [`config.example.json`](config.example.json)：

| 字段 | 说明 |
|------|------|
| `input.file_path` | 输入 Excel 路径 |
| `input.password` | 加密文件密码（明文，可选） |
| `input.password_env` | 从环境变量读取密码（优先于 `password`） |
| `input.use_office_fallback` | Office/WPS 兜底（**仅 Windows**） |
| `probe.sheet_name` | 工作表名称 |
| `probe.has_header` | 首行是否为表头 |
| `probe.url_column` | 列号 / 列字母（如 `B`）/ 表头名 |
| `probe.concurrency` | 并发数（1~100） |
| `probe.timeout_seconds` | 超时秒数（1~300） |
| `probe.retry_count` | 失败重试（0~5） |
| `probe.success_mode` | `2xx_3xx` / `2xx_3xx_or_no_status` / `any_http` / `custom` |
| `probe.custom_status_codes` | `success_mode=custom` 时必填 |
| `output.path` | 输出路径（省略则自动生成） |

### 3.3 运行

```bash
./run_cli.sh --dry-run                        # 仅校验（默认读取当前目录 config.json）
./run_cli.sh                                  # 执行拨测
./run_cli.sh --config /path/to/other.json     # 指定其他配置文件（优先于默认）
```

`run_cli.sh` 会自动使用项目内 `.venv` 的 Python，**无需**手动 `source .venv/bin/activate`。未传 `--config` 时使用**当前工作目录**下的 `config.json`。

若需进入 venv 交互调试，仍可执行：`source .venv/bin/activate`

退出码：`0` 成功 / `1` 错误 / `130` 用户中断（SIGINT/SIGTERM）。

### 3.4 cron 示例

```cron
0 2 * * * cd /opt/SyntheticMonitoringTool && ./run_cli.sh >> /var/log/probe.log 2>&1
```

> Linux 下加密 Excel 仅支持 `msoffcrypto` 解密，不支持 Office/WPS COM 兜底。

### 3.5 从 Windows 打包脚本到 Ubuntu 部署（推荐）

在 **Windows** 开发机上仅打包脚本与 JSON 配置（不含依赖与浏览器）：

```bat
pack_ubuntu_scripts.bat
```

产物：

- `dist\SyntheticMonitoringTool-scripts\` — 脚本目录
- `dist\SyntheticMonitoringTool-scripts.zip` — 上传到服务器

**Ubuntu 服务器（首次）：**

```bash
cd /opt
unzip SyntheticMonitoringTool-scripts.zip -d SyntheticMonitoringTool
cd SyntheticMonitoringTool
chmod +x setup_ubuntu.sh run_cli.sh
./setup_ubuntu.sh          # 安装 Python 依赖、Playwright 浏览器
cp config.example.json config.json && vim config.json
./run_cli.sh --dry-run
./run_cli.sh
```

日常拨测只需 `./run_cli.sh`。CLI 启动时会输出公网 IP、运营商、国家/省份/城市（与 Windows GUI 一致），随后输出拨测进度。

详见包内 `DEPLOY-ubuntu.md`。

### 3.6 可选：Ubuntu 本机打完整离线包

若希望在 Ubuntu 构建机上一次性打入 `vendor/` 与 `pw-browsers/`（服务器无需 pip），可使用：

```bash
./build_ubuntu_package.sh
```

详见包内 `DEPLOY.md`。

## 4. 打包与发布（Windows）

### 4.0 打包 Ubuntu CLI 脚本包

```bat
pack_ubuntu_scripts.bat
```

仅复制 `.py` / `.sh` / `.json` / `requirements-linux.txt` 等脚本与配置，供 Ubuntu 上执行 `setup_ubuntu.sh` 安装依赖。详见上文 **3.5**。

### 4.1 一键构建 EXE + 安装器

```bat
build_exe.bat
```

默认成功后产物：

- `dist\SyntheticMonitoringTool\SyntheticMonitoringTool.exe`
- `installer\SyntheticMonitoringTool-Setup.exe`

## 4.2 增量构建参数（与 `build_exe.bat` 保持一致）

- `--skip-installer`：仅构建 EXE，不生成安装器。
- `--full-clean`：全量清理构建（会启用 PyInstaller `--clean` 并清理 work 子目录）。
- `--force-deps`：强制重新安装 `requirements.txt` 依赖。
- `--force-browser`：强制重新安装 Playwright Chromium 到本地 `pw-browsers`。
- `--add-defender-exclusions`：尝试自动添加 Windows Defender 排除项（通常需管理员权限）。

示例：

```bat
build_exe.bat --skip-installer --full-clean
```

## 4.3 仅做安装器预检并调用完整构建

```bat
build_installer.bat
```

说明：`build_installer.bat` 会先做环境预检（Python、依赖、PyInstaller、ISCC、关键文件），通过后调用 `build_exe.bat` 执行完整构建流程。

## 4.4 打包目录约定

- `pw-browsers`：本地 Playwright 浏览器缓存/分发目录。
- `build`：PyInstaller 工作目录（中间产物）。
- `dist`：EXE 分发目录。
- `installer`：Inno Setup 安装器输出目录。

## 5. 运行时目录约定

通过 `run_tool.py` 启动时，程序会在应用根目录下创建并使用：

- `runtime-data\temp`：临时文件目录（同时设置到 `TMP/TEMP/TMPDIR`）。
- `runtime-data\cache`：缓存目录（`XDG_CACHE_HOME`、`PYTHON_EGG_CACHE` 等）。

若存在 `pw-browsers` 目录，程序会将 `PLAYWRIGHT_BROWSERS_PATH` 指向该目录，便于离线运行。

## 6. 常见问题与排障

### 6.1 文件读取失败或提示加密

- 先确认密码是否正确。
- 若已勾选“密码解密失败时尝试 Office/WPS”，请确认本机 COM 可调用 Office/WPS。
- 未安装 `msoffcrypto-tool` 时无法走纯 Python 解密链路。

### 6.2 拨测结果大量超时

- 先降低并发，观察网络或目标站点限流情况。
- 适当提高超时秒数与重试次数。
- 检查目标 URL 是否需要内网/VPN 或特定 DNS 环境。

### 6.3 打包失败（PyInstaller 或文件占用）

- 关闭正在运行的 `SyntheticMonitoringTool.exe` 后重试。
- 使用 `--full-clean` 进行全量构建。
- 杀软可能导致构建慢或锁文件，建议排除 `build`、`dist`、`pw-browsers` 和 Python `site-packages`。

### 6.4 找不到 Inno Setup 编译器（ISCC）

- 安装 Inno Setup 6：<https://jrsoftware.org/isinfo.php>
- 或使用脚本中提示的 `winget` 命令安装后重试。

### 6.5 目标机器离线运行

- 推荐携带完整 `dist\SyntheticMonitoringTool` 目录分发，而非仅复制单个 EXE。
- 确保目录中包含 `pw-browsers`（如需离线拨测）。
