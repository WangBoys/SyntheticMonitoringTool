# Ubuntu 脚本包部署说明

本包仅含脚本与配置文件，**不含** Python 依赖与浏览器。请在 Ubuntu 上首次执行 `./setup_ubuntu.sh` 安装依赖。

## 1. 解压

```bash
cd /opt
unzip SyntheticMonitoringTool-scripts.zip -d SyntheticMonitoringTool
cd SyntheticMonitoringTool
```

## 2. 首次安装依赖（只需一次）

```bash
chmod +x setup_ubuntu.sh run_cli.sh install_server_deps.sh
./setup_ubuntu.sh
```

若出现 `/usr/bin/env: 'bash\r': No such file or directory`，说明脚本为 Windows 换行符，执行：

```bash
sed -i 's/\r$//' *.sh
chmod +x *.sh
./setup_ubuntu.sh
```

（重新在 Windows 上运行 `pack_ubuntu_scripts.bat` 打出的包已自动转为 LF。）

Ubuntu 24.04 若 `libasound2` 安装失败，请使用最新脚本包（含 `ubuntu_apt_deps.sh`），或安装完 Python 依赖后执行：`python -m playwright install-deps chromium`。

## 3. 配置与运行

```bash
cp config.example.json config.json
vim config.json
./run_cli.sh --dry-run
./run_cli.sh
```

运行时会输出公网 IP、运营商、国家/省份/城市，随后为拨测进度。

## 4. 新增配置选项

从 v1.1.0 开始，`config.json` 支持以下新增字段：

### probe_mode - 探测模式

- `lightweight`（默认）：使用 httpx + BeautifulSoup，速度快，内存占用低，适合大规模批量测试
- `browser`：使用 Playwright Chromium 浏览器，支持 JavaScript 渲染页面，适合需要测试 SPA 应用的场景

```json
"probe_mode": "lightweight"
```

### extract_title - 提取网站标题

- `true`（默认）：从响应中解析 HTML 并提取 `<title>` 标签内容
- `false`：跳过标题提取，提升性能

```json
"extract_title": true
```

### track_redirects - 跟踪重定向

- `true`（默认）：记录重定向次数和最终 URL
- `false`：不跟踪重定向信息

```json
"track_redirects": true
```

## 5. 依赖说明

新增依赖包：
- `httpx>=0.27.0`：轻量级 HTTP 客户端，用于 lightweight 模式
- `beautifulsoup4>=4.12.0`：HTML 解析库，用于提取网站标题

这些依赖已在 `requirements-linux.txt` 中包含，执行 `./setup_ubuntu.sh` 时会自动安装。
