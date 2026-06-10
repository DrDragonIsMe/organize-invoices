# 发票邮件下载工具

自动从邮箱中检索标题含「发票」的邮件，下载附件和邮件正文中的发票下载链接，按月份分目录存放。

![image-20260610090016861](/Users/changxinglong/Library/Application Support/typora-user-images/image-20260610090016861.png)

![image-20260610090239170](/Users/changxinglong/Library/Application Support/typora-user-images/image-20260610090239170.png)

## 功能特性

- **IMAP 协议支持**：兼容 QQ邮箱、163邮箱、Gmail、Outlook 等主流邮箱
- **中文文件夹名自动编码**：内置 IMAP-UTF-7 编解码，支持 `收件箱`、`发票` 等中文文件夹名
- **自动发现文件夹**：配置 `"folders": ["ALL"]` 即可自动扫描所有非系统文件夹
- **附件自动下载**：支持 PDF、PNG、JPG、JPEG、GIF、WEBP、OFD、XML 等格式
- **链接智能追踪**：自动识别邮件正文中的发票下载链接并下载
- **短链接解析**：自动展开 `wxaurl.cn`、`t.cn`、`bit.ly` 等短链接
- **智能过滤**：自动过滤邮件 banner、二维码、广告追踪、缩略图、官网首页等无关链接
- **多格式优先级**：同一发票存在多种下载格式时，按 **PDF > 图片 > OFD > XML > HTML** 优先级只下载一个
- **Playwright SPA 渲染**：对百望云、票通等 Vue/React 预览页自动使用浏览器渲染提取真实下载链接
- **内容查重 + 跨格式去重**：基于 SHA-256 内容哈希查重，发票级跨格式去重避免重复保存
- **按月份归档**：自动创建 `2026-01/`、`2026-02/` 等目录
- **智能去重**：文件已存在自动加序号，避免覆盖
- **Supplemental 补充目录**：支持手动补充发票，自动按日期归类并去重
- **清理工具**：删除无效 HTML 阅读器页面、重建数据库
- **核查报告**：自动生成 HTML + JSON 格式的核查报告，含金额统计与月度汇总
- **Web 配置界面**：可视化配置界面，支持配置自动保存/加载、一键运行/停止和实时日志查看
- **报告列表**：在 Web 界面中直接查看所有已生成的 HTML 报告，点击另开标签页浏览
- **目录管理**：在 Web 界面中一键打开发票目录、补充目录或报告目录（调用系统文件管理器）
- **扫描并生成报告**：无需重新下载邮件，直接在 Web 界面中扫描现有发票目录并重新生成核查报告
- **进度显示**：下载进度条 + 详细日志

## 项目结构

```
.
├── download_invoices.py    # 主程序：IMAP 检索、附件/链接下载、去重归档
├── cleanup_invoices.py     # 清理工具：删除无效 HTML、内容哈希去重、重建数据库
├── generate_report.py      # 报告生成器：扫描目录、解析金额日期、生成 HTML/JSON 报告
├── start.py                # Web 配置界面：启动本地 HTTP 服务，可视化配置和一键运行
├── Start.html              # Web 配置界面主页面
├── Start.bat               # Windows 一键启动脚本
├── Start.command           # macOS 一键启动脚本
├── config.json.example     # 配置文件模板（复制后填写邮箱密码）
├── requirements.txt        # Python 依赖
├── README.md               # 本文件
├── invoices/               # 下载的发票文件（按月份分目录）
├── Supplemental/           # 手动补充的发票（脚本自动归类到月份目录）
├── reports/                # 核查报告（HTML + JSON）
└── web/                    # Web 界面静态文件（由 start.py 提供）
```

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置邮箱信息

复制配置文件模板并填写你的信息：

```bash
cp config.json.example config.json
```

编辑 `config.json`：

```json
{
  "imap_server": "imap.163.com",
  "imap_port": 993,
  "use_ssl": true,
  "email": "your-email@163.com",
  "password": "你的授权码",
  "date_range": {
    "since": "2026-01-01",
    "before": "2026-06-30"
  },
  "target_dir": "./invoices",
  "search_keywords": ["发票", "电票"],
  "folders": ["ALL"]
}
```

> **注意**：QQ邮箱、163邮箱等需要使用**授权码**而非登录密码。

常见邮箱 IMAP 配置：

| 邮箱 | IMAP 服务器 | 端口 | 备注 |
|------|------------|------|------|
| QQ邮箱 | imap.qq.com | 993 | 需开启 IMAP，使用授权码 |
| 163邮箱 | imap.163.com | 993 | 需开启 IMAP，使用授权码；注意开启"收取全部邮件" |
| Gmail | imap.gmail.com | 993 | 需开启两步验证，使用应用密码 |
| Outlook | outlook.office365.com | 993 | 使用邮箱密码 |

### 3. 运行

#### 方式一：命令行运行

```bash
python download_invoices.py -c config.json
```

#### 方式二：Web 配置界面（推荐）

**macOS：**
```bash
双击 Start.command
# 或 python3 start.py
```

**Windows：**
```bash
双击 Start.bat
# 或 python start.py
```

启动后自动打开浏览器，提供以下功能：

- **可视化配置**：邮件服务器、搜索条件、黑名单、日期范围、输出目录等，配置自动保存到 `config.json`
- **一键下载**：支持「仅预览」和「开始下载」两种模式，实时查看运行日志
- **目录管理**：一键打开发票目录、Supplemental 补充目录或报告目录
- **扫描并生成报告**：不重新下载邮件，直接扫描现有发票目录生成最新核查报告
- **报告列表**：查看所有已生成的 HTML 报告，点击另开标签页浏览

> **提示**：配置界面中的任何修改都会自动保存，下次打开时自动恢复上次配置。

运行后目录结构：

```
invoices/
├── 2026-04/
│   ├── 发票_XXX.pdf
│   └── 26112000001768503076.pdf
├── 2026-05/
│   ├── 电子发票_123.jpg
│   └── 来自_京东_invoice_456.pdf
└── unknown/
    └── ...
Supplemental/               # 手动补充的发票放入此处，下次运行自动归类
reports/
├── invoice_report_20260610_143052.json
└── invoice_report_20260610_143052.html
```

## 工作流

推荐按以下顺序使用各工具：

```
1. 下载（download_invoices.py）
   └─ 从邮箱自动下载发票，按月份归档，生成初步报告

2. 补充（Supplemental/ 目录）
   └─ 将手动下载/拍照的发票放入 Supplemental/，下次运行时自动归类

3. 清理（cleanup_invoices.py）
   └─ 删除残留的无效 HTML 阅读器页面，按内容哈希去重，重建数据库

4. 报告（generate_report.py）
   └─ 重新扫描目录，解析所有发票金额和日期，生成最终核查报告
```

## 核心机制

### 链接智能过滤

脚本内置多层过滤机制，避免下载无关内容：

| 过滤类型 | 示例 | 处理方式 |
|---------|------|---------|
| 邮件模板图 | `email_banner.png`、`qrCodeImg.png` | 黑名单直接过滤 |
| 广告追踪 | `ad.efapiao.com` | 黑名单直接过滤 |
| XML 命名空间 | `ns.adobe.com` | 黑名单直接过滤 |
| 官网首页 | `www.baiwang.com` | 黑名单直接过滤 |
| 服务商预览页 | `newtimeai.com` | 黑名单直接过滤 |
| 缩略图 | `meituan.net/scarlett/...` | 黑名单直接过滤 |
| 无法解析的短链接 | `wxaurl.cn/...`（无 302 重定向）| 存在直链时自动丢弃 |
| SPA 空白页 | Vue/React 渲染的空白预览页 | 检测后跳过或 Playwright 渲染 |
| HTML 阅读器 | 绿页云、PDF Viewer 等在线阅读器 | 检测后跳过 |

### 多格式优先级下载

同一封邮件中，如果一张发票提供了多个下载链接（如 OFD + XML + 预览页），脚本会按以下优先级只下载一个：

```
PDF (4分) > 图片 PNG/JPG/WEBP (3分) > OFD (2分) > XML (1分) > HTML (0分)
```

### 发票 ID 分组

脚本会自动提取链接中的发票标识进行分组：

| 平台 | 分组依据 | 示例 |
|------|---------|------|
| 百望云 | `param` 参数 | `bw:A20C11C829C2E7A0B986...` |
| 税务局 | `Fphm`（发票号码）| `tax:26112000001768503076` |
| 美团 | 文件名 hash | `mt:969f96cb1ee9f53f` |
| 票通 | URL 路径 | `vt:6EZSyT7HLx9P4Mead1gn` |

### 金额与日期解析

脚本自动从文件名、PDF 文本、OFD XML 中提取发票金额和开票日期：

- **文件名**：匹配 `金额123.45`、`20260508`、`2026-05-08` 等模式
- **PDF**：提取 "价税合计"、"合计" 后的金额，以及 "开票日期"
- **OFD**：读取 `CustomTag.xml` 中的 `TaxInclusiveTotalAmount` 和 `IssueDate`

## 各脚本详细说明

### download_invoices.py - 主程序

从邮箱检索并下载发票，是核心下载工具。

```bash
# 基本使用
python download_invoices.py

# 指定配置文件
python download_invoices.py -c config.json

# 仅预览模式（不实际下载）
python download_invoices.py --dry-run
```

**主要功能**：
- IMAP 连接与中文文件夹编码
- 附件提取与保存（带内容查重）
- 邮件正文链接提取、短链接解析、智能过滤
- Playwright SPA 页面渲染（百望云、票通等）
- 多格式优先级下载与发票级跨格式去重
- Supplemental 目录自动归类
- HTML + JSON 核查报告生成

### cleanup_invoices.py - 清理工具

清理已下载目录中的无效文件和重复文件，重建数据库。

```bash
# 基本使用（清理 invoices 目录）
python cleanup_invoices.py

# 指定目录
python cleanup_invoices.py --target-dir ./invoices

# 仅预览，不实际删除
python cleanup_invoices.py --dry-run
```

**清理内容**：
1. 删除所有 `.html` 文件（残留的在线阅读器页面）
2. 按内容哈希去重，相同内容只保留一个副本（优先保留文件名含发票号码的）
3. 删除空目录
4. 重建 SQLite 数据库（`.invoice_db`）

### generate_report.py - 报告生成器

扫描发票目录，解析金额和日期，生成核查报告。既可独立运行，也被 `download_invoices.py` 导入调用。

```bash
# 扫描 ./invoices 目录并生成报告
python generate_report.py

# 扫描指定目录
python generate_report.py --dir ./invoices

# 同时处理 Supplemental 目录
python generate_report.py --supplement

# 重建数据库
python generate_report.py --rebuild-db

# 仅预览（不移动 Supplemental 文件）
python generate_report.py --supplement --dry-run
```

**报告内容**：
- 每封邮件/文件的详细信息（标题、发件人、日期、文件列表）
- 发票金额自动解析与汇总
- 按月份汇总金额
- 按发票标识汇总金额
- HTML 可视化报告（可直接浏览器打开）
- JSON 结构化数据

### start.py - Web 配置界面

启动本地 HTTP 服务，提供可视化配置界面和一键运行功能。**零额外依赖**，仅需 Python 标准库。

```bash
# macOS / Linux
python3 start.py

# Windows
python start.py
```

**一键启动脚本**：
- **macOS**：双击 `Start.command`（已添加执行权限）
- **Windows**：双击 `Start.bat`

**功能**：
- 自动查找可用端口（默认从 8765 开始）
- 自动打开浏览器访问配置界面
- 可视化编辑所有配置项，**自动保存到 `config.json`**，下次打开自动恢复
- 一键运行/停止 `download_invoices.py`，实时查看运行日志
- 支持「仅预览」模式，先查看会下载哪些文件再正式执行

**目录管理**：
- **打开发票目录**：在系统文件管理器中打开 `invoices/` 目录
- **打开补充目录**：在系统文件管理器中打开 `Supplemental/` 目录
- **打开报告目录**：在系统文件管理器中打开 `reports/` 目录

**扫描并生成报告**：
无需重新下载邮件，直接在界面中扫描现有发票目录并生成最新核查报告：

| 选项 | 说明 | 默认 |
|------|------|------|
| 处理 Supplemental 补充目录 | 将手动补充的发票自动归类到对应月份 | ☑ 勾选 |
| 重建查重索引 | 重新计算所有文件哈希，重建 SQLite 去重数据库。仅在手动增删改了 `invoices/` 里的文件时需要勾选 | ☐ 不勾选 |
| 仅预览 | 只模拟运行，不实际移动文件 | ☐ 不勾选 |

> **注意**：每次执行都会**重新扫描 `invoices/` 目录**，报告会反映文件的最新增删改情况。只有当你手动改动过发票文件、担心查重数据库不同步时，才需要额外勾选「重建查重索引」。

**报告列表**：
界面会列出 `reports/` 目录下所有已生成的 HTML 报告，按时间倒序排列，点击「查看」按钮即可在浏览器新标签页中打开报告。

**API 端点**：
- `GET /api/config` - 读取当前配置
- `POST /api/config` - 保存配置（合并更新，保留未修改字段）
- `POST /api/run` - 启动下载任务（支持 `dry_run` 参数）
- `POST /api/stop` - 停止运行中的任务
- `GET /api/status` - 获取任务运行状态
- `GET /api/log` - 获取运行日志（后台任务实时写入）
- `GET /api/reports` - 列出已生成的报告文件列表
- `POST /api/open-folder` - 用系统文件管理器打开指定目录（`invoices` / `supplemental` / `reports`）
- `POST /api/regenerate` - 调用 `generate_report.py` 重新扫描并生成报告

## 高级配置

### 指定多个文件夹

```json
{
  "folders": ["INBOX", "发票", "已发送"]
}
```

或使用 `"ALL"` 自动发现所有非系统文件夹：

```json
{
  "folders": ["ALL"]
}
```

### 自定义请求头

```json
{
  "headers": {
    "User-Agent": "Mozilla/5.0 ...",
    "Cookie": "session=xxx"
  }
}
```

### 屏蔽特定发件人

```json
{
  "blocked_senders": ["spam.com", "noreply@ads.com"]
}
```

## 仅预览模式

不实际下载，只查看会处理哪些邮件：

```bash
python download_invoices.py --dry-run
```

在 Web 界面中也可勾选预览模式后运行。

## 核查报告

每次运行结束后，脚本会在 `reports/` 目录下生成两份报告：

- **`invoice_report_YYYYMMDD_HHMMSS.json`**：结构化数据，包含每封邮件的标题、发件人、日期、提取的文件列表
- **`invoice_report_YYYYMMDD_HHMMSS.html`**：可视化报告，可直接在浏览器中打开查看，含金额统计和月度汇总

### 查看报告

**方式一：Web 配置界面（推荐）**

在 Web 界面的「已生成报告」卡片中，会列出所有历史报告，按生成时间倒序排列。点击「查看」按钮即可在浏览器新标签页中打开 HTML 报告。

**方式二：直接打开文件**

```bash
open reports/invoice_report_*.html        # macOS
start reports\invoice_report_*.html       # Windows
```

### 重新生成报告

如果手动补充了发票或修改了目录内容，无需重新下载邮件，可直接重新生成报告：

```bash
# 处理 Supplemental 并生成报告
python generate_report.py --supplement

# 同时重建查重索引（手动改动过 invoices/ 目录时推荐）
python generate_report.py --supplement --rebuild-db
```

或在 Web 配置界面中点击「扫描并生成报告」按钮。"

## 注意事项

1. **授权码 vs 密码**：国内邮箱（QQ、163）一般需要使用授权码
2. **IMAP 开启**：需要在邮箱设置中开启 IMAP 服务
3. **163 邮箱特殊注意**：默认只同步最近 30 天的邮件，如需抓取历史邮件，请登录 163 网页版 → 设置 → 客户端设置 → 开启"收取全部邮件"
4. **网络环境**：部分邮箱（如 Gmail）可能需要代理
5. **下载频率**：建议不要过于频繁运行，避免触发邮箱风控
6. **Playwright**：首次使用 SPA 渲染功能前需执行 `playwright install chromium`
7. **Supplemental 目录**：手动补充的发票放入 `Supplemental/` 目录即可，无需指定日期，脚本会自动从文件名或内容中提取日期并归类

## 日志

运行时会输出详细日志到控制台，也可配置 `log_file`：

```json
{
  "log_file": "./logs/download.log"
}
```

Web 界面中可实时查看运行日志。
