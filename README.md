# 发票邮件下载工具

自动从邮箱中检索标题含「发票」的邮件，下载附件和邮件正文中的发票下载链接，按月份分目录存放。

## 功能特性

- ✅ **IMAP 协议支持**：兼容 QQ邮箱、163邮箱、Gmail、Outlook 等主流邮箱
- ✅ **中文文件夹名自动编码**：内置 IMAP-UTF-7 编解码，支持 `收件箱`、`发票` 等中文文件夹名
- ✅ **自动发现文件夹**：配置 `"folders": ["ALL"]` 即可自动扫描所有非系统文件夹
- ✅ **附件自动下载**：支持 PDF、PNG、JPG、JPEG、GIF、WEBP、OFD、XML 等格式
- ✅ **链接智能追踪**：自动识别邮件正文中的发票下载链接并下载
- ✅ **短链接解析**：自动展开 `wxaurl.cn`、`t.cn`、`bit.ly` 等短链接
- ✅ **智能过滤**：自动过滤邮件 banner、二维码、广告追踪、缩略图、官网首页等无关链接
- ✅ **多格式优先级**：同一发票存在多种下载格式时，按 **PDF > 图片 > OFD > XML > HTML** 优先级只下载一个
- ✅ **按月份归档**：自动创建 `2026-01/`、`2026-02/` 等目录
- ✅ **智能去重**：文件已存在自动加序号，避免覆盖
- ✅ **核查报告**：自动生成 HTML + JSON 格式的核查报告
- ✅ **进度显示**：下载进度条 + 详细日志

## 项目结构

```
.
├── download_invoices.py   # 主程序
├── config.json.example    # 配置文件模板（复制后填写邮箱密码）
├── requirements.txt       # Python 依赖
├── README.md              # 本文件
├── invoices/              # 下载的发票文件（按月份分目录）
├── Supplemental/          # 手动补充的发票（脚本自动归类到月份目录）
└── reports/               # 核查报告（HTML + JSON）
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

```bash
python download_invoices.py -c config.json
```

运行后目录结构：

```
invoices/
├── 2026-04/
│   ├── 发票_XXX.pdf
│   └── 北京某某公司_ofd_查阅需OFD阅读器_1.ofd
├── 2026-05/
│   ├── 电子发票_123.jpg
│   └── 来自_京东_invoice_456.pdf
└── unknown/
    └── ...
Supplemental/              # 手动补充的发票放入此处，下次运行自动归类
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

### 链接过滤白名单

```json
{
  "link_filters": {
    "domain_whitelist": ["example.com", "invoice.provider.com"]
  }
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

## 仅预览模式

不实际下载，只查看会处理哪些邮件：

```bash
python download_invoices.py --dry-run
```

## 核查报告

每次运行结束后，脚本会在 `reports/` 目录下生成两份报告：

- **`invoice_report_YYYYMMDD_HHMMSS.json`**：结构化数据，包含每封邮件的标题、发件人、日期、提取的文件列表
- **`invoice_report_YYYYMMDD_HHMMSS.html`**：可视化报告，可直接在浏览器中打开查看

## 注意事项

1. **授权码 vs 密码**：国内邮箱（QQ、163）一般需要使用授权码
2. **IMAP 开启**：需要在邮箱设置中开启 IMAP 服务
3. **163 邮箱特殊注意**：默认只同步最近 30 天的邮件，如需抓取历史邮件，请登录 163 网页版 → 设置 → 客户端设置 → 开启"收取全部邮件"
4. **网络环境**：部分邮箱（如 Gmail）可能需要代理
5. **下载频率**：建议不要过于频繁运行，避免触发邮箱风控

## 日志

运行时会输出详细日志到控制台，也可配置 `log_file`：

```json
{
  "log_file": "./logs/download.log"
}
```
