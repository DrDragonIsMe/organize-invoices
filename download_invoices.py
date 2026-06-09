#!/usr/bin/env python3
"""
发票邮件附件及链接追踪下载工具

功能：
1. 通过 IMAP 检索标题含指定关键词的邮件，支持多关键词和日期范围过滤
2. 下载邮件中的附件（PDF、图片、OFD、XML 等），按月份分目录存放
3. 追踪邮件正文中的发票链接并下载，支持短链接解析和链接打分排序
4. 内容哈希查重 + URL 级别查重 + 发票级跨格式去重（同一张发票只保留最优格式）
5. Playwright SPA 渲染：自动识别百望云、票通等 Vue/React 预览页，提取真实下载链接
6. HTML 阅读器/空白页检测：自动跳过绿页云等在线阅读器框架页和 JS 空白页
7. 从 HTML 中提取 base64 编码的 PDF/图片数据
8. 支持 blocked_senders 发件人黑名单过滤
9. 自动从文件内容（OFD、PDF）中提取发票号码并重命名文件
10. 自动提取发票金额和开票日期，汇总到报告
11. 处理 Supplemental 目录（手动补充的发票自动归类到对应月份）
12. 生成 JSON 和 HTML 核查报告，含月度/发票级金额汇总

使用方法：
    python download_invoices.py
    python download_invoices.py --config custom_config.json
    python download_invoices.py --dry-run    # 仅预览，不下载
"""

import argparse
import base64
import email
import hashlib
import imaplib
import json
import os
import re
import sqlite3
import ssl
import subprocess
import sys

from generate_report import ReportGenerator
import tempfile
from datetime import datetime
from email.header import decode_header
from pathlib import Path
from time import sleep
from urllib.parse import urljoin, urlparse
from typing import Optional, List, Dict, Any, Tuple

import requests
from bs4 import BeautifulSoup
from dateutil import parser as date_parser
from tqdm import tqdm

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False


# ---- IMAP-UTF-7 编解码（用于中文文件夹名） ----

def _imap_utf7_encode(s: str) -> str:
    """将字符串编码为 IMAP-UTF-7（RFC 3501）"""
    result = []
    buf = bytearray()
    for ch in s:
        o = ord(ch)
        if 0x20 <= o <= 0x7E:
            if buf:
                b64 = base64.b64encode(bytes(buf)).decode('ascii').rstrip('=').replace('/', ',')
                result.append('&' + b64 + '-')
                buf = bytearray()
            if ch == '&':
                result.append('&-')
            else:
                result.append(ch)
        else:
            buf.extend(ch.encode('utf-16-be'))
    if buf:
        b64 = base64.b64encode(bytes(buf)).decode('ascii').rstrip('=').replace('/', ',')
        result.append('&' + b64 + '-')
    return ''.join(result)


def _imap_utf7_decode(s: str) -> str:
    """将 IMAP-UTF-7 解码为字符串"""
    result = []
    i = 0
    while i < len(s):
        if s[i] == '&' and i + 1 < len(s) and s[i+1] != '-':
            end = s.find('-', i)
            if end == -1:
                result.append(s[i:])
                break
            b64_part = s[i+1:end].replace(',', '/')
            while len(b64_part) % 4:
                b64_part += '='
            try:
                decoded = base64.b64decode(b64_part).decode('utf-16-be')
                result.append(decoded)
            except Exception:
                result.append(s[i:end+1])
            i = end + 1
        elif s[i:i+2] == '&-':
            result.append('&')
            i += 2
        else:
            result.append(s[i])
            i += 1
    return ''.join(result)


# ---- 日志工具 ----

class Logger:
    def __init__(self, log_file: Optional[str] = None):
        self.log_file = log_file
        if log_file:
            Path(log_file).parent.mkdir(parents=True, exist_ok=True)
            open(log_file, 'a', encoding='utf-8').close()

    def _write(self, level: str, message: str):
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        line = f"[{ts}] [{level}] {message}"
        print(line)
        if self.log_file:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(line + '\n')

    def info(self, message: str):
        self._write('INFO', message)

    def warning(self, message: str):
        self._write('WARN', message)

    def error(self, message: str):
        self._write('ERROR', message)

    def success(self, message: str):
        self._write('OK', message)


# ---- 发票数据库（内容查重 + 跨格式去重） ----

class InvoiceDatabase:
    """SQLite 数据库，记录已下载发票的内容哈希和格式优先级，实现查重和跨格式去重。"""

    # 格式优先级：数值越高越优先
    FORMAT_PRIORITY = {
        '.pdf': 4,
        '.png': 3,
        '.jpg': 3,
        '.jpeg': 3,
        '.gif': 3,
        '.webp': 3,
        '.ofd': 2,
        '.xml': 1,
        '.html': 0,
        '.bin': 0,
        '.tmp': 0,
        '': 0,
    }

    def __init__(self, db_path: str):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path)
        self._init_tables()

    def _init_tables(self):
        self.conn.executescript('''
            CREATE TABLE IF NOT EXISTS downloaded_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_id TEXT,
                content_hash TEXT UNIQUE,
                file_path TEXT,
                file_format TEXT,
                format_priority INTEGER,
                file_size INTEGER,
                source_email TEXT,
                source_url TEXT,
                downloaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            CREATE INDEX IF NOT EXISTS idx_invoice_id ON downloaded_files(invoice_id);
            CREATE INDEX IF NOT EXISTS idx_content_hash ON downloaded_files(content_hash);

            CREATE TABLE IF NOT EXISTS invoice_best_format (
                invoice_id TEXT PRIMARY KEY,
                best_format TEXT,
                best_priority INTEGER,
                file_path TEXT,
                content_hash TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        ''')
        self.conn.commit()

    def get_priority(self, ext: str) -> int:
        return self.FORMAT_PRIORITY.get(ext.lower(), 0)

    def hash_exists(self, content_hash: str) -> bool:
        cur = self.conn.execute(
            'SELECT 1 FROM downloaded_files WHERE content_hash = ? LIMIT 1',
            (content_hash,)
        )
        return cur.fetchone() is not None

    def url_exists(self, source_url: str) -> bool:
        """检查该 URL 是否已下载过"""
        if not source_url:
            return False
        cur = self.conn.execute(
            'SELECT 1 FROM downloaded_files WHERE source_url = ? LIMIT 1',
            (source_url,)
        )
        return cur.fetchone() is not None

    def get_best_for_invoice(self, invoice_id: str) -> Optional[Dict[str, Any]]:
        """获取某发票已下载的最高优先级格式记录"""
        if not invoice_id:
            return None
        cur = self.conn.execute(
            'SELECT best_format, best_priority, file_path, content_hash FROM invoice_best_format WHERE invoice_id = ?',
            (invoice_id,)
        )
        row = cur.fetchone()
        if row:
            return {
                'best_format': row[0],
                'best_priority': row[1],
                'file_path': row[2],
                'content_hash': row[3],
            }
        return None

    def record_file(self, invoice_id: str, content_hash: str, file_path: str,
                    file_format: str, file_size: int, source_email: str = '',
                    source_url: str = '') -> bool:
        """记录下载的文件。如果该发票已有更高优先级格式，返回 False 并删除新文件。

        如果新文件优先级更高，会删除旧文件并更新数据库记录。
        """
        priority = self.get_priority(file_format)
        abs_path = str(Path(file_path).absolute())

        # 1. 内容哈希去重
        if self.hash_exists(content_hash):
            self._remove_file(abs_path)
            return False

        # 2. 发票级跨格式去重
        if invoice_id:
            best = self.get_best_for_invoice(invoice_id)
            if best:
                if best['best_priority'] >= priority:
                    # 已有更高或相等优先级，删除新文件
                    self._remove_file(abs_path)
                    return False
                else:
                    # 新文件优先级更高，删除旧文件
                    self._remove_file(best['file_path'])

        # 3. 写入数据库（存储绝对路径）
        try:
            self.conn.execute(
                '''INSERT INTO downloaded_files
                   (invoice_id, content_hash, file_path, file_format, format_priority, file_size, source_email, source_url)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                (invoice_id, content_hash, abs_path, file_format, priority, file_size, source_email, source_url)
            )
        except sqlite3.IntegrityError:
            # 哈希冲突（并发场景）
            self._remove_file(abs_path)
            return False

        # 4. 更新该发票的最佳格式记录
        if invoice_id:
            self.conn.execute(
                '''INSERT INTO invoice_best_format (invoice_id, best_format, best_priority, file_path, content_hash)
                   VALUES (?, ?, ?, ?, ?)
                   ON CONFLICT(invoice_id) DO UPDATE SET
                       best_format=excluded.best_format,
                       best_priority=excluded.best_priority,
                       file_path=excluded.file_path,
                       content_hash=excluded.content_hash,
                       updated_at=CURRENT_TIMESTAMP
                   WHERE excluded.best_priority >= invoice_best_format.best_priority''',
                (invoice_id, file_format, priority, abs_path, content_hash)
            )

        self.conn.commit()
        return True

    @staticmethod
    def _remove_file(file_path: str):
        try:
            p = Path(file_path)
            # 相对路径尝试从项目目录解析
            if not p.is_absolute():
                project_dir = Path(__file__).parent.resolve()
                p = project_dir / p
            p.unlink(missing_ok=True)
        except Exception:
            pass

    def close(self):
        self.conn.close()


# ---- 文件下载工具 ----

class Downloader:
    """通用下载器，支持重试、进度条、限速等"""

    def __init__(self, headers: Dict[str, str], timeout: int = 60, max_retries: int = 3):
        self.session = requests.Session()
        self.session.headers.update(headers)
        self.timeout = timeout
        self.max_retries = max_retries

    def download(self, url: str, dest_path: Path, expected_size: Optional[int] = None) -> bool:
        """下载文件到指定路径，返回是否成功"""
        dest_path.parent.mkdir(parents=True, exist_ok=True)

        for attempt in range(1, self.max_retries + 1):
            try:
                with self.session.get(url, stream=True, timeout=self.timeout, allow_redirects=True) as resp:
                    resp.raise_for_status()

                    cl = resp.headers.get('content-length', '')
                    try:
                        total = expected_size or int(cl)
                    except (ValueError, TypeError):
                        total = expected_size or 0
                    desc = dest_path.name[:30]

                    with open(dest_path, 'wb') as f, tqdm(
                        total=total, unit='B', unit_scale=True, desc=desc, leave=False
                    ) as pbar:
                        for chunk in resp.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                                pbar.update(len(chunk))
                return True
            except Exception as e:
                if attempt == self.max_retries:
                    return False
                sleep(2 ** attempt)
        return False

    def fetch_text(self, url: str) -> Optional[str]:
        """获取网页文本内容"""
        try:
            resp = self.session.get(url, timeout=self.timeout, allow_redirects=True)
            resp.raise_for_status()
            return resp.text
        except Exception:
            return None


# ---- 邮件解析工具 ----

class EmailParser:
    """解析邮件内容，提取附件信息和链接"""

    # 常见发票附件扩展名
    INVOICE_EXTENSIONS = {'.pdf', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.ofd', '.xml'}

    # 邮件中常见的链接文本特征
    LINK_TEXT_PATTERNS = [
        re.compile(r'发票', re.I),
        re.compile(r'下载', re.I),
        re.compile(r'电子发票', re.I),
        re.compile(r'查看', re.I),
        re.compile(r'点击.*下载', re.I),
    ]

    # URL 路径特征
    URL_PATH_PATTERNS = [
        re.compile(r'invoice', re.I),
        re.compile(r'fapiao', re.I),
        re.compile(r'fp[-_]?\d+', re.I),
        re.compile(r'pdf|png|jpg|jpeg', re.I),
    ]

    # URL 黑名单：邮件模板图、广告追踪、logo、XML命名空间、官网首页、缩略图等
    URL_BLACKLIST_PATTERNS = [
        re.compile(r'email_banner|qrCodeImg|bwcloud_banner|angwlogo', re.I),
        re.compile(r'ad\.efapiao\.com', re.I),
        re.compile(r'ns\.adobe\.com', re.I),
        re.compile(r'templates/img|cms/templates', re.I),
        re.compile(r'www\.baiwang\.com', re.I),
        re.compile(r'meituan\.net/scarlett', re.I),
    ]

    def __init__(self, downloader: Downloader, logger: Logger):
        self.downloader = downloader
        self.logger = logger

    @staticmethod
    def decode_header_str(header_value: str) -> str:
        """解码邮件头中的编码字符串"""
        if not header_value:
            return ''
        parts = decode_header(header_value)
        result = []
        for part, charset in parts:
            if isinstance(part, bytes):
                if not charset or charset.lower() in ('unknown-8bit',):
                    charset = 'utf-8'
                result.append(part.decode(charset, errors='replace'))
            else:
                result.append(part)
        return ''.join(result)

    @staticmethod
    def get_email_date(msg: email.message.EmailMessage) -> Optional[datetime]:
        """提取邮件日期，支持多种格式"""
        date_str = msg.get('Date')
        if date_str:
            try:
                return date_parser.parse(date_str)
            except Exception:
                pass
        # 尝试从 Received 头提取（最后一个 Received 通常是最早的）
        received = msg.get_all('Received', [])
        for r in reversed(received):
            try:
                # Received: from ... by ... with ...; Mon, 18 Apr 2026 12:34:56 +0800
                if ';' in r:
                    dt_part = r.split(';')[-1].strip()
                    return date_parser.parse(dt_part)
            except Exception:
                pass
        return None

    def extract_attachments(self, msg: email.message.EmailMessage) -> List[Dict[str, Any]]:
        """从邮件中提取所有附件信息"""
        attachments = []
        for part in msg.walk():
            content_disposition = part.get_content_disposition() or ''
            filename = part.get_filename()
            if filename:
                filename = self.decode_header_str(filename)
            else:
                continue

            payload = part.get_payload(decode=True)
            if not payload:
                continue

            # 区分 attachment 和 inline：
            # - attachment：按扩展名判断（常见发票格式）
            # - inline / 无 Content-Disposition：必须文件名明确像发票才接受
            #   防止账单、广告邮件中的 banner 图片被误下载
            if 'attachment' not in content_disposition.lower():
                if not self._looks_like_invoice_name(filename):
                    continue

            ext = Path(filename).suffix.lower()
            if ext not in self.INVOICE_EXTENSIONS and not self._looks_like_invoice_name(filename):
                continue

            attachments.append({
                'filename': filename,
                'content': payload,
                'size': len(payload),
                'type': 'attachment',
            })

        return attachments

    def extract_links(self, msg: email.message.EmailMessage, base_url: str = '') -> List[Dict[str, Any]]:
        """从邮件正文中提取可能的发票下载链接"""
        links = []
        seen = set()

        html_body = self._get_html_body(msg)
        text_body = self._get_text_body(msg)

        if html_body:
            links.extend(self._parse_html_links(html_body, base_url))

        if text_body:
            links.extend(self._parse_text_links(text_body))

        unique = []
        for link in links:
            url = link['url']
            if not url.startswith(('http://', 'https://')):
                continue
            if url not in seen:
                seen.add(url)
                unique.append(link)
        return unique

    def _get_html_body(self, msg: email.message.EmailMessage) -> Optional[str]:
        """获取邮件 HTML 正文"""
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == 'text/html':
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or 'utf-8'
                    try:
                        return payload.decode(charset, errors='replace')
                    except Exception:
                        return payload.decode('utf-8', errors='replace')
        return None

    def _get_text_body(self, msg: email.message.EmailMessage) -> Optional[str]:
        """获取邮件纯文本正文"""
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == 'text/plain':
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or 'utf-8'
                    try:
                        return payload.decode(charset, errors='replace')
                    except Exception:
                        return payload.decode('utf-8', errors='replace')
        return None

    def _parse_html_links(self, html: str, base_url: str) -> List[Dict[str, Any]]:
        """解析 HTML 中的链接"""
        links = []
        try:
            soup = BeautifulSoup(html, 'html.parser')
        except Exception:
            return links

        for tag in soup.find_all('a', href=True):
            url = urljoin(base_url, tag['href'].strip())
            text = tag.get_text(strip=True)
            score = self._score_link(url, text)
            if score > 0:
                links.append({
                    'url': url,
                    'text': text,
                    'score': score,
                    'type': 'html_link',
                })

        for tag in soup.find_all('img', src=True):
            url = urljoin(base_url, tag['src'].strip())
            if self._is_invoice_image_url(url):
                links.append({
                    'url': url,
                    'text': tag.get('alt', ''),
                    'score': 5,
                    'type': 'image_src',
                })

        return links

    def _parse_text_links(self, text: str) -> List[Dict[str, Any]]:
        """从纯文本中提取 URL"""
        links = []
        url_pattern = re.compile(r'https?://[^\s<>"\'`\)\]\}]+', re.I)
        for match in url_pattern.finditer(text):
            url = match.group(0).rstrip('.,;:!?')
            score = self._score_link(url, '')
            if score > 0:
                links.append({
                    'url': url,
                    'text': '',
                    'score': score,
                    'type': 'text_link',
                })
        return links

    def _score_link(self, url: str, text: str) -> int:
        """给链接打分，判断是否为发票相关链接"""
        for pattern in self.URL_BLACKLIST_PATTERNS:
            if pattern.search(url):
                return 0

        score = 0
        url_lower = url.lower()
        text_lower = text.lower()

        for pattern in self.LINK_TEXT_PATTERNS:
            if pattern.search(text_lower):
                score += 2

        for pattern in self.URL_PATH_PATTERNS:
            if pattern.search(url_lower):
                score += 2

        parsed = urlparse(url)
        path = parsed.path.lower()
        if any(path.endswith(ext) for ext in self.INVOICE_EXTENSIONS):
            score += 3

        invoice_domains = [
            'fapiao', 'invoice', 'einvoice', 'fp',
            'alicdn', 'taobao', 'tmall', 'jd.com',
            'wechat', 'qq.com', 'dingtalk',
            'baiwang.com', 'chinatax.gov.cn', 'vpiaotong.com',
            'crestv.cn', 'meituan.net',
        ]
        for domain in invoice_domains:
            if domain in parsed.netloc.lower():
                score += 1

        if parsed.netloc.lower() in {'t.cn', 'bit.ly', 'tinyurl.com', 'goo.gl', 'dwz.cn'}:
            score = max(1, score - 2)

        return score

    def _is_invoice_image_url(self, url: str) -> bool:
        """判断 URL 是否可能是发票图片"""
        for pattern in self.URL_BLACKLIST_PATTERNS:
            if pattern.search(url):
                return False

        url_lower = url.lower()
        if not any(url_lower.endswith(ext) for ext in {'.png', '.jpg', '.jpeg', '.gif', '.webp'}):
            return False
        return any(pattern.search(url_lower) for pattern in self.URL_PATH_PATTERNS)

    @staticmethod
    def _looks_like_invoice_name(filename: str) -> bool:
        """判断文件名是否像发票"""
        fn_lower = filename.lower()
        keywords = ['发票', 'invoice', 'fapiao', 'fp', '电子发票', '增值税']
        return any(kw in fn_lower for kw in keywords)

    def resolve_short_link(self, url: str) -> Optional[str]:
        """尝试解析短链接到真实 URL"""
        try:
            resp = self.downloader.session.head(url, allow_redirects=True, timeout=30)
            final = resp.url
            if final != url:
                return final
        except Exception:
            pass
        return None

    def guess_filename_from_url(self, url: str, content_type: Optional[str] = None) -> str:
        """从 URL 和响应头猜测文件名"""
        parsed = urlparse(url)
        path = parsed.path
        if path and '/' in path:
            name = path.split('/')[-1]
            if name and '.' in name:
                return name

        ext = '.bin'
        if content_type:
            mapping = {
                'application/pdf': '.pdf',
                'image/png': '.png',
                'image/jpeg': '.jpg',
                'image/jpg': '.jpg',
                'image/gif': '.gif',
                'image/webp': '.webp',
            }
            for ct, e in mapping.items():
                if ct in content_type:
                    ext = e
                    break

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        return f"invoice_{timestamp}{ext}"


# ---- 主程序 ----

class InvoiceDownloader:
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.logger = Logger(log_file=config.get('log_file'))
        self.downloader = Downloader(
            headers=config.get('headers', {}),
            timeout=config.get('timeout', 60),
            max_retries=config.get('max_retries', 3),
        )
        self.parser = EmailParser(self.downloader, self.logger)
        self.target_dir = Path(config['target_dir'])
        self.target_dir.mkdir(parents=True, exist_ok=True)
        self.dry_run = config.get('dry_run', False)

        # 初始化查重数据库
        db_path = str(self.target_dir / '.invoice_db')
        self.db = InvoiceDatabase(db_path)

        # 统计
        self.stats = {
            'emails_checked': 0,
            'attachments_found': 0,
            'attachments_saved': 0,
            'attachments_skipped': 0,
            'links_found': 0,
            'links_downloaded': 0,
            'links_skipped_dup': 0,
            'links_skipped_spa': 0,
            'links_failed': 0,
            'links_manual': 0,
        }

        # 邮件-文件对应记录
        self.records: List[Dict[str, Any]] = []

    def run(self):
        """主入口"""
        self.logger.info("=" * 50)
        self.logger.info("发票邮件下载工具启动")
        self.logger.info("=" * 50)

        imap = self._connect_imap()
        if not imap:
            self.logger.error("IMAP 连接失败，退出")
            return

        try:
            configured_folders = self.config.get('folders', ['INBOX'])
            folders_to_process = []

            for folder in configured_folders:
                if folder == 'ALL':
                    all_folders = self._list_imap_folders(imap)
                    skip_names = {'草稿箱', '已发送', '已删除', '垃圾邮件', '已删除邮件', 'Deleted Items', 'Sent Messages'}
                    for encoded_name, decoded_name in all_folders:
                        if decoded_name not in skip_names:
                            folders_to_process.append((encoded_name, decoded_name))
                    break
                else:
                    folders_to_process.append((folder, folder))

            for encoded_name, decoded_name in folders_to_process:
                self.logger.info(f"正在处理文件夹: {decoded_name}")
                try:
                    self._process_folder(imap, encoded_name, decoded_name)
                except Exception as e:
                    self.logger.warning(f"处理文件夹 {decoded_name} 时出错: {e}")
        finally:
            try:
                imap.logout()
            except Exception:
                pass

        # ---- 处理 Supplemental 目录 ----
        gen = ReportGenerator(self.target_dir, records=self.records, db=self.db, logger=self.logger)
        supplement_records = gen.process_supplemental(dry_run=self.dry_run)
        self.records.extend(supplement_records)

        # ---- 生成报告 ----
        if self.records:
            gen.generate_reports()
        else:
            self.logger.info("没有匹配到含发票的邮件，跳过报告生成")

        self._print_stats()
        self.db.close()

    def _list_imap_folders(self, imap: imaplib.IMAP4) -> List[tuple]:
        """列出所有 IMAP 文件夹"""
        result = []
        try:
            _, folders = imap.list()
            if folders:
                for f in folders:
                    decoded = f.decode('utf-8', errors='replace')
                    parts = decoded.split('"')
                    if len(parts) >= 3:
                        encoded_name = parts[-2]
                        decoded_name = _imap_utf7_decode(encoded_name)
                        result.append((encoded_name, decoded_name))
        except Exception as e:
            self.logger.warning(f"列出文件夹失败: {e}")
        return result

    def _connect_imap(self) -> Optional[imaplib.IMAP4]:
        """连接 IMAP 服务器"""
        server = self.config['imap_server']
        port = self.config.get('imap_port', 993)
        use_ssl = self.config.get('use_ssl', True)

        try:
            if use_ssl:
                context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
                context.minimum_version = ssl.TLSVersion.TLSv1_2
                context.check_hostname = False
                context.verify_mode = ssl.CERT_NONE
                imap = imaplib.IMAP4_SSL(server, port, ssl_context=context)
            else:
                imap = imaplib.IMAP4(server, port)

            imap.login(self.config['email'], self.config['password'])

            try:
                tag = imap._new_tag()
                imap.send(tag + b' ID ("name" "PythonIMAP" "version" "1.0" "vendor" "Python")\r\n')
                while True:
                    line = imap.readline()
                    if line.startswith(tag):
                        break
            except Exception:
                pass

            self.logger.success(f"成功登录 IMAP 服务器: {server}")
            return imap
        except Exception as e:
            self.logger.error(f"IMAP 登录失败: {e}")
            return None

    def _process_folder(self, imap: imaplib.IMAP4, encoded_name: str, decoded_name: str = ''):
        """处理指定文件夹中的邮件"""
        display_name = decoded_name or encoded_name
        status, data = imap.select(encoded_name)
        if status != 'OK':
            self.logger.warning(f"无法选择文件夹 {display_name}: {data}")
            return

        search_criteria = self._build_search_criteria()
        self.logger.info(f"搜索条件: {search_criteria}")

        status, msg_ids = imap.search(None, search_criteria)
        if status != 'OK' or not msg_ids[0]:
            self.logger.info(f"文件夹 {display_name} 中没有匹配的邮件")
            return

        ids = msg_ids[0].split()
        self.logger.info(f"找到 {len(ids)} 封匹配邮件")

        for msg_id in tqdm(ids, desc="处理邮件"):
            self._process_email(imap, msg_id)

    def _build_search_criteria(self) -> str:
        """构建 IMAP 搜索条件"""
        criteria = []
        date_range = self.config.get('date_range', {})
        if 'since' in date_range:
            since = datetime.strptime(date_range['since'], '%Y-%m-%d')
            criteria.append(f'SINCE {since.strftime("%d-%b-%Y")}')
        if 'before' in date_range:
            before = datetime.strptime(date_range['before'], '%Y-%m-%d')
            criteria.append(f'BEFORE {before.strftime("%d-%b-%Y")}')
        criteria.append('NOT DELETED')
        return '(' + ' '.join(criteria) + ')'

    def _process_email(self, imap: imaplib.IMAP4, msg_id: bytes):
        """处理单封邮件"""
        status, data = imap.fetch(msg_id, '(RFC822)')
        if status != 'OK' or not data or not data[0]:
            return

        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        self.stats['emails_checked'] += 1

        subject = self.parser.decode_header_str(msg.get('Subject', ''))
        from_addr = self.parser.decode_header_str(msg.get('From', ''))
        email_date = self.parser.get_email_date(msg)

        # 如果 Date 头无法解析，尝试从 IMAP INTERNALDATE 获取
        if not email_date:
            try:
                _, idata = imap.fetch(msg_id, '(INTERNALDATE)')
                if idata and idata[0]:
                    # 格式: b'123 (INTERNALDATE "18-Apr-2026 12:34:56 +0800")'
                    internal_str = idata[0].decode('utf-8', errors='replace')
                    m = re.search(r'INTERNALDATE "([^"]+)"', internal_str)
                    if m:
                        email_date = date_parser.parse(m.group(1))
            except Exception:
                pass

        # 支持单关键词或关键词列表
        keywords = self.config.get('search_keywords', self.config.get('search_keyword', '发票'))
        if isinstance(keywords, str):
            keywords = [keywords]

        # 策略 1: 标题匹配（最可靠）
        subject_lower = subject.lower()
        matched = any(kw.lower() in subject_lower for kw in keywords)

        # 策略 2: 如果标题不匹配，检查发件人域名 + 正文关键词
        # 默认黑名单为空，所有邮件都有效；可通过配置 blocked_senders 排除特定域名
        if not matched:
            from_lower = from_addr.lower()
            blocked_senders = self.config.get('blocked_senders', [])
            is_blocked_sender = any(domain in from_lower for domain in blocked_senders)

            body_text = ''
            html_body = self.parser._get_html_body(msg)
            text_body = self.parser._get_text_body(msg)
            if text_body:
                body_text = text_body
            elif html_body:
                body_text = re.sub(r'<[^>]+>', ' ', html_body)
            body_lower = body_text.lower()

            if not is_blocked_sender:
                matched = any(kw.lower() in body_lower for kw in keywords)
            else:
                self.logger.info(f"  发件人 {from_addr} 在黑名单中，跳过正文匹配")
                matched = False

        if not matched:
            return

        self.logger.info(f"处理邮件: [{subject}] 来自: {from_addr}")

        record = {
            'msg_id': msg_id.decode(),
            'subject': subject,
            'from': from_addr,
            'date': email_date.isoformat() if email_date else '',
            'date_display': email_date.strftime('%Y-%m-%d %H:%M') if email_date else '未知',
            'month': email_date.strftime('%Y-%m') if email_date else 'unknown',
            'files': [],
        }

        if email_date:
            month_dir = self.target_dir / email_date.strftime('%Y-%m')
        else:
            month_dir = self.target_dir / 'unknown'
        month_dir.mkdir(parents=True, exist_ok=True)

        # ---- 1. 处理附件 ----
        attachments = self.parser.extract_attachments(msg)
        if attachments:
            self.logger.info(f"  发现 {len(attachments)} 个附件: {[a['filename'] for a in attachments]}")
        self.stats['attachments_found'] += len(attachments)

        for att in attachments:
            file_info = self._save_attachment(att, month_dir, subject)
            if file_info:
                if file_info.get('status') == '成功':
                    self.stats['attachments_saved'] += 1
                elif file_info.get('status') == '已存在/重复':
                    self.stats['attachments_skipped'] += 1
                record['files'].append(file_info)

        # ---- 2. 处理邮件中的链接 ----
        links = self.parser.extract_links(msg)
        if links:
            self.logger.info(f"  发现 {len(links)} 个链接: {[(l['url'][:60], l['score']) for l in links]}")
        self.stats['links_found'] += len(links)

        link_groups = self._group_links_by_invoice(links)

        for group in link_groups:
            file_info = self._download_best_link(group, month_dir, subject)
            if file_info:
                if file_info.get('status') == '成功':
                    self.stats['links_downloaded'] += 1
                elif file_info.get('status') in ('已存在/重复', '已存在/重复(URL)'):
                    self.stats['links_skipped_dup'] += 1
                elif file_info.get('status') in ('JS渲染空白页，跳过', '在线阅读器页面，跳过'):
                    self.stats['links_skipped_spa'] += 1
                record['files'].append(file_info)
            else:
                self.stats['links_failed'] += 1
                record['files'].append({
                    'type': '下载失败',
                    'filename': '该发票所有格式均下载失败',
                    'path': '',
                    'size': 0,
                    'source_url': group[0]['url'] if group else '',
                    'status': '错误：PDF/图片/OFD/XML 均下载失败',
                })

        if record['files']:
            self.records.append(record)

    # ---- 查重与发票标识提取 ----

    @staticmethod
    def _compute_hash(data: bytes) -> str:
        """计算内容的 SHA-256 哈希"""
        return hashlib.sha256(data).hexdigest()

    @staticmethod
    def _compute_file_hash(path: Path) -> str:
        """计算文件内容的 SHA-256 哈希"""
        sha = hashlib.sha256()
        with open(path, 'rb') as f:
            for chunk in iter(lambda: f.read(65536), b''):
                sha.update(chunk)
        return sha.hexdigest()

    @staticmethod
    def _extract_invoice_id(filename: str) -> Optional[str]:
        """从文件名中提取发票唯一标识"""
        # 1. 20 位纯数字（发票代码/号码）
        m = re.search(r'\b(\d{20})\b', filename)
        if m:
            return m.group(1)
        # 2. 12-20 位纯数字前缀
        m = re.search(r'^(\d{12,20})', Path(filename).stem)
        if m:
            return m.group(1)
        # 3. 美团模式：商户名 + 金额整数
        # 匹配 "商户名_发票金额..."，提取商户名和金额数字（忽略小数点）
        if '发票金额' in filename:
            merchant_match = re.search(r'^([^_]+)', filename)
            if merchant_match:
                merchant = merchant_match.group(1)
                amt_match = re.search(r'发票金额\D*(\d+)', filename)
                if amt_match:
                    return f"mt:{merchant}_{amt_match.group(1)}"
                return f"mt:{merchant}"
        # 4. 通用模式：取前 20 个字符作为 fallback
        stem = Path(filename).stem
        if len(stem) > 5:
            return stem[:30]
        return None

    @staticmethod
    def _extract_invoice_id_from_file(path: Path) -> Optional[str]:
        """从已下载的文件内容中提取发票号码（20位数字）

        支持 OFD（ZIP 内 XML）和 PDF（文本提取）。
        """
        try:
            ext = path.suffix.lower()
            if ext == '.ofd':
                import zipfile
                with zipfile.ZipFile(path, 'r') as z:
                    for name in z.namelist():
                        if name.endswith('.xml'):
                            with z.open(name) as f:
                                content = f.read().decode('utf-8', errors='replace')
                                m = re.search(r'\b(\d{20})\b', content)
                                if m:
                                    return m.group(1)
            elif ext == '.pdf':
                try:
                    result = subprocess.run(['strings', str(path)], capture_output=True, text=True, timeout=5)
                    m = re.search(r'\b(\d{20})\b', result.stdout)
                    if m:
                        return m.group(1)
                except Exception:
                    pass
        except Exception:
            pass
        return None

    def _try_rename_with_invoice_id(self, path: Path) -> Path:
        """尝试从文件内容中提取发票号码并重命名文件"""
        invoice_id = self._extract_invoice_id_from_file(path)
        if invoice_id and invoice_id not in path.name:
            new_name = f"{invoice_id}{path.suffix}"
            new_path = self._unique_path(path.parent / self._sanitize_filename(new_name))
            try:
                path.rename(new_path)
                self.logger.info(f"  根据发票号码重命名: {path.name} -> {new_path.name}")
                return new_path
            except Exception:
                pass
        return path

    # ---- 金额提取 ----

    @staticmethod
    def _extract_amount_from_filename(filename: str) -> Optional[float]:
        """从文件名提取金额"""
        m = re.search(r'(?:发票)?金额\D*(\d+(?:\.\d+)?)', filename)
        if m:
            try:
                return float(m.group(1))
            except ValueError:
                pass
        return None

    @staticmethod
    def _extract_amount_from_ofd(path: Path) -> Optional[float]:
        """从 OFD 文件提取金额（优先价税合计）"""
        try:
            import zipfile
            with zipfile.ZipFile(path, 'r') as z:
                # 方法1: CustomTag.xml 中 TaxInclusiveTotalAmount
                if 'Doc_0/Tags/CustomTag.xml' in z.namelist():
                    with z.open('Doc_0/Tags/CustomTag.xml') as f:
                        content = f.read().decode('utf-8', errors='replace')
                        m = re.search(r'TaxInclusiveTotalAmount[^>]*>(\d+\.\d{2})<', content)
                        if m:
                            return float(m.group(1))

                # 方法2: 从所有 XML 的 TextCode 中提取合理金额
                amounts = []
                for name in z.namelist():
                    if name.endswith('.xml'):
                        with z.open(name) as f:
                            content = f.read().decode('utf-8', errors='replace')
                            for m in re.finditer(r'>(\d+\.\d{2})<', content):
                                val = float(m.group(1))
                                if 0.01 <= val <= 999999:
                                    amounts.append(val)
                if amounts:
                    return max(amounts)
        except Exception:
            pass
        return None

    @staticmethod
    def _extract_amount_from_pdf(path: Path) -> Optional[float]:
        """从 PDF 文件提取金额"""
        try:
            import pdfplumber
            with pdfplumber.open(str(path)) as pdf:
                text = ''
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + '\n'

            # 优先匹配"价税合计"后的金额
            patterns = [
                r'(?:价税合计|合计|总金额)[^\d]{0,30}([¥￥]?\s*)(\d{1,8}(?:\.\d{1,2})?)',
                r'([¥￥]\s*)(\d{1,8}(?:\.\d{1,2})?)',
            ]
            amounts = []
            for pattern in patterns:
                for m in re.finditer(pattern, text, re.I):
                    try:
                        val = float(m.group(2))
                        if 0.01 <= val <= 999999:
                            amounts.append(val)
                    except ValueError:
                        pass
            if amounts:
                return max(amounts)

            # 备用：提取所有合理的两位小数数字，取最大
            for m in re.finditer(r'(\d{1,8}\.\d{2})', text):
                try:
                    val = float(m.group(1))
                    if 0.01 <= val <= 999999:
                        amounts.append(val)
                except ValueError:
                    pass
            if amounts:
                return max(amounts)
        except Exception:
            pass
        return None

    def _extract_amount(self, path: Path) -> Optional[float]:
        """从文件提取金额（综合所有方法）"""
        # 1. 优先从文件名提取（最可靠）
        amt = self._extract_amount_from_filename(path.name)
        if amt:
            return amt

        # 2. 按文件类型提取
        ext = path.suffix.lower()
        if ext == '.ofd':
            return self._extract_amount_from_ofd(path)
        elif ext == '.pdf':
            return self._extract_amount_from_pdf(path)
        return None

    # ---- 日期提取（用于 Supplemental 目录归类） ----

    @staticmethod
    def _extract_date_from_filename(filename: str) -> Optional[datetime]:
        """从文件名提取日期"""
        patterns = [
            r'(\d{4})[-_年]?(\d{2})[-_月]?(\d{2})',  # 20260508 / 2026-05-08 / 2026年05月08日
            r'(\d{4})(\d{2})(\d{2})',  # 20260508
        ]
        for pattern in patterns:
            m = re.search(pattern, filename)
            if m:
                try:
                    year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
                    if 2020 <= year <= 2030 and 1 <= month <= 12 and 1 <= day <= 31:
                        return datetime(year, month, day)
                except (ValueError, IndexError):
                    pass
        return None

    @staticmethod
    def _extract_date_from_ofd(path: Path) -> Optional[datetime]:
        """从 OFD 文件提取开票日期"""
        try:
            import zipfile
            with zipfile.ZipFile(path, 'r') as z:
                # 方法1: CustomTag.xml 中 IssueDate
                if 'Doc_0/Tags/CustomTag.xml' in z.namelist():
                    with z.open('Doc_0/Tags/CustomTag.xml') as f:
                        content = f.read().decode('utf-8', errors='replace')
                        m = re.search(r'IssueDate[^>]*>.*?<([^<]+)<', content)
                        if m:
                            try:
                                return date_parser.parse(m.group(1))
                            except Exception:
                                pass

                # 方法2: 从 Content.xml 提取日期文本
                for name in z.namelist():
                    if name.endswith('.xml'):
                        with z.open(name) as f:
                            content = f.read().decode('utf-8', errors='replace')
                            # 匹配 "2026年05月16日"
                            m = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', content)
                            if m:
                                try:
                                    return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                                except ValueError:
                                    pass
        except Exception:
            pass
        return None

    @staticmethod
    def _extract_date_from_pdf(path: Path) -> Optional[datetime]:
        """从 PDF 文件提取开票日期"""
        try:
            import pdfplumber
            with pdfplumber.open(str(path)) as pdf:
                text = ''
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + '\n'

            # 匹配 "开票日期：2026年05月08日" 或 "2026-05-08"
            patterns = [
                r'开票日期[:：]?\s*(\d{4})年(\d{1,2})月(\d{1,2})日',
                r'开票日期[:：]?\s*(\d{4})[-/](\d{1,2})[-/](\d{1,2})',
                r'(\d{4})年(\d{1,2})月(\d{1,2})日',
            ]
            for pattern in patterns:
                m = re.search(pattern, text)
                if m:
                    try:
                        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                    except ValueError:
                        pass
        except Exception:
            pass
        return None

    def _extract_invoice_date(self, path: Path) -> Optional[datetime]:
        """从发票文件提取日期（综合所有方法）"""
        # 1. 优先从文件名提取
        dt = self._extract_date_from_filename(path.name)
        if dt:
            return dt

        # 2. 按文件类型提取
        ext = path.suffix.lower()
        if ext == '.ofd':
            return self._extract_date_from_ofd(path)
        elif ext == '.pdf':
            return self._extract_date_from_pdf(path)
        return None

    @staticmethod
    def _is_spa_empty_html(html_content: str) -> bool:
        """判断 HTML 是否为空白 SPA 页面（Vue/React 等前端框架渲染的空白页）"""
        if not html_content:
            return True
        html_lower = html_content.lower()

        # SPA 标志
        spa_markers = [
            '<div id="app"></div>',
            '<div id="root"></div>',
            '<div id=app></div>',
            '<div id=root></div>',
            "we're sorry but",
            "javascript enabled",
            "<noscript>",
        ]
        has_spa_marker = any(marker in html_lower for marker in spa_markers)

        # 内容极少（< 5KB 且没有实际的 base64 数据或 iframe）
        has_base64 = 'base64,' in html_content
        has_iframe = '<iframe' in html_lower
        has_embed = '<embed' in html_lower or '<object' in html_lower

        # 如果包含 SPA 标志，且没有 base64/iframe/embed 实际内容 → 认为是空白页
        if has_spa_marker and not (has_base64 or has_iframe or has_embed):
            return True

        # 如果 HTML 内容极少（< 2KB）且主要是 script/link 标签
        if len(html_content) < 2048:
            text_content = re.sub(r'<[^>]+>', '', html_content)
            text_content = re.sub(r'\s+', '', text_content)
            if len(text_content) < 100:
                return True

        return False

    @staticmethod
    def _extract_base64_from_html(html_content: str) -> List[bytes]:
        """从 HTML 中提取 base64 编码的数据（图片/PDF）"""
        results = []
        # data:image/xxx;base64,...
        img_pattern = re.compile(r'data:image/[^;]+;base64,([A-Za-z0-9+/=]+)')
        for m in img_pattern.finditer(html_content):
            try:
                data = base64.b64decode(m.group(1))
                if len(data) > 1024:
                    results.append(data)
            except Exception:
                pass
        # data:application/pdf;base64,...
        pdf_pattern = re.compile(r'data:application/pdf;base64,([A-Za-z0-9+/=]+)')
        for m in pdf_pattern.finditer(html_content):
            try:
                data = base64.b64decode(m.group(1))
                if len(data) > 1024:
                    results.append(data)
            except Exception:
                pass
        return results

    # ---- Playwright SPA 渲染 ----

    # 需要通过浏览器渲染才能提取下载链接的域名/URL 模式
    SPA_RENDER_PATTERNS = [
        re.compile(r'pis\.baiwang\.com/smkp-vue/previewInvoiceAllEle', re.I),
        re.compile(r'bwfp\.baiwang\.com/fp/qrcode', re.I),
        re.compile(r'webapp\.crestv\.com', re.I),
        re.compile(r'fpkj\.vpiaotong\.com', re.I),
        re.compile(r'scan\.vpiaotong\.com', re.I),
    ]

    @staticmethod
    def _needs_spa_render(url: str) -> bool:
        """判断 URL 是否需要用 Playwright 浏览器渲染"""
        return any(p.search(url) for p in InvoiceDownloader.SPA_RENDER_PATTERNS)

    def _render_spa_page(self, url: str) -> Tuple[List[str], List[bytes]]:
        """使用 Playwright 渲染 SPA 页面，提取真实的下载链接和 base64 数据

        返回 (urls, base64_datas)。
        """
        if not PLAYWRIGHT_AVAILABLE:
            self.logger.warning(f"  Playwright 未安装，无法渲染 SPA 页面: {url[:60]}...")
            return [], []

        urls = []
        base64_datas: List[bytes] = []
        try:
            self.logger.info(f"  Playwright 渲染 SPA 页面: {url[:80]}...")
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                page = browser.new_page(
                    user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
                )
                # 设置超时
                page.set_default_timeout(30000)
                page.set_default_navigation_timeout(30000)

                try:
                    page.goto(url, wait_until='networkidle', timeout=30000)
                except Exception as e:
                    self.logger.warning(f"  页面加载超时: {e}")
                    browser.close()
                    return [], []

                # 等待 Vue/React 渲染完成（额外等待 2 秒让 JS 执行）
                page.wait_for_timeout(2000)

                # 策略 0: 检测「链接过期/失效/失败」弹窗，点击确认后刷新页面重新加载
                try:
                    page_text = page.inner_text('body') or ''
                    error_keywords = [
                        '链接已过期', '链接过期', '链接失效', '已失效', '失效',
                        '下载失败', '请求失败', '出错了', '错误提示', '票据不存在',
                        '发票不存在', '无法查看', '加载失败', '请刷新',
                    ]
                    if any(kw in page_text for kw in error_keywords):
                        self.logger.info(f"  检测到错误/过期提示，尝试点击确认并刷新页面")
                        for btn_text in ['确认', '确定', '知道了', '关闭', '重试', '刷新']:
                            try:
                                locator = page.locator(
                                    f'button:has-text("{btn_text}"), '
                                    f'.el-button:has-text("{btn_text}"), '
                                    f'a:has-text("{btn_text}"), '
                                    f'div[role="button"]:has-text("{btn_text}"), '
                                    f'span:has-text("{btn_text}")'
                                ).first
                                if locator.is_visible():
                                    locator.click()
                                    self.logger.info(f"  点击「{btn_text}」按钮")
                                    page.wait_for_timeout(1500)
                                    break
                            except Exception:
                                pass
                        try:
                            page.reload(wait_until='networkidle', timeout=30000)
                            page.wait_for_timeout(3000)
                            self.logger.info(f"  页面已刷新，重新提取内容")
                        except Exception as e:
                            self.logger.warning(f"  刷新页面失败: {e}")
                except Exception:
                    pass

                # 策略 1: 查找页面中所有的 <a> 链接，按关键词筛选
                hrefs = page.eval_on_selector_all('a[href]', 'elements => elements.map(e => ({href: e.href, text: e.innerText.trim()}))')
                for link in hrefs:
                    href = link.get('href', '')
                    text = link.get('text', '')
                    if href and any(k in href.lower() for k in ['pdf', 'ofd', 'png', 'jpg', 'jpeg', 'download', 'invoice', 'fapiao']):
                        urls.append(href)
                    elif href and any(k in text.lower() for k in ['下载', 'pdf', '发票', 'invoice', 'download']):
                        urls.append(href)

                # 策略 2: 查找页面中所有的 <img> src
                imgs = page.eval_on_selector_all('img[src]', 'elements => elements.map(e => e.src)')
                for src in imgs:
                    if src and any(src.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.webp']):
                        # 排除极小的图片（图标/logo）
                        try:
                            size_info = page.eval_on_selector(f'img[src={json.dumps(src)}]', 'e => e.naturalWidth')
                            if size_info and size_info > 200:  # 大于 200px 的图片才可能是发票
                                urls.append(src)
                        except Exception:
                            urls.append(src)

                # 策略 3: 查找页面中所有的 iframe src
                iframes = page.eval_on_selector_all('iframe[src]', 'elements => elements.map(e => e.src)')
                for src in iframes:
                    if src:
                        urls.append(src)

                # 策略 4: 查找页面中 base64 编码的 PDF/图片
                page_content = page.content()
                base64_datas = self._extract_base64_from_html(page_content)

                # 策略 5: 百望云特殊处理 - 查找 Vue 组件中的下载按钮
                try:
                    # 百望云常见下载按钮 class
                    download_btns = page.query_selector_all('button, a, div')
                    for btn in download_btns:
                        btn_text = btn.inner_text()
                        if btn_text and any(k in btn_text.lower() for k in ['下载', 'pdf', 'ofd', '保存', 'download']):
                            # 尝试获取按钮关联的点击事件或 href
                            btn_html = btn.inner_html()
                            hrefs_in_btn = re.findall(r'href=["\']([^"\']+)["\']', btn_html)
                            for h in hrefs_in_btn:
                                if h.startswith('http'):
                                    urls.append(h)
                except Exception:
                    pass

                browser.close()

            # 去重 + 过滤
            seen = set()
            unique = []
            for u in urls:
                if u in seen:
                    continue
                # 跳过黑名单 URL
                if any(p.search(u) for p in self.parser.URL_BLACKLIST_PATTERNS):
                    continue
                # 使用现有打分逻辑过滤低质量链接
                score = self.parser._score_link(u, '')
                if score <= 0:
                    continue
                seen.add(u)
                unique.append(u)

            self.logger.info(f"  Playwright 提取到 {len(unique)} 个有效链接, {len(base64_datas)} 个 base64 数据")
            return unique, base64_datas

        except Exception as e:
            self.logger.warning(f"  Playwright 渲染失败: {e}")
            return [], []

    # ---- 附件保存（带查重） ----

    def _save_attachment(self, att: Dict[str, Any], month_dir: Path, subject: str) -> Optional[Dict[str, Any]]:
        """保存附件到指定目录，带内容查重和发票级去重"""
        filename = att['filename']
        content = att['content']
        content_hash = self._compute_hash(content)

        # 1. 内容哈希查重
        if self.db.hash_exists(content_hash):
            self.logger.info(f"  附件内容已存在，跳过: {filename}")
            return {
                'type': '附件',
                'filename': filename,
                'path': '',
                'size': att['size'],
                'status': '已存在/重复',
            }

        safe_name = self._sanitize_filename(filename)
        dest = self._unique_path(month_dir / safe_name)

        if self.dry_run:
            self.logger.info(f"  [预览] 将保存附件: {safe_name} ({att['size']} bytes)")
            return {
                'type': '附件',
                'filename': safe_name,
                'path': str(dest.relative_to(self.target_dir)),
                'size': att['size'],
                'status': '预览',
            }

        try:
            with open(dest, 'wb') as f:
                f.write(content)
        except Exception as e:
            self.logger.error(f"  保存附件失败: {e}")
            return None

        # 尝试从文件内容中提取发票号码并重命名
        dest = self._try_rename_with_invoice_id(dest)

        # 2. 发票级跨格式去重
        invoice_id = self._extract_invoice_id(dest.name)
        file_format = Path(dest.name).suffix
        file_size = dest.stat().st_size

        if not self.db.record_file(
            invoice_id=invoice_id,
            content_hash=content_hash,
            file_path=str(dest),
            file_format=file_format,
            file_size=file_size,
            source_email=subject,
            source_url='',
        ):
            self.logger.info(f"  发票 {invoice_id or 'N/A'} 已有更高优先级格式，跳过: {dest.name}")
            return {
                'type': '附件',
                'filename': dest.name,
                'path': str(dest.relative_to(self.target_dir)),
                'size': file_size,
                'status': '已存在/重复',
            }

        amount = self._extract_amount(dest)
        self.logger.success(f"  保存附件: {dest.name} ({att['size']} bytes){f' 金额:¥{amount:.2f}' if amount else ''}")
        return {
            'type': '附件',
            'filename': dest.name,
            'path': str(dest.relative_to(self.target_dir)),
            'size': att['size'],
            'status': '成功',
            'amount': amount,
        }

    # ---- 链接下载（带查重 + SPA 检测） ----

    def _download_link(self, link: Dict[str, Any], month_dir: Path, subject: str) -> Optional[Dict[str, Any]]:
        """下载链接指向的文件，先下到临时文件，验证后再保存"""
        url = link['url']
        self.logger.info(f"  发现链接 [{link['score']}分]: {url[:80]}...")

        resolved = self.parser.resolve_short_link(url)
        if resolved and resolved != url:
            self.logger.info(f"  短链接解析为: {resolved[:80]}...")
            url = resolved

        # URL 级别查重：该 URL 已下载过则直接跳过
        if self.db.url_exists(url):
            self.logger.info(f"  URL 已下载过，跳过: {url[:60]}...")
            return {
                'type': '链接下载',
                'filename': '',
                'path': '',
                'size': 0,
                'source_url': url,
                'status': '已存在/重复(URL)',
            }

        # 生成临时文件名
        try:
            resp = self.downloader.session.head(url, allow_redirects=True, timeout=30)
            content_type = resp.headers.get('content-type', '')
            content_length = resp.headers.get('content-length')
        except Exception:
            content_type = ''
            content_length = None

        filename = self.parser.guess_filename_from_url(url, content_type)
        safe_name = self._sanitize_filename(filename)

        if not self.parser._looks_like_invoice_name(safe_name):
            prefix = self._extract_prefix_from_subject(subject)
            name_part = Path(safe_name).stem
            ext = Path(safe_name).suffix
            safe_name = f"{prefix}_{name_part}{ext}"

        if self.dry_run:
            self.logger.info(f"  [预览] 将下载链接: {safe_name} ({url[:80]}...)")
            return {
                'type': '链接下载',
                'filename': safe_name,
                'path': '',
                'size': self._safe_int(content_length, 0),
                'source_url': url,
                'status': '预览',
            }

        # 先下载到临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=Path(safe_name).suffix) as tmp:
            tmp_path = Path(tmp.name)

        expected_size = self._safe_int(content_length)
        success = self.downloader.download(url, tmp_path, expected_size)

        if not success:
            tmp_path.unlink(missing_ok=True)
            self.logger.error(f"  下载失败: {url[:80]}")
            return None

        # 修正扩展名
        detected_ext = self._detect_extension(tmp_path, url)
        if detected_ext and tmp_path.suffix.lower() != detected_ext:
            new_tmp = tmp_path.with_suffix(detected_ext)
            tmp_path.rename(new_tmp)
            tmp_path = new_tmp
            safe_name = Path(safe_name).stem + detected_ext

        # ---- SPA / 阅读器页面检测 ----
        html_processed = False
        if tmp_path.suffix.lower() == '.html' or detected_ext == '.html':
            try:
                html_content = tmp_path.read_text(encoding='utf-8', errors='replace')

                # 策略1: 如果是已知需要浏览器渲染的页面（百望云预览、票通等），
                # 无论 HTML 是否"空白"，都用 Playwright 尝试提取真实下载链接
                if self._needs_spa_render(url):
                    self.logger.info(f"  检测到 SPA 页面，尝试 Playwright 渲染: {url[:60]}...")
                    spa_urls, spa_base64 = self._render_spa_page(url)
                    if spa_urls:
                        # 删除临时 HTML 文件，尝试下载 SPA 提取到的真实链接
                        tmp_path.unlink(missing_ok=True)
                        for spa_url in spa_urls:
                            spa_file_info = self._download_direct_url(spa_url, month_dir, subject, source_url=url)
                            if spa_file_info:
                                return spa_file_info
                        self.logger.warning(f"  SPA 提取的 {len(spa_urls)} 个链接均下载失败")
                        # 记录需要手动下载的链接
                        self.stats['links_manual'] = self.stats.get('links_manual', 0) + 1
                        return {
                            'type': '需手动下载',
                            'filename': safe_name,
                            'path': '',
                            'size': 0,
                            'source_url': url,
                            'status': 'SPA链接过期/失效，需手动从邮箱网页版下载',
                            'needs_manual_download': True,
                        }
                    if spa_base64:
                        # 从 SPA 渲染后的页面中提取 base64 数据
                        best_data = max(spa_base64, key=len)
                        tmp_path.write_bytes(best_data)
                        self.logger.info(f"  从 SPA 页面提取 base64 数据 ({len(best_data)} bytes)")
                        detected_ext = self._detect_extension(tmp_path, url)
                        if detected_ext and tmp_path.suffix.lower() != detected_ext:
                            new_tmp = tmp_path.with_suffix(detected_ext)
                            tmp_path.rename(new_tmp)
                            tmp_path = new_tmp
                            safe_name = Path(safe_name).stem + detected_ext
                        html_processed = True
                    if not html_processed:
                        self.logger.info(f"  Playwright 未提取到有效链接，跳过")
                        tmp_path.unlink(missing_ok=True)
                        # 记录需要手动下载的链接
                        self.stats['links_manual'] = self.stats.get('links_manual', 0) + 1
                        return {
                            'type': '需手动下载',
                            'filename': safe_name,
                            'path': '',
                            'size': 0,
                            'source_url': url,
                            'status': 'SPA链接过期/失效，需手动从邮箱网页版下载',
                            'needs_manual_download': True,
                        }

                # 策略2: 如果是空白 SPA 页（Vue/React 等前端框架渲染的空白页）
                if not html_processed and self._is_spa_empty_html(html_content):
                    # 先尝试提取 base64 内容
                    base64_datas = self._extract_base64_from_html(html_content)
                    if base64_datas:
                        best_data = max(base64_datas, key=len)
                        tmp_path.write_bytes(best_data)
                        self.logger.info(f"  从 HTML 中提取 base64 数据 ({len(best_data)} bytes)")
                        detected_ext = self._detect_extension(tmp_path, url)
                        if detected_ext and tmp_path.suffix.lower() != detected_ext:
                            new_tmp = tmp_path.with_suffix(detected_ext)
                            tmp_path.rename(new_tmp)
                            tmp_path = new_tmp
                            safe_name = Path(safe_name).stem + detected_ext
                    else:
                        self.logger.info(f"  检测到 JS 渲染空白页，跳过: {safe_name}")
                        tmp_path.unlink(missing_ok=True)
                        return {
                            'type': '链接下载',
                            'filename': safe_name,
                            'path': '',
                            'size': 0,
                            'source_url': url,
                            'status': 'JS渲染空白页，跳过',
                        }

                # 策略3: 检测无实际发票数据的 HTML 阅读器/框架页面
                if not html_processed:
                    has_base64 = len(self._extract_base64_from_html(html_content)) > 0
                    if not has_base64:
                        html_lower = html_content.lower()
                        reader_markers = [
                            '绿页云', 'greenpaper', 'pdfviewer', 'viewercontainer',
                            '在线阅读', '阅读器', 'webviewer', '发票查看', '电子发票查看',
                            '在线预览', '发票预览', '发票阅读', 'greenpaper',
                        ]
                        is_reader_page = any(marker in html_lower for marker in reader_markers)
                        # 或者页面内容极少（主要是脚本和样式，没有实际文本内容）
                        text_only = re.sub(r'<[^>]+>', '', html_content)
                        text_only = re.sub(r'\s+', '', text_only)
                        is_mostly_scripts = len(text_only) < 200 and len(html_content) > 3000

                        if is_reader_page or is_mostly_scripts:
                            self.logger.info(f"  检测到在线阅读器/框架页面，无实际发票数据，跳过: {safe_name}")
                            tmp_path.unlink(missing_ok=True)
                            return {
                                'type': '链接下载',
                                'filename': safe_name,
                                'path': '',
                                'size': 0,
                                'source_url': url,
                                'status': '在线阅读器页面，跳过',
                            }
            except Exception:
                pass

        # ---- 内容查重 + 发票级去重 ----
        content_hash = self._compute_file_hash(tmp_path)

        if self.db.hash_exists(content_hash):
            self.logger.info(f"  链接内容已存在，跳过: {safe_name}")
            tmp_path.unlink(missing_ok=True)
            return {
                'type': '链接下载',
                'filename': safe_name,
                'path': '',
                'size': 0,
                'source_url': url,
                'status': '已存在/重复',
            }

        # 移动到最终位置
        dest = self._unique_path(month_dir / safe_name)
        try:
            tmp_path.rename(dest)
        except Exception:
            import shutil
            shutil.move(str(tmp_path), str(dest))

        # 尝试从文件内容中提取发票号码并重命名
        dest = self._try_rename_with_invoice_id(dest)

        file_size = dest.stat().st_size
        invoice_id = self._extract_invoice_id(dest.name)
        file_format = dest.suffix

        if not self.db.record_file(
            invoice_id=invoice_id,
            content_hash=content_hash,
            file_path=str(dest),
            file_format=file_format,
            file_size=file_size,
            source_email=subject,
            source_url=url,
        ):
            self.logger.info(f"  发票 {invoice_id or 'N/A'} 已有更高优先级格式，跳过: {dest.name}")
            return {
                'type': '链接下载',
                'filename': dest.name,
                'path': str(dest.relative_to(self.target_dir)),
                'size': file_size,
                'source_url': url,
                'status': '已存在/重复',
            }

        amount = self._extract_amount(dest)
        self.logger.success(f"  下载完成: {dest.name}{f' 金额:¥{amount:.2f}' if amount else ''}")
        return {
            'type': '链接下载',
            'filename': dest.name,
            'path': str(dest.relative_to(self.target_dir)),
            'size': file_size,
            'source_url': url,
            'status': '成功',
            'amount': amount,
        }

    def _download_direct_url(self, url: str, month_dir: Path, subject: str, source_url: str = '') -> Optional[Dict[str, Any]]:
        """直接下载 URL，不经过 SPA 检测（用于 Playwright 提取到的真实链接）"""
        self.logger.info(f"  SPA提取链接下载: {url[:80]}...")

        # 检查 URL 黑名单
        for pattern in self.parser.URL_BLACKLIST_PATTERNS:
            if pattern.search(url):
                self.logger.info(f"  SPA提取链接在黑名单中，跳过: {url[:60]}...")
                return None

        # URL 级别查重
        if self.db.url_exists(url):
            self.logger.info(f"  SPA提取链接 URL 已下载过，跳过: {url[:60]}...")
            return None

        try:
            resp = self.downloader.session.head(url, allow_redirects=True, timeout=30)
            content_type = resp.headers.get('content-type', '')
            content_length = resp.headers.get('content-length')
        except Exception:
            content_type = ''
            content_length = None

        filename = self.parser.guess_filename_from_url(url, content_type)
        safe_name = self._sanitize_filename(filename)

        if not self.parser._looks_like_invoice_name(safe_name):
            prefix = self._extract_prefix_from_subject(subject)
            name_part = Path(safe_name).stem
            ext = Path(safe_name).suffix
            safe_name = f"{prefix}_{name_part}{ext}"

        if self.dry_run:
            self.logger.info(f"  [预览] 将下载SPA链接: {safe_name} ({url[:80]}...)")
            return {
                'type': '链接下载',
                'filename': safe_name,
                'path': '',
                'size': self._safe_int(content_length, 0),
                'source_url': source_url or url,
                'status': '预览',
            }

        with tempfile.NamedTemporaryFile(delete=False, suffix=Path(safe_name).suffix) as tmp:
            tmp_path = Path(tmp.name)

        expected_size = self._safe_int(content_length)
        success = self.downloader.download(url, tmp_path, expected_size)

        if not success:
            tmp_path.unlink(missing_ok=True)
            return None

        detected_ext = self._detect_extension(tmp_path, url)
        if detected_ext and tmp_path.suffix.lower() != detected_ext:
            new_tmp = tmp_path.with_suffix(detected_ext)
            tmp_path.rename(new_tmp)
            tmp_path = new_tmp
            safe_name = Path(safe_name).stem + detected_ext

        # 如果下载到的是 HTML，检查是否是 SPA/阅读器页面（避免下载绿页云等阅读器 iframe）
        if tmp_path.suffix.lower() == '.html' or detected_ext == '.html':
            try:
                html_content = tmp_path.read_text(encoding='utf-8', errors='replace')
                has_base64 = len(self._extract_base64_from_html(html_content)) > 0

                # 1. 需要浏览器渲染的已知 SPA 页面
                if self._needs_spa_render(url):
                    self.logger.info(f"  SPA提取链接仍需渲染，跳过: {url[:60]}...")
                    tmp_path.unlink(missing_ok=True)
                    return None

                # 2. 空白 SPA 页
                if self._is_spa_empty_html(html_content):
                    base64_datas = self._extract_base64_from_html(html_content)
                    if base64_datas:
                        best_data = max(base64_datas, key=len)
                        tmp_path.write_bytes(best_data)
                        detected_ext = self._detect_extension(tmp_path, url)
                        if detected_ext and tmp_path.suffix.lower() != detected_ext:
                            new_tmp = tmp_path.with_suffix(detected_ext)
                            tmp_path.rename(new_tmp)
                            tmp_path = new_tmp
                    else:
                        self.logger.info(f"  SPA提取链接返回空白页，跳过: {url[:60]}...")
                        tmp_path.unlink(missing_ok=True)
                        return None

                # 3. 没有 base64 实际数据的 HTML 阅读器页面
                if not has_base64:
                    html_lower = html_content.lower()
                    reader_markers = [
                        '绿页云', 'greenpaper', 'pdfviewer', 'viewercontainer',
                        '在线阅读', '阅读器', 'webviewer', '发票查看', '电子发票查看',
                        '在线预览', '发票预览', '发票阅读',
                    ]
                    is_reader_page = any(marker in html_lower for marker in reader_markers)
                    text_only = re.sub(r'<[^>]+>', '', html_content)
                    text_only = re.sub(r'\s+', '', text_only)
                    is_mostly_scripts = len(text_only) < 200 and len(html_content) > 3000

                    if is_reader_page or is_mostly_scripts:
                        self.logger.info(f"  SPA提取链接返回 HTML 阅读器，跳过: {url[:60]}...")
                        tmp_path.unlink(missing_ok=True)
                        return None
            except Exception:
                pass

        content_hash = self._compute_file_hash(tmp_path)

        if self.db.hash_exists(content_hash):
            tmp_path.unlink(missing_ok=True)
            return None

        dest = self._unique_path(month_dir / safe_name)
        try:
            tmp_path.rename(dest)
        except Exception:
            import shutil
            shutil.move(str(tmp_path), str(dest))

        # 尝试从文件内容中提取发票号码并重命名
        dest = self._try_rename_with_invoice_id(dest)

        file_size = dest.stat().st_size
        invoice_id = self._extract_invoice_id(dest.name)
        file_format = dest.suffix

        if not self.db.record_file(
            invoice_id=invoice_id,
            content_hash=content_hash,
            file_path=str(dest),
            file_format=file_format,
            file_size=file_size,
            source_email=subject,
            source_url=source_url or url,
        ):
            return None

        amount = self._extract_amount(dest)
        self.logger.success(f"  SPA提取下载完成: {dest.name}{f' 金额:¥{amount:.2f}' if amount else ''}")
        return {
            'type': '链接下载',
            'filename': dest.name,
            'path': str(dest.relative_to(self.target_dir)),
            'size': file_size,
            'source_url': source_url or url,
            'status': '成功',
            'amount': amount,
        }

    @staticmethod
    def _detect_extension(path: Path, original_url: str = '') -> Optional[str]:
        """根据文件内容头检测扩展名"""
        if not path.exists() or path.stat().st_size == 0:
            return None

        try:
            with open(path, 'rb') as f:
                header = f.read(16)
        except Exception:
            return None

        if len(header) < 4:
            return None

        ext_map = {
            b'%PDF': '.pdf',
            b'\x89PNG': '.png',
            b'\xff\xd8\xff': '.jpg',
            b'GIF87a': '.gif',
            b'GIF89a': '.gif',
            b'PK': '.zip',
        }

        detected = None
        for magic, ext in ext_map.items():
            if header.startswith(magic):
                detected = ext
                break

        if header[:4] == b'RIFF' and header[8:12] == b'WEBP':
            detected = '.webp'

        text_header = header.lower()
        if b'<!doctype' in text_header or b'<html' in text_header or b'<!DOCTYPE' in header:
            detected = '.html'

        if detected == '.zip' and 'ofd' in original_url.lower():
            detected = '.ofd'

        return detected

    @staticmethod
    def _get_link_priority(url: str) -> int:
        url_lower = url.lower()
        if url_lower.endswith(('.pdf', '_pdf')):
            return 4
        if any(url_lower.endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.webp',
                                                      '_png', '_jpg', '_jpeg', '_gif', '_webp']):
            return 3
        if url_lower.endswith(('.ofd', '_ofd')):
            return 2
        if url_lower.endswith(('.xml', '_xml')):
            return 1
        return 0

    @staticmethod
    def _get_link_base_id(url: str) -> str:
        parsed = urlparse(url)
        query = parsed.query
        path = parsed.path

        if 'baiwang.com' in parsed.netloc.lower() and 'param=' in query:
            m = re.search(r'param=([^&]+)', query)
            if m:
                return f"bw:{m.group(1)[:32]}"

        if 'chinatax.gov.cn' in parsed.netloc.lower() and 'Fphm=' in query:
            m = re.search(r'Fphm=([^&]+)', query)
            if m:
                return f"tax:{m.group(1)}"

        if 'meituan.net' in parsed.netloc.lower():
            name = path.split('/')[-1] if '/' in path else path
            name = re.sub(r'[_\.](png|jpg|jpeg|gif|webp)$', '', name, flags=re.I)
            return f"mt:{name[:16]}"

        if 'vpiaotong.com' in parsed.netloc.lower():
            name = path.split('/')[-1] if '/' in path else path
            return f"vt:{name[:20]}"

        if parsed.netloc.lower() in {'wxaurl.cn', 't.cn'}:
            return f"short:{url}"

        return re.sub(r'[_\.](pdf|ofd|xml|png|jpg|jpeg|gif|webp)$', '', url, flags=re.I)

    SHORT_LINK_DOMAINS = {'wxaurl.cn', 't.cn', 'bit.ly', 'tinyurl.com', 'goo.gl', 'dwz.cn'}

    def _group_links_by_invoice(self, links: List[Dict[str, Any]]) -> List[List[Dict[str, Any]]]:
        groups: Dict[str, List[Dict[str, Any]]] = {}
        unresolved_short_links: List[Dict[str, Any]] = []

        for link in links:
            url = link['url']
            resolved = self.parser.resolve_short_link(url)
            if resolved:
                url = resolved
            base_id = self._get_link_base_id(url)

            is_short = any(domain in link['url'].lower() for domain in self.SHORT_LINK_DOMAINS)
            if is_short and not resolved:
                unresolved_short_links.append(link)
            else:
                groups.setdefault(base_id, []).append(link)

        if not groups and unresolved_short_links:
            for link in unresolved_short_links:
                base_id = self._get_link_base_id(link['url'])
                groups.setdefault(base_id, []).append(link)

        return list(groups.values())

    def _download_best_link(self, group: List[Dict[str, Any]], month_dir: Path, subject: str) -> Optional[Dict[str, Any]]:
        group.sort(key=lambda x: self._get_link_priority(x['url']), reverse=True)

        for link in group:
            file_info = self._download_link(link, month_dir, subject)
            if file_info:
                return file_info

        base_id = self._get_link_base_id(group[0]['url'])
        self.logger.error(f"  该发票所有格式均下载失败: {base_id}")
        return None

    @staticmethod
    def _sanitize_filename(name: str) -> str:
        name = re.sub(r'[\\/:*?"<>|]', '_', name)
        name = re.sub(r'[\x00-\x1f\x7f]', '', name)
        if len(name) > 200:
            stem = Path(name).stem[:100]
            suffix = Path(name).suffix
            name = f"{stem}{suffix}"
        return name.strip()

    @staticmethod
    def _unique_path(path: Path) -> Path:
        if not path.exists():
            return path

        stem = path.stem
        suffix = path.suffix
        parent = path.parent
        counter = 1

        while True:
            new_name = f"{stem}_{counter}{suffix}"
            new_path = parent / new_name
            if not new_path.exists():
                return new_path
            counter += 1

    @staticmethod
    def _extract_prefix_from_subject(subject: str) -> str:
        cleaned = re.sub(r'[【\[\(].*?[\]\)】]', '', subject)
        cleaned = cleaned.strip()
        prefix = cleaned[:20].strip()
        prefix = re.sub(r'[^\w一-鿿]', '_', prefix)
        return prefix or 'invoice'

    @staticmethod
    def _safe_int(value, default=None):
        try:
            return int(value) if value is not None else default
        except (ValueError, TypeError):
            return default

    def _print_stats(self):
        self.logger.info("=" * 50)
        self.logger.info("处理完成！统计信息：")
        self.logger.info(f"  检索邮件数: {self.stats['emails_checked']}")
        self.logger.info(f"  发现附件:   {self.stats['attachments_found']}")
        self.logger.info(f"  保存附件:   {self.stats['attachments_saved']}")
        self.logger.info(f"  跳过附件:   {self.stats['attachments_skipped']}")
        self.logger.info(f"  发现链接:   {self.stats['links_found']}")
        self.logger.info(f"  链接下载成功: {self.stats['links_downloaded']}")
        self.logger.info(f"  链接跳过(重复): {self.stats['links_skipped_dup']}")
        self.logger.info(f"  链接跳过(JS空白): {self.stats['links_skipped_spa']}")
        self.logger.info(f"  链接下载失败: {self.stats['links_failed']}")
        manual = self.stats.get('links_manual', 0)
        if manual > 0:
            self.logger.info(f"  ⚠️  需手动下载: {manual} 个（SPA链接过期，请登录邮箱网页版下载）")
        self.logger.info(f"  保存目录:   {self.target_dir.absolute()}")
        self.logger.info("=" * 50)
        if manual > 0:
            self.logger.info("提示：部分发票因SPA预览链接过期无法自动下载，")
            self.logger.info("      请登录 163邮箱网页版查看邮件并手动下载。")
            self.logger.info("      将下载好的发票放入 Supplemental 目录，")
            self.logger.info("      下次运行脚本时会自动归类到对应月份。")
            self.logger.info("=" * 50)


def load_config(path: str) -> Dict[str, Any]:
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def main():
    parser = argparse.ArgumentParser(description='发票邮件附件及链接下载工具')
    parser.add_argument('--config', '-c', default='config.json', help='配置文件路径')
    parser.add_argument('--dry-run', action='store_true', help='仅预览，不下载')
    args = parser.parse_args()

    if not os.path.exists(args.config):
        print(f"配置文件不存在: {args.config}")
        print("请复制 config.json 并填写你的邮箱信息")
        sys.exit(1)

    config = load_config(args.config)

    if args.dry_run:
        print("【仅预览模式】不会实际下载文件")
        config['dry_run'] = True

    app = InvoiceDownloader(config)
    app.run()


if __name__ == '__main__':
    main()
