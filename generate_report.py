#!/usr/bin/env python3
"""
发票报告生成器

功能：
1. 扫描 invoices 目录下的所有发票文件（PDF、图片、OFD、XML 等）
2. 解析发票信息：从文件名和内容（OFD、PDF）中提取金额、开票日期、发票号码
3. 处理 Supplemental 目录：将手动补充的发票按日期归类到对应月份，自动去重
4. 生成 JSON 和 HTML 核查报告，含月度金额汇总和按发票汇总
5. 重建 SQLite 数据库（内容哈希、发票最佳格式、来源记录）
6. 自动根据发票号码重命名文件

既可以独立运行，也可以被 download_invoices.py 导入调用。

独立使用方法：
    python generate_report.py                    # 扫描 ./invoices，生成报告
    python generate_report.py --dir ./invoices   # 扫描指定目录
    python generate_report.py --supplement       # 同时处理 Supplemental 目录
    python generate_report.py --rebuild-db       # 重建数据库
    python generate_report.py --dry-run          # 仅预览，不移动文件
"""

import argparse
import hashlib
import html as html_module
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

try:
    from dateutil import parser as date_parser
    DATEUTIL_AVAILABLE = True
except ImportError:
    DATEUTIL_AVAILABLE = False


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

    FORMAT_PRIORITY = {
        '.pdf': 4, '.png': 3, '.jpg': 3, '.jpeg': 3,
        '.gif': 3, '.webp': 3, '.ofd': 2, '.xml': 1,
        '.html': 0, '.bin': 0, '.tmp': 0, '': 0,
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
        if not source_url:
            return False
        cur = self.conn.execute(
            'SELECT 1 FROM downloaded_files WHERE source_url = ? LIMIT 1',
            (source_url,)
        )
        return cur.fetchone() is not None

    def get_best_for_invoice(self, invoice_id: str) -> Optional[Dict[str, Any]]:
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
        priority = self.get_priority(file_format)
        abs_path = str(Path(file_path).absolute())

        if self.hash_exists(content_hash):
            self._remove_file(abs_path)
            return False

        if invoice_id:
            best = self.get_best_for_invoice(invoice_id)
            if best:
                if best['best_priority'] >= priority:
                    self._remove_file(abs_path)
                    return False
                else:
                    self._remove_file(best['file_path'])

        try:
            self.conn.execute(
                '''INSERT INTO downloaded_files
                   (invoice_id, content_hash, file_path, file_format, format_priority, file_size, source_email, source_url)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                (invoice_id, content_hash, abs_path, file_format, priority, file_size, source_email, source_url)
            )
        except sqlite3.IntegrityError:
            self._remove_file(abs_path)
            return False

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
            Path(file_path).unlink(missing_ok=True)
        except Exception:
            pass

    def close(self):
        self.conn.close()


# ---- 发票文件解析器 ----

class InvoiceParser:
    """从文件名和内容提取发票信息"""

    INVOICE_EXTENSIONS = {'.pdf', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.ofd', '.xml'}

    @staticmethod
    def compute_hash(data: bytes) -> str:
        return hashlib.sha256(data).hexdigest()

    @staticmethod
    def compute_file_hash(path: Path) -> str:
        sha = hashlib.sha256()
        with open(path, 'rb') as f:
            for chunk in iter(lambda: f.read(65536), b''):
                sha.update(chunk)
        return sha.hexdigest()

    @staticmethod
    def extract_invoice_id_from_filename(filename: str) -> Optional[str]:
        m = re.search(r'\b(\d{20})\b', filename)
        if m:
            return m.group(1)
        m = re.search(r'^(\d{12,20})', Path(filename).stem)
        if m:
            return m.group(1)
        if '发票金额' in filename:
            merchant_match = re.search(r'^([^_]+)', filename)
            if merchant_match:
                merchant = merchant_match.group(1)
                amt_match = re.search(r'发票金额\D*(\d+)', filename)
                if amt_match:
                    return f"mt:{merchant}_{amt_match.group(1)}"
                return f"mt:{merchant}"
        stem = Path(filename).stem
        if len(stem) > 5:
            return stem[:30]
        return None

    @staticmethod
    def extract_invoice_id_from_file(path: Path) -> Optional[str]:
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
                # 1) 先用 strings 快速尝试
                try:
                    result = subprocess.run(['strings', str(path)], capture_output=True, text=True, timeout=5)
                    m = re.search(r'\b(\d{20})\b', result.stdout)
                    if m:
                        return m.group(1)
                except Exception:
                    pass

                # 2) 回退到 pdfplumber，能处理更多编码/排版
                try:
                    import pdfplumber
                    with pdfplumber.open(str(path)) as pdf:
                        text = ''
                        for page in pdf.pages:
                            page_text = page.extract_text()
                            if page_text:
                                text += page_text + '\n'
                    # 先尝试连续 20 位
                    m = re.search(r'\b(\d{20})\b', text)
                    if m:
                        return m.group(1)
                    # 再尝试带空格的 20 位（某些竖排/分栏发票）
                    m = re.search(r'(\d(?:\s*\d){19})', text)
                    if m:
                        cleaned = re.sub(r'\s', '', m.group(1))
                        if len(cleaned) == 20 and cleaned.isdigit():
                            return cleaned
                except Exception:
                    pass
        except Exception:
            pass
        return None

    @staticmethod
    def try_rename_with_invoice_id(path: Path, logger: Optional[Logger] = None) -> Path:
        invoice_id = InvoiceParser.extract_invoice_id_from_file(path)
        if invoice_id and invoice_id not in path.name:
            new_name = f"{invoice_id}{path.suffix}"
            new_path = InvoiceParser.unique_path(path.parent / InvoiceParser.sanitize_filename(new_name))
            try:
                path.rename(new_path)
                if logger:
                    logger.info(f"  根据发票号码重命名: {path.name} -> {new_path.name}")
                return new_path
            except Exception:
                pass
        return path

    @staticmethod
    def extract_amount_from_filename(filename: str) -> Optional[float]:
        m = re.search(r'(?:发票)?金额\D*(\d+(?:\.\d+)?)', filename)
        if m:
            try:
                return float(m.group(1))
            except ValueError:
                pass
        return None

    @staticmethod
    def extract_amount_from_ofd(path: Path) -> Optional[float]:
        try:
            import zipfile
            with zipfile.ZipFile(path, 'r') as z:
                if 'Doc_0/Tags/CustomTag.xml' in z.namelist():
                    with z.open('Doc_0/Tags/CustomTag.xml') as f:
                        content = f.read().decode('utf-8', errors='replace')
                        m = re.search(r'TaxInclusiveTotalAmount[^>]*>(\d+\.\d{2})<', content)
                        if m:
                            return float(m.group(1))

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
    def extract_amount_from_pdf(path: Path) -> Optional[float]:
        try:
            import pdfplumber
            with pdfplumber.open(str(path)) as pdf:
                text = ''
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + '\n'

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

    @staticmethod
    def extract_amount(path: Path) -> Optional[float]:
        amt = InvoiceParser.extract_amount_from_filename(path.name)
        if amt:
            return amt

        ext = path.suffix.lower()
        if ext == '.ofd':
            return InvoiceParser.extract_amount_from_ofd(path)
        elif ext == '.pdf':
            return InvoiceParser.extract_amount_from_pdf(path)
        return None

    @staticmethod
    def extract_date_from_filename(filename: str) -> Optional[datetime]:
        patterns = [
            r'(\d{4})[-_年]?(\d{2})[-_月]?(\d{2})',
            r'(\d{4})(\d{2})(\d{2})',
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
    def extract_date_from_ofd(path: Path) -> Optional[datetime]:
        try:
            import zipfile
            with zipfile.ZipFile(path, 'r') as z:
                if 'Doc_0/Tags/CustomTag.xml' in z.namelist():
                    with z.open('Doc_0/Tags/CustomTag.xml') as f:
                        content = f.read().decode('utf-8', errors='replace')
                        m = re.search(r'IssueDate[^>]*>.*?<([^<]+)<', content)
                        if m:
                            try:
                                if DATEUTIL_AVAILABLE:
                                    return date_parser.parse(m.group(1))
                            except Exception:
                                pass

                for name in z.namelist():
                    if name.endswith('.xml'):
                        with z.open(name) as f:
                            content = f.read().decode('utf-8', errors='replace')
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
    def extract_date_from_pdf(path: Path) -> Optional[datetime]:
        try:
            import pdfplumber
            with pdfplumber.open(str(path)) as pdf:
                text = ''
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + '\n'

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

    @staticmethod
    def extract_date(path: Path) -> Optional[datetime]:
        dt = InvoiceParser.extract_date_from_filename(path.name)
        if dt:
            return dt

        ext = path.suffix.lower()
        if ext == '.ofd':
            return InvoiceParser.extract_date_from_ofd(path)
        elif ext == '.pdf':
            return InvoiceParser.extract_date_from_pdf(path)
        return None

    @staticmethod
    def extract_seller_from_ofd(path: Path) -> Optional[str]:
        try:
            import zipfile
            with zipfile.ZipFile(path, 'r') as z:
                # 1) 尝试从 CustomTag.xml 的 SellerName ObjectRef 定位到页面 TextObject
                seller_ref_id: Optional[str] = None
                if 'Doc_0/Tags/CustomTag.xml' in z.namelist():
                    with z.open('Doc_0/Tags/CustomTag.xml') as f:
                        tag_content = f.read().decode('utf-8', errors='replace')
                    m = re.search(
                        r'<[^>]*SellerName[^>]*>.*?<[^>]*ObjectRef[^>]*>(\d+)</[^>]*ObjectRef>.*?</[^>]*SellerName>',
                        tag_content,
                        re.S,
                    )
                    if m:
                        seller_ref_id = m.group(1)

                if seller_ref_id:
                    for name in z.namelist():
                        if 'Content' in name and name.endswith('.xml'):
                            with z.open(name) as f:
                                content = f.read().decode('utf-8', errors='replace')
                            # 查找对应 ID 的 TextObject，然后取其中的 TextCode 文本
                            obj_pat = (
                                r'<[^>]*TextObject\s+[^>]*\bID="' + re.escape(seller_ref_id) +
                                r'"[^>]*>.*?<[^>]*TextCode[^>]*>([^<]+)</[^>]*TextCode>.*?</[^>]*TextObject>'
                            )
                            m = re.search(obj_pat, content, re.S)
                            if m:
                                val = m.group(1).strip()
                                if val:
                                    return val
                            # 某些 OFD 把文本放在 TextCode 的 X/Y 属性外，值可能在标签之间或属性中
                            m = re.search(
                                r'<[^>]*TextObject\s+[^>]*\bID="' + re.escape(seller_ref_id) +
                                r'"[^>]*>(.*?)</[^>]*TextObject>',
                                content,
                                re.S,
                            )
                            if m:
                                inner = m.group(1)
                                m2 = re.search(r'<[^>]*TextCode[^>]*>([^<]+)<', inner)
                                if m2:
                                    val = m2.group(1).strip()
                                    if val:
                                        return val

                # 2) 回退：直接在 XML 中搜索 SellerName 标签内联值
                for name in z.namelist():
                    if name.endswith('.xml'):
                        with z.open(name) as f:
                            content = f.read().decode('utf-8', errors='replace')
                        for tag in ['SellerName', 'Xfsmc', '销售方名称']:
                            m = re.search(rf'<[^>]*{re.escape(tag)}[^>]*>([^<]+)<', content)
                            if m:
                                val = m.group(1).strip()
                                if val:
                                    return val
        except Exception:
            pass
        return None

    @staticmethod
    def extract_seller_from_pdf(path: Path) -> Optional[str]:
        try:
            import pdfplumber
            with pdfplumber.open(str(path)) as pdf:
                text = ''
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + '\n'

            # 处理竖排/分栏导致字符间出现空格的情况，如 "销 名称：..."
            patterns = [
                r'销\s*售\s*方\s*名\s*称\s*[:：]?\s*\n?\s*([^\n]{2,80})',
                r'销\s*名称\s*[:：]?\s*([^\n]{2,80})',
                r'销\s*售\s*方\s*[:：]?\s*\n?\s*名\s*称\s*[:：]?\s*\n?\s*([^\n]{2,80})',
                r'销售方名称\s*[:：]?\s*\n?\s*([^\n]{2,80})',
            ]
            for pattern in patterns:
                m = re.search(pattern, text)
                if m:
                    seller = m.group(1).strip()
                    seller = re.sub(r'^[：:\s]+', '', seller)
                    # 过滤掉纳税人识别号（15~20 位纯数字/字母），保留短名称如 IBM、3M
                    if len(seller) > 1 and not re.fullmatch(r'[A-Z0-9]{15,20}', seller):
                        return seller
        except Exception:
            pass
        return None

    @staticmethod
    def extract_seller(path: Path) -> Optional[str]:
        ext = path.suffix.lower()
        if ext == '.ofd':
            return InvoiceParser.extract_seller_from_ofd(path)
        elif ext == '.pdf':
            return InvoiceParser.extract_seller_from_pdf(path)
        return None

    @staticmethod
    def sanitize_filename(name: str) -> str:
        name = re.sub(r'[\\/:*?"<>|]', '_', name)
        name = re.sub(r'[\x00-\x1f\x7f]', '', name)
        if len(name) > 200:
            stem = Path(name).stem[:100]
            suffix = Path(name).suffix
            name = f"{stem}{suffix}"
        return name.strip()

    @staticmethod
    def unique_path(path: Path) -> Path:
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
    def format_size(size: int) -> str:
        if size < 1024:
            return f"{size}B"
        elif size < 1024 * 1024:
            return f"{size / 1024:.1f}KB"
        else:
            return f"{size / (1024 * 1024):.1f}MB"


# ---- 报告生成器 ----

class ReportGenerator:
    """扫描目录、解析发票、生成报告"""

    def __init__(self, target_dir: Path, records: Optional[List[Dict[str, Any]]] = None,
                 db: Optional[InvoiceDatabase] = None, logger: Optional[Logger] = None):
        self.target_dir = Path(target_dir)
        self.target_dir.mkdir(parents=True, exist_ok=True)
        self.logger = logger or Logger()
        self.db = db or InvoiceDatabase(str(self.target_dir / '.invoice_db'))
        self.parser = InvoiceParser()
        self.records = records or []

    def scan_directory(self) -> List[Dict[str, Any]]:
        """扫描目录下的所有发票文件，构建记录列表"""
        records = []
        invoice_exts = InvoiceParser.INVOICE_EXTENSIONS

        for month_dir in self.target_dir.iterdir():
            if not month_dir.is_dir() or month_dir.name.startswith('.'):
                continue

            files = [f for f in month_dir.iterdir() if f.is_file() and f.suffix.lower() in invoice_exts]
            if not files:
                continue

            for f in files:
                amount = self.parser.extract_amount(f)
                invoice_id = self.parser.extract_invoice_id_from_filename(f.name)

                record = {
                    'msg_id': 'scan',
                    'subject': f'【本地扫描】{month_dir.name}',
                    'from': '本地文件系统',
                    'date': '',
                    'date_display': month_dir.name,
                    'month': month_dir.name,
                    'files': [{
                        'type': '本地文件',
                        'filename': f.name,
                        'path': str(f.relative_to(self.target_dir)),
                        'size': f.stat().st_size,
                        'status': '成功',
                        'amount': amount,
                    }],
                }
                records.append(record)

        return records

    def process_supplemental(self, dry_run: bool = False) -> List[Dict[str, Any]]:
        """处理 Supplemental 目录：将手动补充的发票归类到对应月份"""
        supplement_dir = self.target_dir.parent / 'Supplemental'
        supplement_dir.mkdir(parents=True, exist_ok=True)

        invoice_exts = InvoiceParser.INVOICE_EXTENSIONS
        files = [f for f in supplement_dir.iterdir() if f.is_file() and f.suffix.lower() in invoice_exts]

        if not files:
            return []

        self.logger.info(f"📂 发现 {len(files)} 个补充发票，开始归类...")
        moved_count = 0
        skip_count = 0
        supplement_records = []

        for src_path in files:
            content_hash = self.parser.compute_file_hash(src_path)

            if dry_run:
                self.logger.info(f"  [预览] 将归类补充发票: {src_path.name}")
                continue

            if self.db.hash_exists(content_hash):
                self.logger.info(f"  补充发票已存在，跳过: {src_path.name}")
                skip_count += 1
                continue

            invoice_date = self.parser.extract_date(src_path)
            if invoice_date:
                month_dir = self.target_dir / invoice_date.strftime('%Y-%m')
                date_display = invoice_date.strftime('%Y-%m-%d')
            else:
                month_dir = self.target_dir / 'unknown'
                date_display = '未知'
            month_dir.mkdir(parents=True, exist_ok=True)

            dest = self.parser.unique_path(month_dir / src_path.name)
            try:
                shutil.move(str(src_path), str(dest))
            except Exception as e:
                self.logger.error(f"  移动失败: {src_path.name} -> {e}")
                continue

            dest = self.parser.try_rename_with_invoice_id(dest, logger=self.logger)

            file_size = dest.stat().st_size
            invoice_id = self.parser.extract_invoice_id_from_filename(dest.name)
            # 优先从文件内容提取真实发票号码，避免文件名临时标识（mt:xxx）导致同一发票被重复记录
            content_id = self.parser.extract_invoice_id_from_file(dest)
            if content_id:
                invoice_id = content_id
            file_format = dest.suffix
            amount = self.parser.extract_amount(dest)

            if not self.db.record_file(
                invoice_id=invoice_id,
                content_hash=content_hash,
                file_path=str(dest.absolute()),
                file_format=file_format,
                file_size=file_size,
                source_email='手动补充',
                source_url='',
            ):
                self.logger.info(f"  发票 {invoice_id or 'N/A'} 已有更高优先级格式，跳过: {dest.name}")
                skip_count += 1
                continue

            moved_count += 1
            self.logger.success(
                f"  归类补充发票: {dest.name}"
                f" -> {month_dir.name}"
                f"{f' 金额:¥{amount:.2f}' if amount else ''}"
            )

            record = {
                'msg_id': 'supplement',
                'subject': '【手动补充发票】',
                'from': '手动补充目录',
                'date': invoice_date.isoformat() if invoice_date else '',
                'date_display': date_display,
                'month': invoice_date.strftime('%Y-%m') if invoice_date else 'unknown',
                'files': [{
                    'type': '手动补充',
                    'filename': dest.name,
                    'path': str(dest.relative_to(self.target_dir)),
                    'size': file_size,
                    'status': '成功',
                    'amount': amount,
                }],
            }
            supplement_records.append(record)

        self.logger.info(f"📂 补充发票归类完成: 移动 {moved_count} 个, 跳过 {skip_count} 个")
        return supplement_records

    def build_html_report(self, records: Optional[List[Dict[str, Any]]] = None) -> str:
        """构建 HTML 核查报告"""
        records = records or self.records
        rows = []
        total_files = 0
        total_amount = 0.0
        month_amounts: Dict[str, float] = {}
        invoice_summary: Dict[str, Dict[str, Any]] = {}
        _e = html_module.escape

        for record in records:
            file_list_html = []
            for f in record['files']:
                total_files += 1
                size_str = self.parser.format_size(f.get('size', 0))
                status_color = 'green' if f.get('status') == '成功' else 'red'
                amt = f.get('amount')

                fp = None
                if f.get('path'):
                    fp = self.target_dir / f['path']
                elif f.get('filename'):
                    for month_dir in self.target_dir.iterdir():
                        if month_dir.is_dir():
                            candidate = month_dir / f['filename']
                            if candidate.exists():
                                fp = candidate
                                break

                if amt is None:
                    try:
                        if fp and fp.exists():
                            amt = self.parser.extract_amount(fp)
                    except Exception:
                        pass
                amt_str = f' <span style="color:#e67e22;font-weight:600">¥{amt:.2f}</span>' if amt else ''
                if amt is not None and f.get('status') == '成功':
                    inv_id = self.parser.extract_invoice_id_from_filename(f['filename'])
                    # 文件名标识不可靠（如 mt:xxx）时，回退到文件内容提取真实号码
                    if not inv_id or inv_id.startswith('mt:') or len(inv_id) < 20:
                        if fp and fp.exists():
                            content_id = self.parser.extract_invoice_id_from_file(fp)
                            if content_id:
                                inv_id = content_id
                    if not inv_id:
                        inv_id = f['filename']
                    if inv_id not in invoice_summary:
                        number = ''
                        date_str = ''
                        seller = ''
                        if fp and fp.exists():
                            number = self.parser.extract_invoice_id_from_file(fp) or ''
                            dt = self.parser.extract_date(fp)
                            if dt:
                                date_str = dt.strftime('%Y-%m-%d')
                            seller = self.parser.extract_seller(fp) or ''
                        if not number:
                            m = re.search(r'(?<![\dA-Za-z_])(\d{20})(?![\dA-Za-z_])', f['filename'])
                            if not m:
                                m = re.search(r'^(\d{20})', f['filename'])
                            if m:
                                number = m.group(1)
                        invoice_summary[inv_id] = {
                            'amount': amt,
                            'number': number,
                            'date': date_str,
                            'seller': seller,
                        }
                        total_amount += amt
                        month = record.get('month', 'unknown')
                        month_amounts[month] = month_amounts.get(month, 0.0) + amt

                file_list_html.append(
                    f'<div class="file-item">'
                    f'<span class="file-type">[{_e(f["type"])}]</span> '
                    f'<span class="file-name">{_e(f["filename"])}</span> '
                    f'<span class="file-size">({size_str})</span>{amt_str} '
                    f'<span style="color:{status_color};font-size:12px">{_e(f.get("status",""))}</span>'
                    f'</div>'
                )

            rows.append(f'''
            <tr>
                <td class="date">{_e(record.get('date_display', ''))}</td>
                <td class="subject">{_e(record.get('subject', ''))}</td>
                <td class="from">{_e(record.get('from', ''))}</td>
                <td class="files">{''.join(file_list_html)}</td>
            </tr>
            ''')

        month_total = sum(month_amounts.values())
        month_summary_rows = ''
        for month, amt in sorted(month_amounts.items()):
            month_summary_rows += f'<tr><td>{_e(month)}</td><td style="text-align:right;font-weight:600;color:#e67e22">¥{amt:.2f}</td></tr>'
        month_summary_rows += (
            f'<tr style="border-top:2px solid #fa8c16">'
            f'<td style="font-weight:700">合计</td>'
            f'<td style="text-align:right;font-weight:700;color:#e67e22">¥{month_total:,.2f}</td>'
            f'</tr>'
        )

        invoice_summary_rows = ''
        for inv_id, info in sorted(
            invoice_summary.items(),
            key=lambda x: x[1].get('date', '') or '9999-99-99',
            reverse=True
        ):
            display_id = inv_id[:30] if len(inv_id) > 30 else inv_id
            number = info.get('number', '') or ''
            date_str = info.get('date', '') or ''
            seller = (info.get('seller', '') or '')[:50]
            invoice_summary_rows += (
                f'<tr>'
                f'<td>{_e(display_id)}</td>'
                f'<td>{_e(number)}</td>'
                f'<td>{_e(date_str)}</td>'
                f'<td>{_e(seller)}</td>'
                f'<td style="text-align:right;font-weight:600;color:#e67e22">¥{info["amount"]:.2f}</td>'
                f'</tr>'
            )
        invoice_summary_rows += (
            f'<tr style="border-top:2px solid #fa8c16">'
            f'<td colspan="4" style="font-weight:700">合计</td>'
            f'<td style="text-align:right;font-weight:700;color:#e67e22">¥{total_amount:,.2f}</td>'
            f'</tr>'
        )

        return f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>发票邮件核查报告</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; margin: 40px; background: #f5f6f7; }}
        .container {{ max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}
        h1 {{ color: #1a1a1a; margin-bottom: 8px; }}
        h2 {{ color: #333; margin-top: 30px; margin-bottom: 12px; font-size: 18px; }}
        .meta {{ color: #666; margin-bottom: 24px; font-size: 14px; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
        th {{ background: #3370ff; color: white; padding: 12px 16px; text-align: left; font-weight: 500; }}
        td {{ padding: 12px 16px; border-bottom: 1px solid #e8e8e8; vertical-align: top; }}
        tr:hover {{ background: #f8f9fa; }}
        .date {{ white-space: nowrap; color: #666; width: 120px; }}
        .subject {{ font-weight: 500; color: #1a1a1a; max-width: 400px; }}
        .from {{ color: #666; max-width: 300px; }}
        .file-item {{ margin: 4px 0; padding: 4px 8px; background: #f0f5ff; border-radius: 4px; display: inline-block; }}
        .file-type {{ color: #3370ff; font-size: 12px; }}
        .file-name {{ color: #1a1a1a; }}
        .file-size {{ color: #999; font-size: 12px; }}
        .summary {{ margin-top: 20px; padding: 16px; background: #e8f3ff; border-radius: 8px; color: #1a1a1a; }}
        .amount-box {{ margin-top: 20px; padding: 16px 20px; background: #fff7e6; border-radius: 8px; border-left: 4px solid #fa8c16; }}
        .amount-box h3 {{ margin: 0 0 8px 0; color: #d46b08; font-size: 16px; }}
        .amount-box .total {{ font-size: 28px; font-weight: 700; color: #e67e22; }}
        .summary-table {{ width: 100%; margin-top: 10px; }}
        .summary-table th {{ background: #fa8c16; }}
        .summary-table td {{ padding: 8px 16px; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📧 发票邮件核查报告</h1>
        <div class="meta">
            生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} &nbsp;|&nbsp;
            匹配邮件: {len(records)} 封 &nbsp;|&nbsp;
            提取文件: {total_files} 个
        </div>

        <div class="amount-box">
            <h3>💰 发票金额总计</h3>
            <div class="total">¥{total_amount:,.2f}</div>
        </div>

        <table>
            <thead>
                <tr>
                    <th>日期</th>
                    <th>邮件标题</th>
                    <th>发件人</th>
                    <th>提取的文件</th>
                </tr>
            </thead>
            <tbody>
                {''.join(rows)}
            </tbody>
        </table>

        <div class="amount-box">
            <h3>📅 按月份汇总</h3>
            <table class="summary-table">
                <thead><tr><th>月份</th><th style="text-align:right">金额</th></tr></thead>
                <tbody>{month_summary_rows}</tbody>
            </table>
        </div>
        <div class="amount-box">
            <h3>📋 按发票汇总</h3>
            <table class="summary-table">
                <thead><tr><th>发票标识</th><th>发票号码</th><th>开票日期</th><th>开票单位</th><th style="text-align:right">金额</th></tr></thead>
                <tbody>{invoice_summary_rows}</tbody>
            </table>
        </div>

        <div class="summary">
            <strong>说明：</strong>本报告由脚本自动生成，展示从邮箱中提取的发票邮件及其对应文件。
            文件保存在 <code>{str(self.target_dir.absolute()).replace(str(Path.home()), "~")}</code> 目录下，按月份分文件夹存放。
        </div>
    </div>
</body>
</html>'''

    def generate_reports(self, records: Optional[List[Dict[str, Any]]] = None) -> Tuple[Optional[Path], Optional[Path]]:
        """生成 JSON 和 HTML 报告"""
        records = records or self.records
        if not records:
            self.logger.info("没有发票记录，跳过报告生成")
            return None, None

        report_dir = Path('reports')
        report_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

        json_path = report_dir / f'invoice_report_{timestamp}.json'
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        self.logger.success(f"JSON 报告已保存: {json_path}")

        html_path = report_dir / f'invoice_report_{timestamp}.html'
        html_content = self.build_html_report(records)
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        self.logger.success(f"HTML 报告已保存: {html_path}")

        return json_path, html_path

    def rebuild_database(self, records: Optional[List[Dict[str, Any]]] = None):
        """根据现有文件重建数据库"""
        records = records or self.records
        db_path = Path(self.db.db_path)

        self.logger.info("重建数据库...")
        self.db.close()
        if db_path.exists():
            db_path.unlink()
        self.db = InvoiceDatabase(str(db_path))

        for record in records:
            for f in record.get('files', []):
                if f.get('status') != '成功' or not f.get('path'):
                    continue
                fp = self.target_dir / f['path']
                if not fp.exists():
                    continue

                content_hash = self.parser.compute_file_hash(fp)
                invoice_id = self.parser.extract_invoice_id_from_filename(fp.name)
                # 回退到文件内容提取真实号码，避免临时标识导致重复
                content_id = self.parser.extract_invoice_id_from_file(fp)
                if content_id:
                    invoice_id = content_id
                file_format = fp.suffix
                file_size = fp.stat().st_size

                self.db.record_file(
                    invoice_id=invoice_id,
                    content_hash=content_hash,
                    file_path=str(fp.absolute()),
                    file_format=file_format,
                    file_size=file_size,
                    source_email='',
                    source_url='',
                )

        self.logger.info(f"数据库重建完成")

    def run(self, process_supplement: bool = True, scan_dir: bool = True,
            rebuild_db: bool = False, dry_run: bool = False) -> List[Dict[str, Any]]:
        """主入口：扫描 + 解析 + 生成报告"""
        self.logger.info("=" * 50)
        self.logger.info("发票报告生成器启动")
        self.logger.info("=" * 50)

        # 1. 处理 Supplemental 目录
        if process_supplement:
            supplement_records = self.process_supplemental(dry_run=dry_run)
            self.records.extend(supplement_records)

        # 2. 扫描目录（如果 records 为空或明确要求扫描）
        if scan_dir and not self.records:
            self.records = self.scan_directory()
        elif scan_dir:
            # 合并扫描结果，避免重复
            scanned = self.scan_directory()
            existing_paths = set()
            for r in self.records:
                for f in r.get('files', []):
                    if f.get('path'):
                        existing_paths.add(f['path'])
            for r in scanned:
                for f in r.get('files', []):
                    if f.get('path') and f['path'] not in existing_paths:
                        self.records.append(r)
                        break

        # 3. 重建数据库
        if rebuild_db:
            self.rebuild_database()

        # 4. 生成报告
        if self.records:
            self.generate_reports()
        else:
            self.logger.info("没有发票记录，跳过报告生成")

        return self.records


def main():
    parser = argparse.ArgumentParser(description='发票报告生成器')
    parser.add_argument('--dir', '-d', default='./invoices', help='发票存放目录')
    parser.add_argument('--supplement', '-s', action='store_true', help='处理 Supplemental 目录')
    parser.add_argument('--rebuild-db', '-r', action='store_true', help='重建数据库')
    parser.add_argument('--dry-run', action='store_true', help='仅预览，不实际移动文件')
    args = parser.parse_args()

    target_dir = Path(args.dir).resolve()
    if not target_dir.exists():
        print(f"目录不存在: {target_dir}")
        sys.exit(1)

    gen = ReportGenerator(target_dir)
    gen.run(
        process_supplement=args.supplement,
        scan_dir=True,
        rebuild_db=args.rebuild_db,
        dry_run=args.dry_run,
    )


if __name__ == '__main__':
    main()
