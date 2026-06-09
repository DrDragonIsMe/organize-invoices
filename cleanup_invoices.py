#!/usr/bin/env python3
"""
清理已下载的发票文件

功能：
1. 删除所有无效的 HTML 阅读器页面
2. 按内容哈希去重，只保留一个副本（优先保留文件名含发票号码的）
3. 重建 SQLite 数据库

使用方法：
    python cleanup_invoices.py
    python cleanup_invoices.py --dry-run    # 仅预览，不删除
"""

import argparse
import hashlib
import os
import re
import sqlite3
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple


def compute_file_hash(path: Path) -> str:
    """计算文件的 SHA-256 哈希"""
    sha = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            sha.update(chunk)
    return sha.hexdigest()


def extract_invoice_id_from_filename(filename: str) -> Optional[str]:
    """从文件名中提取发票标识"""
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


def extract_invoice_id_from_file(path: Path) -> Optional[str]:
    """从文件内容中提取发票号码"""
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
                result = os.popen(f'strings "{path}"').read()
                m = re.search(r'\b(\d{20})\b', result)
                if m:
                    return m.group(1)
            except Exception:
                pass
    except Exception:
        pass
    return None


def get_format_priority(ext: str) -> int:
    """格式优先级"""
    priorities = {
        '.pdf': 4, '.png': 3, '.jpg': 3, '.jpeg': 3,
        '.gif': 3, '.webp': 3, '.ofd': 2, '.xml': 1,
        '.html': 0, '.bin': 0,
    }
    return priorities.get(ext.lower(), 0)


def main():
    parser = argparse.ArgumentParser(description='清理发票下载目录')
    parser.add_argument('--dry-run', action='store_true', help='仅预览，不删除')
    parser.add_argument('--target-dir', default='invoices', help='发票存放目录')
    args = parser.parse_args()

    target_dir = Path(args.target_dir).resolve()
    if not target_dir.exists():
        print(f"目录不存在: {target_dir}")
        sys.exit(1)

    db_path = target_dir / '.invoice_db'

    # ---- 第一阶段：收集所有文件 ----
    all_files: List[Path] = []
    html_files: List[Path] = []

    for f in target_dir.rglob('*'):
        if f.is_file() and f.name not in {'.invoice_db', '.DS_Store'}:
            if f.suffix.lower() == '.html':
                html_files.append(f)
            else:
                all_files.append(f)

    print(f"发现 {len(html_files)} 个 HTML 文件（将被删除）")
    print(f"发现 {len(all_files)} 个非 HTML 文件")

    # ---- 第二阶段：删除 HTML 文件 ----
    if not args.dry_run:
        for f in html_files:
            f.unlink()
            print(f"  删除 HTML: {f.relative_to(target_dir)}")
    else:
        for f in html_files:
            print(f"  [预览删除] HTML: {f.relative_to(target_dir)}")

    # ---- 第三阶段：按内容哈希去重 ----
    print("\n正在计算文件哈希并去重...")

    hash_groups: Dict[str, List[Tuple[Path, Optional[str], Optional[str], int]]] = {}

    for f in all_files:
        try:
            h = compute_file_hash(f)
            inv_file = extract_invoice_id_from_filename(f.name)
            inv_content = extract_invoice_id_from_file(f)
            priority = get_format_priority(f.suffix)
            hash_groups.setdefault(h, []).append((f, inv_file, inv_content, priority))
        except Exception as e:
            print(f"  警告: 无法处理 {f}: {e}")

    # 决定保留哪个文件
    to_delete: List[Path] = []
    to_keep: List[Tuple[Path, str, Optional[str], int]] = []

    for h, group in hash_groups.items():
        if len(group) == 1:
            f, inv_file, inv_content, priority = group[0]
            invoice_id = inv_content or inv_file
            to_keep.append((f, h, invoice_id, priority))
            continue

        # 有重复，选择最佳文件保留
        def score(item):
            f, inv_file, inv_content, priority = item
            s = priority
            if inv_file and re.match(r'^\d{20}$', inv_file):
                s += 10
            if inv_content:
                s += 8
            if not re.search(r'_\d+$', f.stem):
                s += 5
            return s

        group.sort(key=score, reverse=True)
        best = group[0]
        to_keep.append((best[0], h, best[2] or best[1], best[3]))

        for item in group[1:]:
            to_delete.append(item[0])

    print(f"去重结果: 保留 {len(to_keep)} 个文件，删除 {len(to_delete)} 个重复文件")

    if args.dry_run:
        for f in to_delete:
            print(f"  [预览删除] 重复: {f.relative_to(target_dir)}")
    else:
        for f in to_delete:
            f.unlink()
            print(f"  删除重复: {f.relative_to(target_dir)}")

    # ---- 第四阶段：清理空目录 ----
    if not args.dry_run:
        empty_dirs = []
        for d in sorted(target_dir.rglob('*'), key=lambda x: len(str(x)), reverse=True):
            if d.is_dir() and not any(d.iterdir()):
                empty_dirs.append(d)
        for d in empty_dirs:
            d.rmdir()
            print(f"  删除空目录: {d.relative_to(target_dir)}")

    # ---- 第五阶段：重建数据库 ----
    if not args.dry_run:
        print("\n重建数据库...")
        if db_path.exists():
            db_path.unlink()

        conn = sqlite3.connect(str(db_path))
        conn.executescript('''
            CREATE TABLE downloaded_files (
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
            CREATE INDEX idx_invoice_id ON downloaded_files(invoice_id);
            CREATE INDEX idx_content_hash ON downloaded_files(content_hash);

            CREATE TABLE invoice_best_format (
                invoice_id TEXT PRIMARY KEY,
                best_format TEXT,
                best_priority INTEGER,
                file_path TEXT,
                content_hash TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        ''')

        for f, h, invoice_id, priority in to_keep:
            file_format = f.suffix
            file_size = f.stat().st_size
            abs_path = str(f.absolute())

            conn.execute(
                '''INSERT INTO downloaded_files
                   (invoice_id, content_hash, file_path, file_format, format_priority, file_size, source_email, source_url)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                (invoice_id, h, abs_path, file_format, priority, file_size, '', '')
            )

            if invoice_id:
                conn.execute(
                    '''INSERT INTO invoice_best_format (invoice_id, best_format, best_priority, file_path, content_hash)
                       VALUES (?, ?, ?, ?, ?)
                       ON CONFLICT(invoice_id) DO UPDATE SET
                           best_format=excluded.best_format,
                           best_priority=excluded.best_priority,
                           file_path=excluded.file_path,
                           content_hash=excluded.content_hash,
                           updated_at=CURRENT_TIMESTAMP
                       WHERE excluded.best_priority >= invoice_best_format.best_priority''',
                    (invoice_id, file_format, priority, abs_path, h)
                )

        conn.commit()
        conn.close()
        print(f"数据库重建完成，共记录 {len(to_keep)} 个文件")
    else:
        print("\n[预览模式] 不重建数据库")

    print("\n清理完成！")


if __name__ == '__main__':
    main()
