#!/usr/bin/env python3
"""
发票下载工具 - Web 配置界面启动器
双击运行此脚本，自动启动本地服务并打开浏览器
"""

import hashlib
import http.server
import json
import os
import re
import socket
import socketserver
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime
from pathlib import Path

# ---- 自动找可用端口 ----
def find_free_port(start=8765, max_try=100):
    for port in range(start, start + max_try):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(("127.0.0.1", port))
            s.close()
            return port
        except OSError:
            continue
    return 0

PORT = find_free_port()
PROJECT_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = PROJECT_DIR / "config.json"
RUN_LOG_PATH = PROJECT_DIR / ".web_run.log"

# ---- 运行时状态 ----
run_process = None
run_lock = threading.Lock()


class APIHandler(http.server.SimpleHTTPRequestHandler):
    """扩展 SimpleHTTPRequestHandler，增加 /api/* 路由"""

    def log_message(self, format, *args):
        # 简化日志，不输出每个请求
        pass

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self.end_headers()

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    def _read_body(self):
        content_length = int(self.headers.get("Content-Length", 0))
        if content_length == 0:
            return {}
        body = self.rfile.read(content_length).decode("utf-8")
        try:
            return json.loads(body)
        except Exception:
            return {}

    def do_GET(self):
        if self.path == "/api/config":
            self._handle_get_config()
        elif self.path == "/api/status":
            self._handle_get_status()
        elif self.path == "/api/log":
            self._handle_get_log()
        elif self.path == "/api/reports":
            self._handle_get_reports()
        else:
            # 静态文件服务（从 web/ 目录或根目录）
            blocked = {'/config.json', '/.env', '/.invoice_db', '/download.log',
                       '/.web_run.log', '/cleanup_invoices.py', '/download_invoices.py',
                       '/generate_report.py'}
            if self.path in blocked or '/..' in self.path:
                self.send_response(403)
                self.end_headers()
                return
            super().do_GET()

    def do_POST(self):
        if self.path == "/api/config":
            self._handle_post_config()
        elif self.path == "/api/run":
            self._handle_run()
        elif self.path == "/api/stop":
            self._handle_stop()
        elif self.path == "/api/open-folder":
            self._handle_open_folder()
        elif self.path == "/api/regenerate":
            self._handle_regenerate()
        elif self.path == "/api/delete-report":
            self._handle_delete_report()
        else:
            self.send_response(404)
            self.end_headers()

    def _handle_get_config(self):
        """读取当前 config.json"""
        if CONFIG_PATH.exists():
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    config = json.load(f)
                self._send_json({"success": True, "config": config})
                return
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 500)
                return
        # 返回默认配置模板
        self._send_json({"success": True, "config": self._default_config()})

    @staticmethod
    def _deep_merge(base: dict, update: dict) -> dict:
        for key, value in update.items():
            if isinstance(value, dict) and key in base and isinstance(base[key], dict):
                APIHandler._deep_merge(base[key], value)
            else:
                base[key] = value
        return base

    def _handle_post_config(self):
        """保存配置到 config.json"""
        data = self._read_body()
        config = data.get("config", {})
        try:
            # 合并到现有配置，保留未修改的字段（深度合并）
            existing = {}
            if CONFIG_PATH.exists():
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    existing = json.load(f)
            self._deep_merge(existing, config)

            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(existing, f, ensure_ascii=False, indent=2)
            self._send_json({"success": True})
        except Exception as e:
            self._send_json({"success": False, "error": str(e)}, 500)

    def _handle_run(self):
        """启动 download_invoices.py"""
        global run_process
        with run_lock:
            if run_process is not None and run_process.poll() is None:
                self._send_json({"success": False, "error": "任务已在运行中"}, 409)
                return

            data = self._read_body()
            dry_run = data.get("dry_run", False)

            # 清空日志文件
            RUN_LOG_PATH.write_text("", encoding="utf-8")

            cmd = [sys.executable, str(PROJECT_DIR / "download_invoices.py"), "--config", str(CONFIG_PATH)]
            if dry_run:
                cmd.append("--dry-run")

            try:
                # 使用 CREATE_NEW_PROCESS_GROUP 在 Windows 上支持优雅终止
                kwargs = {}
                if sys.platform == "win32":
                    kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP

                run_process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    cwd=str(PROJECT_DIR),
                    **kwargs,
                )

                # 启动线程读取输出到日志文件
                def _read_output():
                    proc = run_process
                    if proc is None:
                        return
                    with open(RUN_LOG_PATH, "a", encoding="utf-8") as log_f:
                        log_f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 任务启动\n")
                        log_f.write(f"命令: {' '.join(cmd)}\n")
                        log_f.write("-" * 50 + "\n")
                        try:
                            for line in proc.stdout:
                                log_f.write(line)
                                log_f.flush()
                        except Exception:
                            pass
                        try:
                            proc.wait()
                        except Exception:
                            pass
                        log_f.write("-" * 50 + "\n")
                        log_f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 任务结束，返回码: {getattr(proc, 'returncode', None)}\n")

                threading.Thread(target=_read_output, daemon=True).start()
                self._send_json({"success": True, "pid": run_process.pid})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 500)

    def _handle_stop(self):
        """终止运行中的任务"""
        global run_process
        with run_lock:
            if run_process is None or run_process.poll() is not None:
                self._send_json({"success": False, "error": "没有正在运行的任务"}, 400)
                return
            try:
                if sys.platform == "win32":
                    run_process.send_signal(subprocess.signal.CTRL_BREAK_EVENT)
                else:
                    run_process.terminate()
                run_process = None
                self._send_json({"success": True})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 500)

    def _handle_get_status(self):
        """获取任务运行状态"""
        global run_process
        running = run_process is not None and run_process.poll() is None
        self._send_json({"success": True, "running": running})

    def _handle_get_log(self):
        """获取运行日志"""
        try:
            if RUN_LOG_PATH.exists():
                text = RUN_LOG_PATH.read_text(encoding="utf-8", errors="replace")
            else:
                text = ""
            self._send_json({"success": True, "log": text})
        except Exception as e:
            self._send_json({"success": False, "error": str(e)}, 500)

    def _handle_get_reports(self):
        """列出 reports 目录下的 HTML 报告文件"""
        try:
            reports_dir = PROJECT_DIR / "reports"
            reports = []
            if reports_dir.exists():
                files = list(reports_dir.glob("invoice_report_*.html"))
                files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
                for f in files:
                    stat = f.stat()
                    reports.append({
                        "name": f.name,
                        "path": f"/reports/{f.name}",
                        "size": stat.st_size,
                        "mtime": stat.st_mtime,
                        "mtime_str": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    })
            self._send_json({"success": True, "reports": reports})
        except Exception as e:
            self._send_json({"success": False, "error": str(e)}, 500)

    def _handle_open_folder(self):
        """用系统文件管理器打开指定目录"""
        data = self._read_body()
        folder_type = data.get("type", "")

        if folder_type == "invoices":
            cfg = self._load_config_dict()
            path = PROJECT_DIR / cfg.get("target_dir", "invoices")
        elif folder_type == "supplemental":
            path = PROJECT_DIR / "Supplemental"
        elif folder_type == "reports":
            path = PROJECT_DIR / "reports"
        else:
            self._send_json({"success": False, "error": "未知目录类型"}, 400)
            return

        path = path.resolve()
        if not path.exists():
            path.mkdir(parents=True, exist_ok=True)

        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            elif sys.platform == "win32":
                subprocess.Popen(["explorer", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
            self._send_json({"success": True})
        except Exception as e:
            self._send_json({"success": False, "error": str(e)}, 500)

    def _handle_regenerate(self):
        """调用 generate_report.py 重新生成报告"""
        global run_process
        with run_lock:
            if run_process is not None and run_process.poll() is None:
                self._send_json({"success": False, "error": "任务已在运行中"}, 409)
                return

            data = self._read_body()
            supplement = data.get("supplement", True)
            rebuild_db = data.get("rebuild_db", False)
            dry_run = data.get("dry_run", False)

            cfg = self._load_config_dict()
            target_dir = cfg.get("target_dir", "invoices")

            RUN_LOG_PATH.write_text("", encoding="utf-8")

            cmd = [sys.executable, str(PROJECT_DIR / "generate_report.py"), "--dir", str(target_dir)]
            if supplement:
                cmd.append("--supplement")
            if rebuild_db:
                cmd.append("--rebuild-db")
            if dry_run:
                cmd.append("--dry-run")

            try:
                kwargs = {}
                if sys.platform == "win32":
                    kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP

                run_process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    cwd=str(PROJECT_DIR),
                    **kwargs,
                )

                def _read_output():
                    proc = run_process
                    if proc is None:
                        return
                    with open(RUN_LOG_PATH, "a", encoding="utf-8") as log_f:
                        log_f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 重新生成报告任务启动\n")
                        log_f.write(f"命令: {' '.join(cmd)}\n")
                        log_f.write("-" * 50 + "\n")
                        try:
                            for line in proc.stdout:
                                log_f.write(line)
                                log_f.flush()
                        except Exception:
                            pass
                        try:
                            proc.wait()
                        except Exception:
                            pass
                        log_f.write("-" * 50 + "\n")
                        log_f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 任务结束，返回码: {getattr(proc, 'returncode', None)}\n")

                threading.Thread(target=_read_output, daemon=True).start()
                self._send_json({"success": True, "pid": run_process.pid})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 500)

    # ---- 报告操作 ----

    def _handle_delete_report(self):
        """删除指定报告文件（HTML + 对应 JSON）"""
        data = self._read_body()
        filename = data.get("name", "")
        if not filename or ".." in filename or "/" in filename or "\\" in filename:
            self._send_json({"success": False, "error": "无效文件名"}, 400)
            return

        reports_dir = PROJECT_DIR / "reports"
        target = (reports_dir / filename).resolve()
        json_name = filename.replace(".html", ".json")
        json_target = (reports_dir / json_name).resolve()

        if not str(target).startswith(str(reports_dir)):
            self._send_json({"success": False, "error": "路径越界"}, 403)
            return

        try:
            deleted = []
            if target.exists():
                target.unlink()
                deleted.append(filename)
            if json_target.exists():
                json_target.unlink()
                deleted.append(json_name)
            self._send_json({"success": True, "deleted": deleted})
        except Exception as e:
            self._send_json({"success": False, "error": str(e)}, 500)

    def _load_config_dict(self):
        """读取当前 config.json，返回字典"""
        if CONFIG_PATH.exists():
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return self._default_config()

    @staticmethod
    def _default_config():
        return {
            "imap_server": "imap.163.com",
            "imap_port": 993,
            "use_ssl": True,
            "email": "",
            "password": "",
            "search_keywords": ["发票"],
            "folders": ["INBOX"],
            "blocked_senders": [],
            "date_range": {},
            "target_dir": "invoices",
            "log_file": "download.log",
            "timeout": 60,
            "max_retries": 3,
            "headers": {
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            },
        }


class ThreadedHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    allow_reuse_address = True


def main():
    os.chdir(PROJECT_DIR)
    server = ThreadedHTTPServer(("127.0.0.1", PORT), APIHandler)
    actual_port = server.server_address[1]
    url = f"http://127.0.0.1:{actual_port}/Start.html"

    print(f"\n{'='*50}")
    print(f"发票下载工具 Web 配置界面")
    print(f"{'='*50}")
    print(f"服务地址: {url}")
    print(f"按 Ctrl+C 停止服务")
    print(f"{'='*50}\n")

    # 启动浏览器
    def open_browser():
        import time
        time.sleep(0.8)
        webbrowser.open(url)

    threading.Thread(target=open_browser, daemon=True).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n正在停止服务...")
        server.shutdown()


if __name__ == "__main__":
    main()
