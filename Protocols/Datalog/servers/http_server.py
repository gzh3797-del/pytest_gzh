# -*- coding: utf-8 -*-
"""
HTTP / HTTPS 服务器 — 接收网关 DataLog 文件推送

支持：
  - POST /upload            → 接收原始 body（Content-Type 任意）
  - POST /upload/<filename> → body 存储为指定文件名
  - multipart/form-data     → 自动解析 file 字段

HTTPS 模式：需在 config 中配置自签名证书路径
"""
import cgi
import io
import logging
import os
import ssl
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse

log = logging.getLogger(__name__)


def _make_handler(data_dir: str, counter: list):
    """工厂函数：生成绑定 data_dir 的 HTTP 请求处理类。"""

    class _Handler(BaseHTTPRequestHandler):
        def log_message(self, fmt, *args):
            log.debug("HTTP %s - " + fmt, self.address_string(), *args)

        def _save_file(self, filename: str, data: bytes):
            os.makedirs(data_dir, exist_ok=True)
            fpath = os.path.join(data_dir, filename)
            with open(fpath, "wb") as f:
                f.write(data)
            counter[0] += 1
            log.info("HTTP 收到文件：%s（%d bytes，累计 %d 个）",
                     filename, len(data), counter[0])
            return fpath

        def _respond(self, code: int, msg: str = ""):
            body = msg.encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_POST(self):
            ctype = self.headers.get("Content-Type", "")
            length = int(self.headers.get("Content-Length", 0))

            # 从 URL path 提取文件名（如 /upload/Logger1-xxx.json）
            path = urlparse(self.path).path.lstrip("/")
            parts = path.split("/", 1)
            url_filename = parts[-1] if len(parts) > 1 else ""

            if "multipart/form-data" in ctype:
                # multipart：解析 file 字段
                environ = {
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": ctype,
                    "CONTENT_LENGTH": str(length),
                }
                form = cgi.FieldStorage(
                    fp=self.rfile,
                    headers=self.headers,
                    environ=environ,
                )
                saved = []
                for key in form.keys():
                    item = form[key]
                    if hasattr(item, "filename") and item.filename:
                        fname = os.path.basename(item.filename) or f"datalog_{counter[0]+1}.bin"
                        self._save_file(fname, item.file.read())
                        saved.append(fname)
                if saved:
                    self._respond(200, f"saved: {', '.join(saved)}")
                else:
                    self._respond(400, "no file field found in multipart")
            else:
                # 原始 body
                body = self.rfile.read(length) if length > 0 else b""
                if not body:
                    self._respond(400, "empty body")
                    return
                # 从 Content-Disposition 尝试取文件名
                cd = self.headers.get("Content-Disposition", "")
                fname = ""
                if 'filename="' in cd:
                    fname = cd.split('filename="')[1].split('"')[0]
                fname = fname or url_filename or f"datalog_{counter[0]+1}.bin"
                self._save_file(fname, body)
                self._respond(200, f"saved: {fname}")

        def do_GET(self):
            self._respond(200, "DataLog HTTP Server OK")

    return _Handler


def start_http_server(
    host: str,
    port: int,
    data_dir: str,
    stop_event: threading.Event = None,
    ssl_certfile: str = "",
    ssl_keyfile: str = "",
) -> tuple:
    """
    在后台线程启动 HTTP（或 HTTPS）服务器。
    返回 (thread, stop_event, file_counter)。
    file_counter 是 [int] 列表，实时记录已接收文件数。
    """
    os.makedirs(data_dir, exist_ok=True)

    if stop_event is None:
        stop_event = threading.Event()

    counter = [0]
    handler_cls = _make_handler(data_dir, counter)
    server = HTTPServer((host, port), handler_cls)
    server.timeout = 1.0

    if ssl_certfile and ssl_keyfile:
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ctx.load_cert_chain(ssl_certfile, ssl_keyfile)
        server.socket = ctx.wrap_socket(server.socket, server_side=True)
        proto = "HTTPS"
    else:
        proto = "HTTP"

    def _run():
        log.info("%s 服务器启动：%s:%d  目录：%s", proto, host, port, data_dir)
        while not stop_event.is_set():
            server.handle_request()
        server.server_close()
        log.info("%s 服务器已停止", proto)

    t = threading.Thread(target=_run, daemon=True, name=f"{proto.lower()}-server")
    t.start()
    return t, stop_event, counter
