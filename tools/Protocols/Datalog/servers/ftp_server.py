# -*- coding: utf-8 -*-
"""
FTP 服务器 — 接收网关 DataLog 文件推送
依赖：pip install pyftpdlib
"""
import logging
import os
import threading

log = logging.getLogger(__name__)


def start_ftp_server(
    host: str,
    port: int,
    user: str,
    password: str,
    data_dir: str,
    stop_event: threading.Event = None,
) -> tuple:
    """
    在后台线程启动 FTP 服务器，文件保存到 data_dir。
    返回 (thread, stop_event)，调用 stop_event.set() 停止服务器。
    """
    try:
        from pyftpdlib.authorizers import DummyAuthorizer
        from pyftpdlib.handlers import FTPHandler
        from pyftpdlib.servers import FTPServer
    except ImportError:
        raise ImportError("请先安装 pyftpdlib：pip install pyftpdlib")

    os.makedirs(data_dir, exist_ok=True)

    authorizer = DummyAuthorizer()
    # perm: e=目录列表 l=列文件 r=读 a=追加 d=删除 f=重命名 m=创建目录 w=写 M=修改权限
    authorizer.add_user(user, password, data_dir, perm="elradfmwMT")

    handler = FTPHandler
    handler.authorizer = authorizer
    handler.passive_ports = range(60000, 60100)
    handler.banner = "DataLog FTP Server ready"

    server = FTPServer((host, port), handler)
    server.max_cons = 20
    server.max_cons_per_ip = 5

    if stop_event is None:
        stop_event = threading.Event()

    def _run():
        log.info("FTP 服务器启动：%s:%d  目录：%s", host, port, data_dir)
        try:
            while not stop_event.is_set():
                server.serve_forever(timeout=1, blocking=False)
        finally:
            server.close_all()
            log.info("FTP 服务器已停止")

    t = threading.Thread(target=_run, daemon=True, name="ftp-server")
    t.start()
    return t, stop_event
