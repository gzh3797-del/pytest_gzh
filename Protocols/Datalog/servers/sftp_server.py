# -*- coding: utf-8 -*-
"""
SFTP 服务器 — 接收网关 DataLog 文件推送
依赖：pip install paramiko
"""
import logging
import os
import socket
import stat
import threading

log = logging.getLogger(__name__)


def _patch_paramiko_sha1_kex():
    """
    网关使用 libssh2 1.8.2，仅支持 diffie-hellman-group14-sha1（RFC 4253 固定分组）。
    paramiko 5.0 已移除 SHA1 kex，需注册一个基于 group14 但使用 SHA1 的处理器。
    libssh2 行为：当服务端仅广播 group14-sha1 时，使用标准 KEXDH_INIT（packet 30）流程。
    """
    from hashlib import sha1
    from paramiko.kex_group14 import KexGroup14SHA256
    from paramiko.transport import Transport

    if 'diffie-hellman-group14-sha1' in Transport._kex_info:
        return

    class _KexGroup14SHA1(KexGroup14SHA256):
        """diffie-hellman-group14-sha1：与 KexGroup14SHA256 完全相同，仅将哈希换为 SHA1。"""
        name = 'diffie-hellman-group14-sha1'
        hash_algo = sha1

    Transport._kex_info['diffie-hellman-group14-sha1'] = _KexGroup14SHA1

    # RSAKey.HASHES 未包含 ssh-rsa（SHA1），补充以支持旧式 RSA 签名
    from cryptography.hazmat.primitives import hashes as _hashes
    import paramiko as _paramiko
    if 'ssh-rsa' not in _paramiko.RSAKey.HASHES:
        _paramiko.RSAKey.HASHES['ssh-rsa'] = _hashes.SHA1
        _paramiko.RSAKey.HASHES['ssh-rsa-cert-v01@openssh.com'] = _hashes.SHA1

    log.debug("已向 paramiko 注册 diffie-hellman-group14-sha1 SHA1 处理器（兼容 libssh2 1.8.x）")

# ─────────────────────────────────────────────────────────────────────────────
# Paramiko 实现
# ─────────────────────────────────────────────────────────────────────────────

def _make_sftp_server_class(root_dir: str):
    """工厂函数：生成绑定 root_dir 的 SFTPServerInterface 子类。"""
    import paramiko

    class _SFTPHandle(paramiko.SFTPHandle):
        def stat(self):
            try:
                return paramiko.SFTPAttributes.from_stat(os.fstat(self.readfile.fileno()))
            except Exception:
                return paramiko.SFTP_FAILURE

        def chattr(self, attr):
            return paramiko.SFTP_OK

    class _SFTPServerInterface(paramiko.SFTPServerInterface):
        ROOT = root_dir

        def _realpath(self, path: str) -> str:
            # 防止路径穿越
            real = os.path.realpath(os.path.join(self.ROOT, path.lstrip("/")))
            if not real.startswith(os.path.realpath(self.ROOT)):
                return self.ROOT
            return real

        def list_folder(self, path):
            real = self._realpath(path)
            out = []
            try:
                for name in os.listdir(real):
                    attr = paramiko.SFTPAttributes.from_stat(
                        os.stat(os.path.join(real, name)))
                    attr.filename = name
                    out.append(attr)
            except OSError as e:
                return paramiko.SFTPServer.convert_errno(e.errno)
            return out

        def stat(self, path):
            try:
                return paramiko.SFTPAttributes.from_stat(os.stat(self._realpath(path)))
            except OSError as e:
                return paramiko.SFTPServer.convert_errno(e.errno)

        lstat = stat

        def open(self, path, flags, attr):
            real = self._realpath(path)
            os.makedirs(os.path.dirname(real), exist_ok=True)
            try:
                if flags & os.O_WRONLY:
                    mode = "ab" if (flags & os.O_APPEND) else "wb"
                elif flags & os.O_RDWR:
                    mode = "a+b" if (flags & os.O_APPEND) else "r+b"
                else:
                    mode = "rb"
                f = open(real, mode)
                h = _SFTPHandle(flags)
                h.readfile = f
                h.writefile = f
                return h
            except OSError as e:
                return paramiko.SFTPServer.convert_errno(e.errno)

        def mkdir(self, path, attr):
            try:
                os.mkdir(self._realpath(path))
                return paramiko.SFTP_OK
            except OSError as e:
                return paramiko.SFTPServer.convert_errno(e.errno)

        def rmdir(self, path):
            try:
                os.rmdir(self._realpath(path))
                return paramiko.SFTP_OK
            except OSError as e:
                return paramiko.SFTPServer.convert_errno(e.errno)

        def remove(self, path):
            try:
                os.remove(self._realpath(path))
                return paramiko.SFTP_OK
            except OSError as e:
                return paramiko.SFTPServer.convert_errno(e.errno)

        def rename(self, oldpath, newpath):
            try:
                os.rename(self._realpath(oldpath), self._realpath(newpath))
                return paramiko.SFTP_OK
            except OSError as e:
                return paramiko.SFTPServer.convert_errno(e.errno)

        def chattr(self, path, attr):
            return paramiko.SFTP_OK

    return _SFTPServerInterface


def _make_server_interface(user: str, password: str):
    import paramiko

    class _ServerInterface(paramiko.ServerInterface):
        def __init__(self):
            self._event = threading.Event()

        def check_channel_request(self, kind, chanid):
            if kind == "session":
                return paramiko.OPEN_SUCCEEDED
            return paramiko.OPEN_FAILED_ADMINISTRATIVELY_PROHIBITED

        def check_auth_password(self, username, passwd):
            if username == user and passwd == password:
                return paramiko.AUTH_SUCCESSFUL
            return paramiko.AUTH_FAILED

        def check_auth_none(self, username):
            return paramiko.AUTH_FAILED

        def get_allowed_auths(self, username):
            return "password"

    return _ServerInterface


def _generate_host_key():
    """生成临时 RSA 主机密钥（每次启动不同，不持久化）。"""
    import paramiko
    return paramiko.RSAKey.generate(2048)


def _handle_client(conn: socket.socket, host_key, server_iface_cls, sftp_iface_cls):
    """处理单个 SFTP 客户端连接。"""
    import paramiko
    import time
    _patch_paramiko_sha1_kex()
    transport = None
    try:
        transport = paramiko.Transport(conn)
        # libssh2 1.8.x 兼容：仅广播 group14-sha1，避免与 GEX-SHA256 的 KEXINIT 歧义
        transport._preferred_kex = ('diffie-hellman-group14-sha1',)
        transport._preferred_keys = (
            ('ssh-rsa',)
            + paramiko.Transport._preferred_keys
        )
        transport._preferred_pubkeys = (
            ('ssh-rsa',)
            + paramiko.Transport._preferred_pubkeys
        )
        # 禁止广播 kex-strict-*（libssh2 1.8.x 不支持，可能导致立即断连）
        transport.advertise_strict_kex = False
        transport.add_server_key(host_key)
        transport.set_subsystem_handler(
            "sftp", paramiko.SFTPServer, sftp_iface_cls
        )
        server = server_iface_cls()
        transport.start_server(server=server)
        chan = transport.accept(60)
        if chan is not None:
            # 等待 SFTP 子系统运行完毕（让客户端完成测试/传输后自行断开）
            deadline = time.time() + 120
            while transport.is_active() and time.time() < deadline:
                time.sleep(0.5)
    except Exception as e:
        log.debug("SFTP 连接处理异常：%s", e)
    finally:
        if transport:
            try:
                transport.close()
            except Exception:
                pass
        try:
            conn.close()
        except Exception:
            pass


def start_sftp_server(
    host: str,
    port: int,
    user: str,
    password: str,
    data_dir: str,
    stop_event: threading.Event = None,
) -> tuple:
    """
    在后台线程启动 SFTP 服务器，文件保存到 data_dir。
    返回 (thread, stop_event)，调用 stop_event.set() 停止服务器。
    """
    try:
        import paramiko  # noqa: F401
    except ImportError:
        raise ImportError("请先安装 paramiko：pip install paramiko")

    os.makedirs(data_dir, exist_ok=True)

    host_key = _generate_host_key()
    sftp_iface_cls = _make_sftp_server_class(data_dir)
    server_iface_cls = _make_server_interface(user, password)

    if stop_event is None:
        stop_event = threading.Event()

    def _run():
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, True)
        sock.bind((host, port))
        sock.listen(10)
        sock.settimeout(1.0)
        log.info("SFTP 服务器启动：%s:%d  目录：%s", host, port, data_dir)
        while not stop_event.is_set():
            try:
                conn, addr = sock.accept()
                log.debug("SFTP 新连接：%s", addr)
                t = threading.Thread(
                    target=_handle_client,
                    args=(conn, host_key, server_iface_cls, sftp_iface_cls),
                    daemon=True,
                )
                t.start()
            except socket.timeout:
                continue
            except Exception as e:
                if not stop_event.is_set():
                    log.warning("SFTP accept 异常：%s", e)
        sock.close()
        log.info("SFTP 服务器已停止")

    t = threading.Thread(target=_run, daemon=True, name="sftp-server")
    t.start()
    return t, stop_event
