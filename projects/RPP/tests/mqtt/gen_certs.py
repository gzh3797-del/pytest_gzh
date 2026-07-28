# -*- coding: utf-8 -*-
"""
gen_certs.py — 生成自签 mTLS 证书（CA + 服务器 + 客户端）

使用 OpenSSL CLI（兼容 HMI1-7 / AcuRev4100 等嵌入式设备的 SSL 库）。
输出到 Protocols/MQTT/certs/ 目录（可通过 --out 指定）。

用法：
  python Protocols/MQTT/gen_certs.py
  python Protocols/MQTT/gen_certs.py --host 192.168.2.61
  python Protocols/MQTT/gen_certs.py --host 192.168.2.61 192.168.2.62 --days 365
  python Protocols/MQTT/gen_certs.py --out /tmp/my_certs
"""

from __future__ import annotations

import argparse
import ipaddress
import os
import socket
import subprocess
import sys
import tempfile
from pathlib import Path

# ── 默认输出目录（脚本所在目录下的 certs/） ──────────────────────────────────
_DEFAULT_OUT = Path(__file__).parent / "certs"

# ── OpenSSL 查找路径（按优先级） ──────────────────────────────────────────────
_OPENSSL_CANDIDATES = [
    "openssl",
    r"C:\Program Files\OpenSSL-Win64\bin\openssl.exe",
    r"C:\Program Files (x86)\OpenSSL-Win32\bin\openssl.exe",
    r"C:\OpenSSL-Win64\bin\openssl.exe",
]


def _find_openssl() -> str:
    """返回可用的 openssl 可执行路径，找不到则报错退出。"""
    for candidate in _OPENSSL_CANDIDATES:
        try:
            result = subprocess.run(
                [candidate, "version"],
                capture_output=True, timeout=5,
            )
            if result.returncode == 0:
                return candidate
        except (FileNotFoundError, OSError):
            continue
    print(
        "[ERROR] 未找到 openssl 可执行文件。\n"
        "请安装 OpenSSL 并确保其在 PATH 中，或安装到以下位置之一：\n"
        + "\n".join(f"  {p}" for p in _OPENSSL_CANDIDATES[1:])
    )
    sys.exit(1)


def _run(openssl: str, args: list[str]) -> None:
    """运行 openssl 子命令，失败时报错退出。"""
    cmd = [openssl] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"[ERROR] openssl 命令失败：{' '.join(cmd)}")
        if result.stderr:
            print(result.stderr.strip())
        sys.exit(1)


def _detect_local_ips() -> list[str]:
    """检测本机所有非回环 IPv4 地址。"""
    ips: set[str] = set()
    for target in ("8.8.8.8", "192.168.0.1", "10.0.0.1"):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect((target, 80))
            ips.add(s.getsockname()[0])
            s.close()
            break
        except OSError:
            pass
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if not ip.startswith("127."):
                ips.add(ip)
    except OSError:
        pass
    return sorted(ips)


def _make_ca(openssl: str, out_dir: Path, days: int) -> None:
    """生成 CA 私钥和自签证书。"""
    _run(openssl, ["genrsa", "-out", str(out_dir / "ca.key"), "2048"])
    _run(openssl, [
        "req", "-new", "-x509", "-days", str(days),
        "-key", str(out_dir / "ca.key"),
        "-out", str(out_dir / "ca.crt"),
        "-subj", "/O=AutoTest/CN=MQTT-Test-CA",
    ])


def _make_server_cert(
    openssl: str,
    out_dir: Path,
    extra_hosts: list[str],
    days: int,
) -> None:
    """生成服务器私钥和证书（含 SAN）。"""
    csr = out_dir / "_server.csr"
    ext_file = out_dir / "_server_ext.cnf"

    _run(openssl, ["genrsa", "-out", str(out_dir / "server.key"), "2048"])
    _run(openssl, [
        "req", "-new",
        "-key", str(out_dir / "server.key"),
        "-out", str(csr),
        "-subj", "/O=AutoTest/CN=www.accu.com",
    ])

    # 构造 SAN 列表：始终包含 127.0.0.1 / localhost，再加调用者传入的 IP/域名
    san_lines: list[str] = ["DNS.1 = localhost", "IP.1 = 127.0.0.1"]
    dns_idx, ip_idx = 2, 2
    for h in extra_hosts:
        try:
            ipaddress.ip_address(h)
            san_lines.append(f"IP.{ip_idx} = {h}")
            ip_idx += 1
        except ValueError:
            san_lines.append(f"DNS.{dns_idx} = {h}")
            dns_idx += 1

    ext_content = (
        "[v3_req]\n"
        "basicConstraints = CA:FALSE\n"
        "subjectAltName = @alt_names\n"
        "[alt_names]\n"
        + "\n".join(san_lines) + "\n"
    )
    ext_file.write_text(ext_content, encoding="ascii")

    _run(openssl, [
        "x509", "-req", "-days", str(days),
        "-in", str(csr),
        "-CA", str(out_dir / "ca.crt"),
        "-CAkey", str(out_dir / "ca.key"),
        "-CAcreateserial",
        "-out", str(out_dir / "server.crt"),
        "-extfile", str(ext_file),
        "-extensions", "v3_req",
    ])

    for tmp in (csr, ext_file, out_dir / "ca.srl"):
        tmp.unlink(missing_ok=True)


def _make_client_cert(openssl: str, out_dir: Path, days: int) -> None:
    """生成客户端私钥和证书。"""
    csr = out_dir / "_client.csr"

    _run(openssl, ["genrsa", "-out", str(out_dir / "client.key"), "2048"])
    _run(openssl, [
        "req", "-new",
        "-key", str(out_dir / "client.key"),
        "-out", str(csr),
        "-subj", "/O=AutoTest/CN=mqtt-client",
    ])
    _run(openssl, [
        "x509", "-req", "-days", str(days),
        "-in", str(csr),
        "-CA", str(out_dir / "ca.crt"),
        "-CAkey", str(out_dir / "ca.key"),
        "-CAcreateserial",
        "-out", str(out_dir / "client.crt"),
    ])

    for tmp in (csr, out_dir / "ca.srl"):
        tmp.unlink(missing_ok=True)


def generate_certs(
    out_dir: Path,
    extra_hosts: list[str],
    days: int,
) -> None:
    """生成全套 mTLS 证书（CA + 服务器 + 客户端）并写入 out_dir。"""
    openssl = _find_openssl()
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"[INFO] 证书输出目录：{out_dir.resolve()}")
    print(f"[INFO] 有效期：{days} 天")
    if extra_hosts:
        print(f"[INFO] 服务器 SAN 附加主机：{extra_hosts}")

    print("\n[1/3] 生成 CA 证书…")
    _make_ca(openssl, out_dir, days)

    print("[2/3] 生成服务器证书…")
    _make_server_cert(openssl, out_dir, extra_hosts, days)

    print("[3/3] 生成客户端证书…")
    _make_client_cert(openssl, out_dir, days)

    _print_cert_summary(out_dir)


def generate_server_cert(
    out_dir: Path,
    extra_hosts: list[str],
    days: int,
) -> None:
    """仅重新生成 server.crt + server.key（复用已有 CA），不触动其他文件。

    适用场景：团队成员 IP 不同，各自运行一次本函数即可，无需重新上传设备证书。
    前提：out_dir 下已有 ca.crt 和 ca.key。
    """
    ca_crt_path = out_dir / "ca.crt"
    ca_key_path = out_dir / "ca.key"

    if not ca_crt_path.exists() or not ca_key_path.exists():
        raise FileNotFoundError(
            f"未找到共享 CA 文件（{ca_crt_path} / {ca_key_path}）。\n"
            "请先由一名团队成员运行完整生成（不带 --server-only）并将 ca.* / client.* 提交到 git。"
        )

    openssl = _find_openssl()

    print(f"[INFO] 复用已有 CA：{ca_crt_path}")
    if extra_hosts:
        print(f"[INFO] 服务器 SAN 附加主机：{extra_hosts}")

    _make_server_cert(openssl, out_dir, extra_hosts, days)

    print(f"\n服务器证书已更新：")
    print(f"  {str(out_dir / 'server.crt'):<50s}  Broker 服务器证书（含本机 IP）")
    print(f"  {str(out_dir / 'server.key'):<50s}  Broker 服务器私钥")
    print()
    print("其他文件（ca.crt / client.crt / client.key）未改动，设备无需重新上传。")
    print()
    print("运行 mTLS 测试：")
    print("  python -X utf8 Protocols/MQTT/mqtt_comparator.py --live --all-modules --ssl")


def _print_cert_summary(out_dir: Path) -> None:
    print("\n生成完成，文件说明：")
    print(f"  {str(out_dir / 'ca.key'):<50s}  CA 私钥      [提交 git，团队共享]")
    print(f"  {str(out_dir / 'ca.crt'):<50s}  CA 根证书    [提交 git，团队共享]")
    print(f"  {str(out_dir / 'client.key'):<50s}  客户端私钥  [提交 git，团队共享]")
    print(f"  {str(out_dir / 'client.crt'):<50s}  客户端证书  [提交 git，团队共享]")
    print(f"  {str(out_dir / 'server.key'):<50s}  服务器私钥  [gitignore，各自生成]")
    print(f"  {str(out_dir / 'server.crt'):<50s}  服务器证书  [gitignore，各自生成，含本机 IP]")
    print()
    print("设备导入（只需做一次，换机器不需要重新上传）：")
    print("  将 ca.crt、client.crt、client.key 导入设备的 MQTT SSL 配置。")
    print()
    print("其他成员使用步骤：")
    print("  git pull  （获取共享 CA + 客户端证书）")
    print("  python -X utf8 Protocols/MQTT/gen_certs.py --server-only")
    print("  # 或指定 IP：")
    print("  python -X utf8 Protocols/MQTT/gen_certs.py --server-only --host 192.168.x.x")
    print()
    print("运行 mTLS 测试：")
    print("  python -X utf8 Protocols/MQTT/mqtt_comparator.py --live --all-modules --ssl")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="生成自签 mTLS 证书（CA + 服务器 + 客户端）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
多人协作说明：
  server.crt 的 SAN 里写有本机 IP，不同电脑的 IP 不同，不能直接共享。
  解决方法：
    第一次（由一人执行）：
      python -X utf8 Protocols/MQTT/gen_certs.py --host <本机IP>
      git add Protocols/MQTT/certs/ca.crt Protocols/MQTT/certs/ca.key
      git add Protocols/MQTT/certs/client.crt Protocols/MQTT/certs/client.key
      git commit -m "add shared MQTT mTLS CA and client certs"
      （server.crt / server.key 已在 .gitignore，不会被提交）

    其他成员（git pull 后执行一次）：
      python -X utf8 Protocols/MQTT/gen_certs.py --server-only
      # 自动检测本机 IP 并写入 server.crt，无需重新上传设备证书
""",
    )
    parser.add_argument(
        "--host",
        metavar="IP_OR_HOST",
        nargs="+",
        default=[],
        help="加入服务器证书 SAN 的 IP 或主机名（不填则自动检测本机 IP）",
    )
    parser.add_argument(
        "--days",
        type=int,
        default=3650,
        help="证书有效期（天），默认 3650（10 年）",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=_DEFAULT_OUT,
        help=f"证书输出目录，默认 {_DEFAULT_OUT}",
    )
    parser.add_argument(
        "--server-only",
        action="store_true",
        help="仅重新生成 server.crt（复用已有 CA），不改动 ca.* / client.*；"
             "团队其他成员换 IP 时使用",
    )
    args = parser.parse_args()

    extra_hosts = list(args.host)
    if not extra_hosts:
        extra_hosts = _detect_local_ips()
        if extra_hosts:
            print(f"[INFO] 未指定 --host，自动检测到本机 IP：{extra_hosts}")
        else:
            print("[WARN] 无法自动检测本机 IP，SAN 仅含 127.0.0.1 / localhost。")

    if args.server_only:
        generate_server_cert(out_dir=args.out, extra_hosts=extra_hosts, days=args.days)
    else:
        generate_certs(out_dir=args.out, extra_hosts=extra_hosts, days=args.days)


if __name__ == "__main__":
    main()
