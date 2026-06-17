# -*- coding: utf-8 -*-
"""
gen_certs.py — 生成自签 mTLS 证书（CA + 服务器 + 客户端）

使用 cryptography 库（纯 Python，不依赖 openssl CLI）。
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
import socket
import sys
from datetime import datetime, timezone
from pathlib import Path

# ── 依赖检查 ──────────────────────────────────────────────────────────────────
try:
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import NameOID
except ImportError:
    print(
        "[ERROR] 缺少 cryptography 库，请先安装：\n"
        "  pip install cryptography\n"
        "安装后重新运行此脚本。"
    )
    sys.exit(1)


# ── 默认输出目录（脚本所在目录下的 certs/） ──────────────────────────────────
_DEFAULT_OUT = Path(__file__).parent / "certs"


def _detect_local_ips() -> list[str]:
    """检测本机所有非回环 IPv4 地址（不依赖第三方库）。"""
    ips: set[str] = set()
    # UDP 路由技巧：让系统路由表选出主出口 IP，不发送实际数据包
    for target in ("8.8.8.8", "192.168.0.1", "10.0.0.1"):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect((target, 80))
            ips.add(s.getsockname()[0])
            s.close()
            break
        except OSError:
            pass
    # 通过主机名枚举所有绑定 IP
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if not ip.startswith("127."):
                ips.add(ip)
    except OSError:
        pass
    return sorted(ips)


def _make_key() -> rsa.RSAPrivateKey:
    """生成 2048-bit RSA 私钥。"""
    return rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
    )


def _save_key(key: rsa.RSAPrivateKey, path: Path) -> None:
    path.write_bytes(
        key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        )
    )


def _save_cert(cert: x509.Certificate, path: Path) -> None:
    path.write_bytes(cert.public_bytes(serialization.Encoding.PEM))


def _make_ca(days: int) -> tuple[rsa.RSAPrivateKey, x509.Certificate]:
    """生成自签 CA 私钥和证书。"""
    key = _make_key()
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, "MQTT-Test-CA"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "AutoTest"),
    ])
    now = datetime.now(timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now)
        .not_valid_after(now.replace(year=now.year + days // 365) if days >= 365
                         else now.replace(day=now.day + days))
        .add_extension(x509.BasicConstraints(ca=True, path_length=None), critical=True)
        .add_extension(
            x509.SubjectKeyIdentifier.from_public_key(key.public_key()),
            critical=False,
        )
        .sign(key, hashes.SHA256())
    )
    return key, cert


def _not_after(now: datetime, days: int) -> datetime:
    """计算证书到期时间（简单按天数偏移，避免 timedelta 对 leap-year 的问题）。"""
    try:
        return now.replace(year=now.year + days // 365,
                           month=now.month,
                           day=now.day)
    except ValueError:
        # 例如 2020-02-29 + 1 年 → 2021-02-28
        import calendar
        year  = now.year + days // 365
        month = now.month
        day   = min(now.day, calendar.monthrange(year, month)[1])
        return now.replace(year=year, month=month, day=day)


def _make_server_cert(
    ca_key: rsa.RSAPrivateKey,
    ca_cert: x509.Certificate,
    extra_hosts: list[str],
    days: int,
) -> tuple[rsa.RSAPrivateKey, x509.Certificate]:
    """生成服务器私钥和证书（含 SAN）。"""
    key = _make_key()
    subject = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, "mqtt-broker"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "AutoTest"),
    ])
    now = datetime.now(timezone.utc)

    # SAN 必须始终包含 127.0.0.1 和 localhost
    san_entries: list[x509.GeneralName] = [
        x509.IPAddress(ipaddress.IPv4Address("127.0.0.1")),
        x509.DNSName("localhost"),
    ]
    for h in extra_hosts:
        try:
            san_entries.append(x509.IPAddress(ipaddress.ip_address(h)))
        except ValueError:
            san_entries.append(x509.DNSName(h))

    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(ca_cert.subject)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now)
        .not_valid_after(_not_after(now, days))
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(x509.SubjectAlternativeName(san_entries), critical=False)
        .add_extension(
            x509.SubjectKeyIdentifier.from_public_key(key.public_key()),
            critical=False,
        )
        .sign(ca_key, hashes.SHA256())
    )
    return key, cert


def _make_client_cert(
    ca_key: rsa.RSAPrivateKey,
    ca_cert: x509.Certificate,
    days: int,
) -> tuple[rsa.RSAPrivateKey, x509.Certificate]:
    """生成客户端私钥和证书（用于 paho / 设备侧 mTLS 认证）。"""
    key = _make_key()
    subject = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, "mqtt-client"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "AutoTest"),
    ])
    now = datetime.now(timezone.utc)

    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(ca_cert.subject)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now)
        .not_valid_after(_not_after(now, days))
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(
            x509.SubjectKeyIdentifier.from_public_key(key.public_key()),
            critical=False,
        )
        .sign(ca_key, hashes.SHA256())
    )
    return key, cert


def generate_certs(
    out_dir: Path,
    extra_hosts: list[str],
    days: int,
) -> None:
    """生成全套 mTLS 证书（CA + 服务器 + 客户端）并写入 out_dir。

    多人协作时建议只做一次，把 ca.crt / ca.key / client.crt / client.key
    提交到 git；各自用 generate_server_cert() 重生成含本机 IP 的 server.crt。
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"[INFO] 证书输出目录：{out_dir.resolve()}")
    print(f"[INFO] 有效期：{days} 天")
    if extra_hosts:
        print(f"[INFO] 服务器 SAN 附加主机：{extra_hosts}")

    print("\n[1/3] 生成 CA 证书…")
    ca_key, ca_cert = _make_ca(days)

    print("[2/3] 生成服务器证书…")
    srv_key, srv_cert = _make_server_cert(ca_key, ca_cert, extra_hosts, days)

    print("[3/3] 生成客户端证书…")
    cli_key, cli_cert = _make_client_cert(ca_key, ca_cert, days)

    _save_key(ca_key,    out_dir / "ca.key")
    _save_cert(ca_cert,  out_dir / "ca.crt")
    _save_key(srv_key,   out_dir / "server.key")
    _save_cert(srv_cert, out_dir / "server.crt")
    _save_key(cli_key,   out_dir / "client.key")
    _save_cert(cli_cert, out_dir / "client.crt")

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

    from cryptography.hazmat.primitives.serialization import load_pem_private_key
    from cryptography import x509 as _x509

    ca_key  = load_pem_private_key(ca_key_path.read_bytes(), password=None)
    ca_cert = _x509.load_pem_x509_certificate(ca_crt_path.read_bytes())

    # 计算剩余有效期（以 CA 到期时间为上限）
    from datetime import timezone as _tz
    ca_remaining = (ca_cert.not_valid_after_utc - __import__('datetime').datetime.now(_tz.utc)).days
    srv_days     = min(days, max(ca_remaining, 1))

    print(f"[INFO] 复用已有 CA：{ca_crt_path}")
    print(f"[INFO] 服务器证书有效期：{srv_days} 天")
    if extra_hosts:
        print(f"[INFO] 服务器 SAN 附加主机：{extra_hosts}")

    srv_key, srv_cert = _make_server_cert(ca_key, ca_cert, extra_hosts, srv_days)

    _save_key(srv_key,   out_dir / "server.key")
    _save_cert(srv_cert, out_dir / "server.crt")

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
