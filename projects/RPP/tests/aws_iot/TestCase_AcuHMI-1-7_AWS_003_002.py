"""
FTS编号: FTS_WEB2_AWS_003_002
用例标题: 非法证书或错误密钥连接失败
用例级别: LV2

预置条件:
  - 网关已登录，合法 cert/key 存放在 tests/protocols/aws_iot/certs/ 目录

测试步骤（Enable 一次，依次验证 4 种无效凭证组合，不中途 Disable）:
  场景 A: 合法 cert + 错误 key
  场景 B: 错误 cert + 合法 key
  场景 C: 错误 cert + 错误 key
  场景 D: .txt 格式或内容损坏文件

预期结果:
  - 以上 4 种情况 Test Connection 均失败
  - 页面无崩溃，可重新操作
"""

import os
from pathlib import Path

import pytest

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent  # AcuHMI-1-7/
_CERTS        = _THIS_DIR / "certs"


def _valid_cert_path(aws_cfg):
    return str(_PROJECT_ROOT / aws_cfg["aws_iot"]["cert_file"])


def _valid_key_path(aws_cfg):
    return str(_PROJECT_ROOT / aws_cfg["aws_iot"]["key_file"])


def _make_invalid_pem(tmp_dir, name, pem_type="CERTIFICATE"):
    path = os.path.join(str(tmp_dir), name)
    with open(path, "w") as f:
        f.write(f"-----BEGIN {pem_type}-----\nINVALID_CONTENT_HERE\n-----END {pem_type}-----\n")
    return path


def _make_txt_file(tmp_dir):
    path = os.path.join(str(tmp_dir), "dummy.txt")
    with open(path, "w") as f:
        f.write("this is not a certificate file")
    return path


def _connection_failed(result):
    return any(kw in result.lower() for kw in ("fail", "error", "invalid", "连接失败", "未检测到"))


class TestCase_RPP_AWS_003_002:

    @pytest.mark.aws_iot
    def test_invalid_cert_connection_fails(self, aws_page, aws_cfg, tmp_path):
        """4 种无效凭证组合均应导致 Test Connection 失败。"""
        aws_page.ensure_enabled()
        aws_page.set_url(aws_cfg["aws_iot"]["url"])
        aws_page.select_only_device("")

        failures = []

        # 场景 A：合法 cert + 错误 key（用 RSA PRIVATE KEY 头确保网关不误判为合法证书）
        aws_page.upload_cert_file(_valid_cert_path(aws_cfg))
        aws_page.upload_key_file(_make_invalid_pem(tmp_path, "invalid_key.pem", "RSA PRIVATE KEY"))
        aws_page.save()
        result = aws_page.test_connection()
        if not _connection_failed(result):
            failures.append(f"场景A（合法cert+错误key）应失败，实际：{result!r}")

        # 场景 B：错误 cert + 合法 key
        aws_page.upload_cert_file(_make_invalid_pem(tmp_path, "invalid_cert.pem"))
        aws_page.upload_key_file(_valid_key_path(aws_cfg))
        aws_page.save()
        result = aws_page.test_connection()
        if not _connection_failed(result):
            failures.append(f"场景B（错误cert+合法key）应失败，实际：{result!r}")

        # 场景 C：错误 cert + 错误 key
        aws_page.upload_cert_file(_make_invalid_pem(tmp_path, "invalid_cert2.pem"))
        aws_page.upload_key_file(_make_invalid_pem(tmp_path, "invalid_key2.pem"))
        aws_page.save()
        result = aws_page.test_connection()
        if not _connection_failed(result):
            failures.append(f"场景C（双错误文件）应失败，实际：{result!r}")

        # 场景 D：.txt 格式文件
        txt = _make_txt_file(tmp_path)
        aws_page.upload_cert_file(txt)
        aws_page.upload_key_file(txt)
        aws_page.save()
        result = aws_page.test_connection()
        if not _connection_failed(result):
            failures.append(f"场景D（.txt格式文件）应失败，实际：{result!r}")

        assert not failures, "以下场景未按预期失败：\n" + "\n".join(failures)
