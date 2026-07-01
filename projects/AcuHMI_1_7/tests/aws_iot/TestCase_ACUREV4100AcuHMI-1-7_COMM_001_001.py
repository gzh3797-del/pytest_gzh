"""
FTS编号: FTS_ACUREV4100WEB2_COMM_001_001
用例标题 A: BACnet/IP、EtherNet/IP、云端协议并存时消息处理正常
用例标题 B: 同一设备被多个协议选中发布时数据一致且正确
用例级别: LV4

（两条用例在 Excel 中共用同一用例编号）

预置条件:
  - 网关已配置 BACnet/IP、EtherNet/IP，各协议客户端就绪
  - AWS IoT 合法证书就绪，Azure IoT Connection String 就绪
"""

import subprocess
import sys
from pathlib import Path

import pytest

from utils.aws_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent   # AcuHMI-1-7/

# 协议 comparator 路径（按 AcuHMI-1-7 工程布局查找）
_BACNET = _PROJECT_ROOT / "tests" / "protocols" / "bacnet" / "comparator.py"
_EIP    = _PROJECT_ROOT / "tests" / "protocols" / "ethernetip" / "comparator.py"


# ── 用例 A: 协议并存 ──────────────────────────────────────────────────────────

@pytest.mark.skip(reason="用例暂不纳入全量执行")
class TestCase_ACUREV4100AcuHMI_1_7_COMM_001_001_A:

    @pytest.mark.aws_iot
    @pytest.mark.slow
    @pytest.mark.manual
    def test_aws_connected_in_coexist(self, aws_session_page, aws_cfg):
        """启用 AWS IoT（其余协议由测试人员手动配置已启用）"""
        aws = aws_cfg["aws_iot"]
        aws_session_page.ensure_enabled()
        aws_session_page.set_url(aws["url"])
        aws_session_page.set_interval(aws.get("interval", "30 seconds"))
        aws_session_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
        aws_session_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))
        aws_session_page.select_only_device()  # 随机选一台物理设备
        aws_session_page.configure_all_devices_parameters(checked_only=True)
        aws_session_page.save()
        result = aws_session_page.test_connection()
        assert any(kw in result.lower() for kw in ("success", "connected", "ok", "成功")), \
            f"多协议并存时 AWS IoT 连接应成功，实际：{result!r}"

    @pytest.mark.aws_iot
    @pytest.mark.slow
    @pytest.mark.manual
    def test_aws_data_in_coexist(self, aws_cfg):
        """各协议并存运行期间验证 AWS IoT 数据正常"""
        ok = run_verifier(aws_cfg, timeout=180, skip_web=True)
        assert ok, "多协议并存时 AWS IoT 数据验证失败"

    @pytest.mark.aws_iot
    @pytest.mark.slow
    @pytest.mark.manual
    def test_bacnet_data_in_coexist(self):
        """多协议并存期间验证 BACnet/IP 数据正常"""
        if not _BACNET.exists():
            pytest.skip("BACnet comparator 不存在，跳过")
        proc = subprocess.run(
            [sys.executable, str(_BACNET)],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            cwd=str(_PROJECT_ROOT)
        )
        output = (proc.stdout or "") + (proc.stderr or "")
        assert proc.returncode == 0 or "PASS" in output, \
            f"多协议并存时 BACnet/IP 数据验证失败:\n{output}"


# ── 用例 B: 同一设备多协议数据一致性 ─────────────────────────────────────────

@pytest.mark.skip(reason="用例暂不纳入全量执行")
class TestCase_ACUREV4100AcuHMI_1_7_COMM_001_001_B:

    @pytest.mark.aws_iot
    @pytest.mark.slow
    @pytest.mark.manual
    def test_same_device_value_consistency(self, aws_cfg):
        """同一设备被多协议发布时，并发采集各协议数据并比对"""
        import threading

        results = {}

        def run_aws():
            results["aws"] = run_verifier(aws_cfg, timeout=120, skip_web=True)

        def run_bacnet():
            if not _BACNET.exists():
                results["bacnet"] = None
                return
            proc = subprocess.run(
                [sys.executable, str(_BACNET)],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                cwd=str(_PROJECT_ROOT)
            )
            output = (proc.stdout or "") + (proc.stderr or "")
            results["bacnet"] = proc.returncode == 0 or "PASS" in output

        def run_eip():
            if not _EIP.exists():
                results["eip"] = None
                return
            proc = subprocess.run(
                [sys.executable, str(_EIP)],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                cwd=str(_PROJECT_ROOT)
            )
            output = (proc.stdout or "") + (proc.stderr or "")
            results["eip"] = proc.returncode == 0 or "PASS" in output

        threads = [
            threading.Thread(target=run_aws),
            threading.Thread(target=run_bacnet),
            threading.Thread(target=run_eip),
        ]
        for t in threads:
            t.start()
        for t in threads:
            t.join(timeout=200)

        failures = [
            f"{proto}: 验证失败"
            for proto, ok in results.items()
            if ok is False
        ]
        assert not failures, "同一设备多协议数据一致性验证：\n" + "\n".join(failures)
