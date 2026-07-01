"""
FTS编号: FTS_AcuHMI_AWS_004_007
用例标题: AcuHMI-1-7选择Virtual虚拟设备连接 AWS IoT Core 数据上报
用例级别: LV3

预置条件:
  - 网关已创建虚拟设备，config.yaml 中 virtual_device 字段指定目标设备名
  - 合法证书已配置，MQTT 订阅端就绪

测试步骤（一次浏览器完成）:
  1. 进入 AWS IoT 页面，配置虚拟设备并 Save（同时启动 MQTT 订阅）
  2. 等待虚拟设备按 interval 上报数据
  3. 收到数据后，在同一浏览器导航到 Virtual Devices → 设备 → Reading 页面读取参数值
  4. 比对 MQTT 上报值与 Reading 页面值（误差在 5% 容差内）

预期结果:
  - MQTT 收到虚拟设备数据
  - 上报参数单位与 Reading 页面一致，数值误差在 5% 容差内
"""

import copy
from pathlib import Path

import pytest

from utils.aws_iot_verifier import run_verifier

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent.parent

DEVICE_NAME = ""  # 无特定设备要求，使用虚拟设备


class TestCase_AcuHMI_1_7_AWS_004_007:

    @pytest.mark.aws_iot
    @pytest.mark.slow
    def test_virtual_device_data_vs_reading(self, aws_page, aws_cfg):
        """配置虚拟设备 → MQTT 收数据 → 读 Reading 页面 → run_verifier 比对"""
        aws = aws_cfg["aws_iot"]
        virtual_device = aws.get("virtual_device", "")
        assert virtual_device, "config.yaml 未配置 aws_iot.virtual_device，无法执行本用例"

        aws_page.ensure_enabled()
        aws_page.set_url(aws["url"])
        aws_page.set_interval(aws.get("interval", "30 seconds"))
        aws_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
        aws_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))
        aws_page.select_only_device(virtual_device, exact=True)
        aws_page.configure_all_devices_parameters(checked_only=True)
        aws_page.save()

        # 读取 Reading 页面当前值，作为 stage [4] 的比对基准
        reading_params = aws_page.get_virtual_device_readings(virtual_device)
        assert reading_params, (
            f"Reading 页面未读取到任何参数，设备：{virtual_device!r}。"
            f"请确认网关上该虚拟设备已配置公式参数。"
        )
        print(f"\n[Reading] {virtual_device!r}: {len(reading_params)} 个参数")

        interval_s = int(str(aws.get("interval", "30 seconds")).split()[0])
        timeout = 60 + interval_s * 3

        single_cfg = copy.deepcopy(aws_cfg)
        single_cfg["aws_iot"]["expected_devices"] = [virtual_device]

        ok = run_verifier(
            single_cfg,
            timeout=timeout,
            virtual_readings={virtual_device: reading_params},
        )
        assert ok, f"虚拟设备 {virtual_device!r} 数据与 Reading 页面比对失败，请查看 HTML 报告"
