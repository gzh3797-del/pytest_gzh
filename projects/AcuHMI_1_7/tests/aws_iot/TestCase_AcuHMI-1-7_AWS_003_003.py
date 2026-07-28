"""
FTS编号: FTS_WEB2_AWS_003_003
用例标题: AWS IoT Core 上报参数数量及单位与模板一致性验证（全设备）
用例级别: LV2（slow）

预置条件:
  - 网关已登录，MQTT 订阅端就绪
  - tests/protocols/aws_iot/certs/ 目录下存放合法的证书和密钥文件

测试步骤:
  1. Enable AWS IoT，填写合法 URL、Interval
  2. 上传合法 cert 文件和 key 文件
  3. 全选所有设备（物理 + 虚拟）并配置参数，点击 Save
  4. 读取虚拟设备 Reading 页的参数/单位作为期望模板
  5. 订阅 AWS IoT Core，验证各设备上报的参数数量和单位：
     - 物理设备：与 Excel 模板 MQTT 列比对
     - 虚拟设备：与 UI Virtual Device Reading 页比对

预期结果:
  - 各设备上报的参数数量与对应模板一致
  - 各设备上报的参数单位与对应模板一致
"""

import logging
from pathlib import Path

import pytest

from utils.aws_iot_verifier import run_verifier_params_only

_THIS_DIR     = Path(__file__).resolve().parent
_PROJECT_ROOT = _THIS_DIR.parent.parent  # AcuHMI-1-7/

log = logging.getLogger(__name__)


class TestCase_AcuHMI_1_7_AWS_003_003:

    @pytest.mark.aws_iot
    @pytest.mark.slow
    def test_params_and_units_match_template(self, aws_page, aws_cfg):
        """全选设备 → 读取虚拟设备 Reading 页模板 → 验证所有设备上报参数数量和单位一致"""
        aws = aws_cfg["aws_iot"]

        # ── 1. 配置 AWS IoT ──────────────────────────────────────────────────
        aws_page.ensure_enabled()
        aws_page.set_url(aws["url"])
        aws_page.set_interval(aws.get("interval", "30 seconds"))
        aws_page.upload_cert_file(str(_PROJECT_ROOT / aws["cert_file"]))
        aws_page.upload_key_file(str(_PROJECT_ROOT / aws["key_file"]))

        # ── 2. 全选所有设备（物理 + 虚拟）──────────────────────────────────
        all_devices = aws_page.get_all_devices_from_table()
        if not all_devices:
            pytest.skip("设备表格为空，无法执行验证")
        all_names    = [d["name"] for d in all_devices]
        virtual_devs = [d["name"] for d in all_devices if d["is_virtual"]]
        log.info("全选设备：%s（虚拟：%s）", all_names, virtual_devs)

        aws_page.select_devices(all_names)
        aws_page.configure_all_devices_parameters(checked_only=True)
        aws_page.save()

        # ── 3. 读取虚拟设备 Reading 页参数/单位 ─────────────────────────────
        extra_templates: dict[str, dict[str, str]] = {}
        for vdev in virtual_devs:
            try:
                readings = aws_page.get_virtual_device_readings(vdev)
                # readings: {param_name: {"value": str, "unit": str}}
                extra_templates[vdev] = {p: info["unit"] for p, info in readings.items()}
                log.info("虚拟设备 %r Reading 页读取 %d 个参数", vdev, len(extra_templates[vdev]))
            except Exception as exc:
                log.warning("读取虚拟设备 %r Reading 页失败：%s", vdev, exc)

        # ── 4. 订阅验证（物理设备对 Excel 模板，虚拟设备对 Reading 页模板）──
        ok = run_verifier_params_only(
            aws_cfg,
            timeout=120,
            extra_templates=extra_templates if extra_templates else None,
        )
        assert ok, "参数数量或单位验证失败，请查看 reports/ 下的 HTML 报告"
