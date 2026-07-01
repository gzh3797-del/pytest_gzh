"""
FTS编号: FTS_AcuHMI_AWS_005_003
用例标题: 断网超 72 小时边界行为与预期一致
用例级别: LV2

预置条件: AWS IoT 已正常配置并连接，已了解设备缓存上限（72 小时为预期边界）

说明: 涉及长时间断网重传，需手动执行，不纳入自动化。
"""

import pytest


class TestCase_AcuHMI_1_7_AWS_005_003:

    @pytest.mark.aws_iot
    @pytest.mark.manual
    def test_boundary_behavior_over_72h_offline(self):
        pytest.skip("断网重传测试需手动执行，不纳入自动化")
