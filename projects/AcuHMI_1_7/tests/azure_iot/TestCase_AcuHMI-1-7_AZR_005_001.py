"""
FTS编号: FTS_AcuHMI_AZR_005_001
用例标题: 断网 24 小时内重连补传数据完整有序
用例级别: LV1

预置条件: Azure IoT 已正常配置并连接，Event Hub 订阅端就绪

说明: 涉及长时间断网重传，需手动执行，不纳入自动化。
"""

import pytest


class TestCase_AcuHMI_1_7_AZR_005_001:

    @pytest.mark.azure_iot
    @pytest.mark.manual
    def test_reconnect_backfill_within_24h(self):
        pytest.skip("断网重传测试需手动执行，不纳入自动化")
