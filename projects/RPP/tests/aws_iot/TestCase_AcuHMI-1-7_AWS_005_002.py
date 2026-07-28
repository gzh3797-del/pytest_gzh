"""
FTS编号: FTS_AcuHMI_AWS_005_002
用例标题: 断网 72 小时内重连补传完整有序
用例级别: LV2

预置条件: AWS IoT 已正常配置并连接，网关有足够存储空间缓存 72 小时数据

说明: 涉及长时间断网重传，需手动执行，不纳入自动化。
"""

import pytest


class TestCase_RPP_AWS_005_002:

    @pytest.mark.aws_iot
    @pytest.mark.manual
    def test_reconnect_backfill_within_72h(self):
        pytest.skip("断网重传测试需手动执行，不纳入自动化")
