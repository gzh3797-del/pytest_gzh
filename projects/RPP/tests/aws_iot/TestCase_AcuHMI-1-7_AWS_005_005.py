"""
FTS编号: FTS_AcuHMI_AWS_005_005
用例标题: 重连后补传与实时采集同时进行互不阻塞
用例级别: LV3

预置条件: AWS IoT 正常连接，MQTT 订阅端持续监听

说明: 涉及长时间断网重传，需手动执行，不纳入自动化。
"""

import pytest


class TestCase_RPP_AWS_005_005:

    @pytest.mark.aws_iot
    @pytest.mark.manual
    def test_backfill_nonblocking_with_realtime(self):
        pytest.skip("断网重传测试需手动执行，不纳入自动化")
