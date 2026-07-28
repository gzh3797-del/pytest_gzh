import pytest

from playwright.sync_api import Page


# 用例编号：TestCase_ACUHMI17_BZ_006_003
# 用例标题：被监控设备离线时 Unacknowledged Alarms 页面状态合理
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 设备在线时触发一条告警（未确认）
#   2. 使该设备离线
#   3. 查看 Unacknowledged Alarms 页面中该告警条目
# 预期结果：
#   3. 告警条目仍在 Unacknowledged Alarms 列表中，状态展示合理（如标注设备
#      离线）；系统不崩溃，蜂鸣器行为不受异常影响


@pytest.mark.skip(reason="步骤 2 需使被监控设备真实离线（断电/拔网线），"
                         "Web 端关闭轮询开关不等价于设备离线，自动化环境不可控，"
                         "需人工配合执行")
def test_TestCase_ACUHMI17_BZ_006_003(app_page: Page):
    pass
