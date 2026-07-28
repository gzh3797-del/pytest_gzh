from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_003_002
# 用例标题：Disable 状态下 Ack Status 列隐藏
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 将 Alarm Acknowledgement Enable 设为关闭
#   2. 触发一条告警，进入 Alarm Logs 页面
#   3. 查看列表列头
# 预期结果：
#   3. Alarm Logs 中 Ack Status 列不显示，页面无确认相关入口

_LABEL = "at_bz003002"


def test_TestCase_ACUHMI17_BZ_003_002(app_page: Page):
    page = app_page
    try:
        # ── Step 1: 确认开关置为 Disable ──
        ha.set_ack_enable(page, False)

        # ── Step 2: 触发一条告警并等待其进入 Alarm Logs ──
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_log_rows(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，Alarm Logs 未在轮询周期内出现记录"

        # ── Step 3: Ack Status 列不显示 ──
        headers = ha.table_headers(page)
        assert "Ack Status" not in headers, (
            f"Disable 状态下 Alarm Logs 不应显示 Ack Status 列，实际列: {headers}"
        )
    finally:
        # 恢复默认 Enable 并清理规则
        try:
            ha.set_ack_enable(page, True)
        except Exception:
            pass
        try:
            ha.ack_alarm(page, _LABEL)
        except Exception:
            pass
        ha.cleanup_test_rules(page)
