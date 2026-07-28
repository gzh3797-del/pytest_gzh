from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_005_002
# 用例标题：Alarm Logs 含 Trigger DO/RO 信息列，无数据时显示"-"
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 触发一条配置了 Trigger DO Device / RO 设备的告警，进入 Alarm Logs
#   2. 触发一条未配置 Trigger 设备的告警，进入 Alarm Logs
# 预期结果：
#   1. 对应行 Trigger DO Device / DO / Trigger RO Device / RO 列显示
#      配置的设备及端口名称
#   2. 对应行 Trigger DO Device / DO / Trigger RO Device / RO 列显示 "-"
# 说明：步骤 1 需要网关下挂含 DO/RO 端口的设备（如 AcuIOM），当前测试台架
#      未接入此类设备，本脚本覆盖步骤 2（未配置 Trigger 时显示 "-"）并断言
#      四列存在；步骤 1 待台架具备 DO/RO 设备后补充。

_LABEL = "at_bz005002"
_TRIGGER_COLUMNS = ("Trigger DO Device", "DO", "Trigger RO Device", "RO")


def test_TestCase_ACUHMI17_BZ_005_002(app_page: Page):
    page = app_page
    try:
        # ── Step 2: 触发一条未配置 Trigger DO/RO 的告警 ──
        ha.set_ack_enable(page, True)
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_log_rows(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，Alarm Logs 未在轮询周期内出现记录"

        # 四列均存在
        headers = ha.table_headers(page)
        for col in _TRIGGER_COLUMNS:
            assert col in headers, \
                f"Alarm Logs 应包含 {col!r} 列，实际列: {headers}"

        # 未配置 Trigger 设备的告警行，四列均显示 "-"
        row = ha.data_rows(page, _LABEL).first
        for col in _TRIGGER_COLUMNS:
            val = ha.cell_text(page, row, col)
            assert val == "-", (
                f"未配置 Trigger 设备时 {col!r} 列应显示 '-'，实际: {val!r}"
            )
    finally:
        try:
            ha.ack_alarm(page, _LABEL)
        except Exception:
            pass
        ha.cleanup_test_rules(page)
