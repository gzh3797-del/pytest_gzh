import re
from datetime import datetime, timedelta

from playwright.sync_api import Page

from projects.RPP.tests.Alarm import helpers_alarm as ha


# 用例编号：TestCase_ACUHMI17_BZ_004_003
# 用例标题：确认操作后 Alarm Log 新增确认日志
# 预置条件：
#   1. 设备已正常上电，相关服务正常启动
# 测试步骤：
#   1. 在 Unacknowledged Alarms 页面确认一条告警
#   2. 进入 Alarm Logs 页面查看日志
# 预期结果：
#   2. Alarm Logs 中新增一条确认日志，Ack Status 为 Acknowledged，
#      时间戳与确认操作时间一致

_LABEL = "at_bz004003"


def test_TestCase_ACUHMI17_BZ_004_003(app_page: Page):
    page = app_page
    try:
        # ── 前置: 制造一条未确认告警 ──
        ha.set_ack_enable(page, True)
        ha.trigger_alarm(page, _LABEL)
        assert ha.wait_for_unack(page, _LABEL), \
            f"创建必触发规则 {_LABEL!r} 后，未在轮询周期内出现未确认告警"

        ha.goto_global_alarm(page, "Alarm Logs")
        acked_before = ha.count_rows_by_ack(page, _LABEL, acknowledged=True)

        # ── Step 1: 确认该告警，记录确认时刻 ──
        ack_time = datetime.now()
        ha.ack_alarm(page, _LABEL)

        # ── Step 2: Alarm Logs 新增 Acknowledged 记录 ──
        assert ha.wait_for_log_status(page, _LABEL, acknowledged=True,
                                      min_count=acked_before + 1), (
            f"确认后全局 Alarm Logs 中 {_LABEL!r} 的 Acknowledged 记录数应增加"
            f"（确认前 {acked_before} 条）"
        )

        # 时间戳与确认操作时间一致（可解析时校验在确认时刻 ±15 分钟内，
        # 覆盖网关与本机的小幅时钟偏差）
        rows = ha.data_rows(page, _LABEL)
        latest_ts = ""
        for i in range(rows.count()):
            row = rows.nth(i)
            if ha.ack_status_of_row(row).lower().startswith("unacknowl"):
                continue
            latest_ts = ha.cell_text(page, row, "Timestamp")
            break
        assert latest_ts, "确认日志的 Timestamp 不应为空"
        m = re.search(r"(\d{4})-(\d{2})-(\d{2})\D+(\d{1,2}):(\d{2}):(\d{2})",
                      latest_ts)
        if m:
            logged = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                              int(m.group(4)) % 24, int(m.group(5)),
                              int(m.group(6)))
            # 12 小时制无法直接判 AM/PM，按小时数取最接近的解释
            candidates = [logged, logged + timedelta(hours=12)]
            diff = min(abs((c - ack_time).total_seconds()) for c in candidates)
            assert diff <= 15 * 60, (
                f"确认日志时间戳 {latest_ts!r} 与确认操作时刻 "
                f"{ack_time:%Y-%m-%d %H:%M:%S} 相差超过 15 分钟"
            )
    finally:
        ha.cleanup_test_rules(page)
