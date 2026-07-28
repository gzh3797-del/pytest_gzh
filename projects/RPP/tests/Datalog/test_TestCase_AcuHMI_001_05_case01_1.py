from pathlib import Path

import pytest
from playwright.sync_api import Page

from projects.RPP.tests.Datalog import helpers_datalog as hd


# 用例编号：TestCase_AcuHMI_001_05_case01_1
# 用例标题：接入设备 Datalog 栏，通过日志产生时间区间和间隔
#          （1hour、6hours、12hours），可下载日志到本地 PC，导出日志格式为 .csv
# 预置条件：
#   1. AcuHMI 上电
#   2. 设备已接入并在线，且已采集数据超过 1 小时
# 测试步骤：
#   2-7. 选中有 datalog 的时间段，依次选择 1hour/6hours/12hours 间隔，
#        点击导出并检查文件内容和格式
# 预期结果：各间隔导出文件数据时间戳及时间间隔正确
# 设计约束：Log Interval 必须小于数据时长；数据跨度不足的档位后端拒绝导出，
#      属正常防护，本脚本仅测数据跨度足够的档位，不足档位 skip 待补。

_INTERVALS = ("1 hour", "6 hours", "12 hours")


def test_TestCase_AcuHMI_001_05_case01_1(app_page: Page, tmp_path: Path):
    page = app_page
    hd.goto_device_datalog(page)
    start_date, end_date = hd.date_range_values(page)
    assert start_date and end_date, "Data Log 页应有默认时间区间"

    span = hd.current_data_span_seconds(page, tmp_path)
    ok, lack = hd.testable_intervals(_INTERVALS, span)
    if not ok:
        pytest.skip(f"前置不满足：设备现有数据跨度 {span / 3600:.2f}h，"
                    f"不足以测试任一档位 {_INTERVALS}（设计约束：间隔须小于"
                    f"数据时长），待数据累计后重跑")

    for interval in ok:
        hd.select_interval(page, interval)
        stamps, name = hd.download_csv(page, tmp_path)
        assert stamps is not None, \
            f"间隔 {interval} 下载失败（页面提示: {name!r}）"
        assert ".csv" in name, f"导出文件应为 CSV 格式，实际文件名: {name!r}"
        hd.verify_timestamps(stamps, interval, start_date, end_date)

    if lack:
        pytest.skip(f"部分档位因数据跨度（{span / 3600:.2f}h）不足未测，"
                    f"待数据累计后重跑: {lack}（本轮已测通过: {ok}）")
