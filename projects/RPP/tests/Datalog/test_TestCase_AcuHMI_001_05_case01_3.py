from pathlib import Path

import pytest
from playwright.sync_api import Page

from projects.RPP.tests.Datalog import helpers_datalog as hd


# 用例编号：TestCase_AcuHMI_001_05_case01_3
# 用例标题：接入设备 Datalog 栏，通过日志产生时间区间（部分区间没有 datalog）
#          和间隔（10mins），可下载日志到本地 PC，导出日志格式为 .csv
# 预置条件：
#   1. AcuHMI 上电
#   2. 设备已接入并在线，且已采集数据超过 10mins
# 测试步骤（用例原文）：
#   2-7. 分别选中"起止都无 datalog / 起有止无 / 起无止有"的时间段，
#        间隔 1mins，导出并检查文件
# 预期结果：导出文件数据时间戳及时间间隔正确
# 实装差异说明（2026-07-17 实测）：日期面板**只允许选择有 datalog 数据的
#   日期**，无数据日（含未来日期）全部 disabled——产品从交互上禁止构造
#   "无数据区间"，用例原文的三种区间均无法选择。本脚本改为验证该禁选
#   防护本身 + 可选区间（起止均有数据）导出正确，是否接受该实装方式
#   由需求方判定。


def test_TestCase_AcuHMI_001_05_case01_3(app_page: Page, tmp_path: Path):
    page = app_page
    hd.goto_device_datalog(page)

    # ── 实装防护验证: 无数据日期在面板中禁选 ──
    state = hd.panel_day_state(page)
    left_avail = state["is-left"]["available"]
    assert left_avail, "面板应存在可选（有 datalog 数据）的日期"
    assert state["is-left"]["disabled"] > 0, \
        "面板应将无 datalog 数据的日期置为禁选"
    assert not state["is-right"]["available"], (
        "下月（未来）日期应全部禁选，实际可选: "
        f"{state['is-right']['available']}"
    )

    # ── 可执行部分: 起止均有数据的区间 + 1 minute 导出正确 ──
    start_date, end_date = hd.date_range_values(page)
    span = hd.current_data_span_seconds(page, tmp_path)
    if span < hd.INTERVAL_SECONDS["1 minute"]:
        pytest.skip(f"前置不满足：设备现有数据跨度 {span:.0f}s 不足 1 分钟"
                    "（如刚执行过 Clear Logs），待数据累计后重跑")
    hd.select_interval(page, "1 minute")
    stamps, name = hd.download_csv(page, tmp_path)
    assert stamps is not None, f"下载失败（页面提示: {name!r}）"
    assert ".csv" in name, f"导出文件应为 CSV 格式，实际文件名: {name!r}"
    hd.verify_timestamps(stamps, "1 minute", start_date, end_date)
