"""接入设备 Data Log（Datalog 下载）用例组共享操作封装。

页面：设备详情 → Logs → Data Log（/#/physicalDevices/deviceDetails/<id>/3:1）。
2026-07-17 实测要点：
- 控件：数据源 el-select（'Current'）、间隔 el-select（1 minute ~ 1 month 共 11 档）、
  daterange 日期区间（默认 昨天~今天）、Download / Clear Logs 按钮；
- **日期面板只允许选择有 datalog 数据的日期**，无数据日（含未来）全部 disabled；
- 下载产物是 gzip 压缩的 CSV（文件名 *.csv.gz），首列 TimeTag，
  时间戳 ISO 格式如 2026-07-16T15:29:00+0800。
"""
from __future__ import annotations

import gzip
from datetime import datetime
from pathlib import Path

from playwright.sync_api import Page

from projects.RPP.tests.Alarm import config_alarm as cfg
from projects.RPP.tests.Alarm import helpers_alarm as ha

# 间隔下拉选项 → 秒数（'1 month' 因月长不定单独处理）
INTERVAL_SECONDS = {
    "1 minute": 60,
    "5 minutes": 300,
    "10 minutes": 600,
    "15 minutes": 900,
    "30 minutes": 1800,
    "1 hour": 3600,
    "6 hours": 21600,
    "12 hours": 43200,
    "1 day": 86400,
    "7 days": 604800,
    "1 month": None,
}


def goto_device_datalog(page: Page,
                        device_name: str = cfg.TRIGGER_DEVICE) -> None:
    """进入 Physical Devices → 设备详情 → Logs → Data Log。"""
    ha.ensure_devices_module(page)
    page.locator(".left-nav-item").filter(has_text="Physical Devices").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)
    row = page.locator("tr.el-table__row").filter(has_text=device_name).first
    row.wait_for(timeout=10_000)
    row.locator("td").first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)
    item = page.get_by_role("menuitem", name="Data Log")
    logs_sub = page.locator("li.el-sub-menu").filter(has_text="Logs")
    if logs_sub.count() > 0 and (item.count() == 0 or not item.first.is_visible()):
        logs_sub.first.click()
        page.wait_for_timeout(500)
        item = page.get_by_role("menuitem", name="Data Log")
    item.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def select_interval(page: Page, interval: str) -> None:
    """选择时间间隔（页面第 2 个可见 el-select；第 1 个为数据源 'Current'）。"""
    assert interval in INTERVAL_SECONDS, f"未知间隔: {interval!r}"
    page.locator(".el-select:visible").nth(1).click()
    page.wait_for_timeout(400)
    page.locator("li.el-select-dropdown__item:visible").filter(
        has_text=interval).first.click()
    page.wait_for_timeout(300)


def date_range_values(page: Page) -> tuple[str, str]:
    """当前区间输入框的起止日期（'YYYY-MM-DD'）。"""
    return (page.locator("input.el-range-input").nth(0).input_value(),
            page.locator("input.el-range-input").nth(1).input_value())


def panel_day_state(page: Page) -> dict:
    """打开日期面板，返回左/右月历的可选与禁选日数量及可选日列表，然后关闭。"""
    page.locator("input.el-range-input").nth(0).click()
    page.wait_for_timeout(600)
    panel = page.locator(".el-picker-panel:visible").first
    state = {}
    for side in ("is-left", "is-right"):
        content = panel.locator(f".el-picker-panel__content.{side}").first
        avail = content.locator("td.available")
        state[side] = {
            "available": [avail.nth(i).inner_text().strip()
                          for i in range(avail.count())],
            "disabled": content.locator("td.disabled").count(),
        }
    page.keyboard.press("Escape")  # 只读未改选，Escape 安全关闭
    page.wait_for_timeout(300)
    return state


def download_csv(page: Page, save_dir: Path) -> tuple[list[datetime] | None, str]:
    """点击 Download 并解析产物。

    返回 (时间戳列表, 文件名)；产品判无数据不产生下载时返回 (None, 页面提示文本)。
    """
    try:
        with page.expect_download(timeout=30_000) as dl_info:
            page.get_by_role("button", name="Download").click()
    except Exception:
        msg = ""
        toast = page.locator(".el-message")
        if toast.count() > 0:
            msg = toast.first.inner_text().strip()
        return None, msg
    download = dl_info.value
    name = download.suggested_filename
    path = save_dir / name
    download.save_as(str(path))
    if name.endswith(".gz"):
        with gzip.open(path, "rt", encoding="utf-8-sig", errors="replace") as f:
            lines = f.read().splitlines()
    else:
        lines = path.read_text(encoding="utf-8-sig",
                               errors="replace").splitlines()
    assert lines and lines[0].startswith("TimeTag"), \
        f"CSV 首列表头应为 TimeTag，实际首行: {lines[0][:80]!r}"
    stamps = []
    for line in lines[1:]:
        cell = line.split(",", 1)[0].strip()
        if cell:
            stamps.append(datetime.strptime(cell, "%Y-%m-%dT%H:%M:%S%z"))
    return stamps, name


def verify_timestamps(stamps: list[datetime], interval: str,
                      start_date: str, end_date: str) -> None:
    """校验导出数据：时间戳落在所选区间内、相邻间隔与所选档位一致。

    数据可能存在采集中断（缺采样点），故间隔校验为：所有相邻差值都是
    间隔的正整数倍，且最小差值恰等于间隔；数据不足 2 行时跳过间隔校验
    （历史数据长度不够大间隔档位，属环境限制而非产品缺陷）。
    """
    assert stamps, "导出文件应包含数据行"
    for ts in stamps:
        day = ts.strftime("%Y-%m-%d")
        assert start_date <= day <= end_date, (
            f"数据时间戳 {ts} 超出所选区间 [{start_date}, {end_date}]"
        )
    if len(stamps) < 2:
        return
    seconds = INTERVAL_SECONDS[interval]
    diffs = [(b - a).total_seconds()
             for a, b in zip(stamps, stamps[1:])]
    if seconds is None:  # 1 month：月长 28~31 天
        for d in diffs:
            assert 28 * 86400 <= d <= 32 * 86400, \
                f"1 month 间隔的相邻时间差异常: {d}s"
        return
    for d in diffs:
        assert d > 0 and d % seconds == 0, (
            f"相邻时间差 {d}s 不是所选间隔 {interval}（{seconds}s）的整数倍"
        )
    assert min(diffs) == seconds, (
        f"最小相邻时间差 {min(diffs)}s 应等于所选间隔 {interval}（{seconds}s）"
    )


def current_data_span_seconds(page: Page, save_dir: Path) -> float:
    """用 1 minute 档下载一次，返回设备现有数据的时间跨度（秒）。

    设计约束：数据时长必须大于所选 Log Interval，跨度不足的档位选了会被
    后端拒绝（HTTP 500），属正常防护而非缺陷。各下载用例先调本函数，
    只对"数据跨度 ≥ 档位间隔"的档位做下载校验。
    """
    select_interval(page, "1 minute")
    stamps, _ = download_csv(page, save_dir)
    if not stamps or len(stamps) < 2:
        return 0.0
    return (stamps[-1] - stamps[0]).total_seconds()


def testable_intervals(intervals: tuple[str, ...],
                       span_seconds: float) -> tuple[list[str], list[str]]:
    """按设计约束把档位分为（可测, 数据不足跳过）两组。1 month 按 28 天算。"""
    ok, lack = [], []
    for iv in intervals:
        seconds = INTERVAL_SECONDS[iv] or 28 * 86400
        (ok if span_seconds >= seconds else lack).append(iv)
    return ok, lack


def clear_logs(page: Page) -> None:
    page.get_by_role("button", name="Clear Logs").click()
    page.wait_for_timeout(400)
    ha._confirm_dialog(page)
    page.wait_for_timeout(1000)
