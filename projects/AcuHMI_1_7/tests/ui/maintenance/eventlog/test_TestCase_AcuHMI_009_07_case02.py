import pytest
import warnings
import csv
import io
import os
import tarfile
import tempfile
from datetime import date, timedelta, datetime
from playwright.sync_api import expect
from projects.AcuHMI_1_7.settings import BASE_URL
from projects.AcuHMI_1_7.pages.login_page import LoginPage


def _nav_to_maintenance(page, submenu: str):
    if "/maintenance" not in page.url:
        if not any(s in page.url for s in [
            "/systemSettings", "/userManagement", "/protocols",
            "/templates", "/firmwareUpdate", "/diagnostics",
        ]):
            page.locator("header span").filter(has_text="AcuHMI").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Maintenance").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    page.get_by_role("menuitem", name=submenu).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(500)


def _collect_and_verify_all_pages(page, start_dt, end_dt, expected_level, start_date, end_date):
    """遍历所有筛选结果分页，校验 Level 和 Timestamp，返回 (总行数, [(ts, level, msg)])。"""
    all_rows = []
    total = 0
    page_num = 1
    while True:
        rows = page.locator(".el-table__body tr.el-table__row").all()
        for idx, row in enumerate(rows, start=total + 1):
            cells = row.locator("td").all()
            timestamp_text = cells[0].inner_text().strip() if len(cells) > 0 else ""
            level_text     = cells[1].inner_text().strip() if len(cells) > 1 else ""
            message_text   = cells[3].inner_text().strip() if len(cells) > 3 else ""

            assert level_text == expected_level, \
                f"第{idx}行（第{page_num}页）Level='{level_text}'，期望 '{expected_level}'"
            try:
                row_dt = datetime.strptime(timestamp_text, "%Y-%m-%d %H:%M:%S")
                assert start_dt <= row_dt <= end_dt, \
                    f"第{idx}行（第{page_num}页）时间戳 '{timestamp_text}' 不在 {start_date} ~ {end_date} 范围内"
            except ValueError:
                pytest.fail(f"第{idx}行（第{page_num}页）时间戳格式无法解析: '{timestamp_text}'")
            all_rows.append((timestamp_text, level_text, message_text))

        total += len(rows)
        next_btn = page.locator(".el-pagination .btn-next")
        if next_btn.count() == 0 or next_btn.is_disabled():
            break
        next_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
        page_num += 1
    return total, all_rows


def _collect_all_pages_unfiltered(page):
    """遍历无筛选状态下的所有分页，收集全量数据。返回 (总行数, [(ts, level, msg)])。"""
    all_rows = []
    total = 0
    while True:
        rows = page.locator(".el-table__body tr.el-table__row").all()
        for row in rows:
            cells = row.locator("td").all()
            ts  = cells[0].inner_text().strip() if len(cells) > 0 else ""
            lvl = cells[1].inner_text().strip() if len(cells) > 1 else ""
            msg = cells[3].inner_text().strip() if len(cells) > 3 else ""
            all_rows.append((ts, lvl, msg))
        total += len(rows)
        next_btn = page.locator(".el-pagination .btn-next")
        if next_btn.count() == 0 or next_btn.is_disabled():
            break
        next_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    return total, all_rows


def _read_all_export_rows(download):
    """从导出 tar 文件中读取 CSV 全部行。返回 [(time_local, level, message)]。"""
    tmp_path = os.path.join(tempfile.gettempdir(), download.suggested_filename)
    download.save_as(tmp_path)
    with tarfile.open(tmp_path, "r:") as tar:
        csv_member = next((m for m in tar.getmembers() if m.name.endswith(".csv")), None)
        assert csv_member is not None, "导出 tar 文件中未找到 CSV"
        raw = tar.extractfile(csv_member).read()
    reader = csv.DictReader(io.StringIO(raw.decode("utf-8-sig")))
    result = []
    for row in reader:
        time_str = row.get("Time", "").strip()
        time_local = time_str[:19].replace("T", " ") if time_str else ""
        result.append((time_local, row.get("Level", "").strip(), row.get("Message", "").strip()))
    return result


# 用例编号：TestCase_AcuHMI_009_07_case02
# 用例标题：选择间隔时间为1天，日志等级为危急，搜索事件日志并导出，
#           日志显示事件戳、等级及消息显示准确，并重置筛选条件
# 预置条件：1、管理权限登录AcuHMI网页
# 测试步骤：
#   1. Maintenance->Event Log设置间隔时间为1天，日志等级为Critical
#   2. 点击Search，校验筛选结果（逐页校验Level和Timestamp）
#   3. 点击Reset，校验筛选条件清空
#   4. 无筛选Search，收集全量页面数据，导出日志Export Logs，比对内容是否一致
# 预期结果：
#   1. 设置成功
#   2. 筛选结果 Level 全部为 Critical，Timestamp 在日期范围内
#   3. 筛选条件被重置
#   4. 导出文件行数与无筛选页面显示总行数一致，内容匹配
def test_TestCase_AcuHMI_009_07_case02(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_maintenance(page, "Event Log")

    today = date.today()
    yesterday = today - timedelta(days=1)
    start_date = yesterday.strftime("%Y-%m-%d")
    end_date = today.strftime("%Y-%m-%d")

    # Step 1: 设置 Interval=1天，Level=Critical
    page.locator(".el-form-item").filter(has_text="Interval").locator("input").first.click()
    page.wait_for_timeout(500)
    page.get_by_role("textbox", name="Start Date").fill(start_date)
    page.keyboard.press("Tab")
    page.wait_for_timeout(200)
    page.get_by_role("textbox", name="End Date").fill(end_date)
    page.keyboard.press("Tab")
    page.wait_for_timeout(200)
    page.locator(".el-picker-panel__link-btn").filter(has_text="OK").click()
    page.wait_for_timeout(300)

    page.locator(".el-form-item").filter(has_text="Level").locator(".el-select").click()
    page.wait_for_timeout(200)
    page.get_by_role("option", name="Critical").click()
    page.wait_for_timeout(200)

    # Step 2: Search，逐页校验筛选结果
    page.get_by_role("button", name="Search").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    assert page.locator("table, .el-table").count() > 0, "搜索后应显示日志列表"

    start_dt = datetime.strptime(start_date + " 00:00:00", "%Y-%m-%d %H:%M:%S")
    end_dt   = datetime.strptime(end_date   + " 23:59:59", "%Y-%m-%d %H:%M:%S")
    filtered_total, _ = _collect_and_verify_all_pages(
        page, start_dt, end_dt, "Critical", start_date, end_date
    )
    if filtered_total == 0:
        warnings.warn(
            f"筛选结果为0条Critical日志（{start_date} ~ {end_date}），"
            "请确认该时间段内是否存在Critical级别日志。",
            UserWarning, stacklevel=2,
        )

    # Step 3: Reset，校验筛选条件清空
    page.get_by_role("button", name="Reset").click()
    page.wait_for_timeout(500)
    interval_val = page.locator(".el-form-item").filter(has_text="Interval").locator("input").first.input_value()
    level_val    = page.locator(".el-form-item").filter(has_text="Level").locator(".el-select input").input_value()
    assert interval_val == "" or level_val == "", "Reset后筛选条件应被清空"

    # Step 4: 无筛选 Search，收集全量页面数据后立即导出，两者时间点一致
    page.get_by_role("button", name="Search").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    unfiltered_total, unfiltered_rows = _collect_all_pages_unfiltered(page)

    with page.expect_download(timeout=15000) as dl_info:
        page.get_by_role("button", name="Export Logs").click()
    download = dl_info.value
    assert download.suggested_filename != "", "导出日志应触发文件下载，文件名不为空"

    export_rows = _read_all_export_rows(download)

    assert len(export_rows) == unfiltered_total, \
        f"导出文件行数({len(export_rows)})与无筛选页面总行数({unfiltered_total})不一致"

    export_set = {(r[0], r[2]) for r in export_rows}
    missing = [(r[0], r[2]) for r in unfiltered_rows if (r[0], r[2]) not in export_set]
    assert len(missing) == 0, \
        f"以下{len(missing)}条无筛选页面记录在导出文件中未找到：\n" + \
        "\n".join(f"  {ts} | {msg}" for ts, msg in missing[:5])
