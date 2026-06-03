import warnings
import csv
import io
import os
import tarfile
import tempfile
from playwright.sync_api import expect
from config.settings import BASE_URL
from pages.login_page import LoginPage


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


def _read_export_rows(download):
    """从导出 tar 文件中读取 CSV，返回所有行 [(time_local, level, message)]。"""
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


# 用例编号：TestCase_AcuHMI_009_07_case04
# 用例标题：清除事件日志，日志全部被清理
# 预置条件：
#   1. 管理权限登录AcuHMI网页
#   2. 已有日志
# 测试步骤：
#   1. 点击Search加载现有日志，确认有数据
#   2. 点击Clear Logs，在弹窗中点击Yes确认
#   3. 重新Search，校验页面日志列表已清空
#   4. 导出日志Export Logs，校验导出文件内容已清空
# 预期结果：
#   1. 清除前存在日志数据
#   2. 弹窗点击Yes后确认操作
#   3. 日志列表为空（或仅剩系统自动写入的清除操作记录）
#   4. 导出文件中包含清除日志操作的记录（消息含关键字 "clear"）
def test_TestCase_AcuHMI_009_07_case04(login_page: LoginPage):
    login_page.open()
    login_page.login()
    page = login_page.page

    _nav_to_maintenance(page, "Event Log")

    # Step 1: Search 加载现有日志，确认清除前有数据
    page.get_by_role("button", name="Search").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    rows_before = page.locator(".el-table__body tr.el-table__row").all()
    if len(rows_before) == 0:
        warnings.warn(
            "清除前日志列表已为空，无法验证Clear Logs功能是否真正清除了数据。",
            UserWarning,
            stacklevel=2,
        )

    # Step 2: 点击 Clear Logs，弹窗中点击 Yes
    page.get_by_role("button", name="Clear Logs").click()
    page.wait_for_timeout(500)

    confirmed = False
    for btn_name in ["Yes", "Yes, continue", "Confirm", "OK"]:
        try:
            page.get_by_role("button", name=btn_name).click(timeout=2000)
            page.wait_for_timeout(500)
            confirmed = True
            break
        except Exception:
            pass
    assert confirmed, "未能点击确认弹窗按钮（Yes/Yes, continue/Confirm/OK），请确认弹窗按钮名称"

    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)

    # Step 3: 重新 Search 刷新列表，校验日志已清空
    page.get_by_role("button", name="Search").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    rows_after = page.locator(".el-table__body tr.el-table__row").all()
    remaining = len(rows_after)

    if remaining == 0:
        pass  # 理想结果：完全清空
    elif remaining == 1:
        # 系统可能在清除后自动写入一条操作记录，视为正常
        warnings.warn(
            f"Clear Logs后页面仍有1条记录，可能为系统自动写入的清除操作日志，视为正常。",
            UserWarning,
            stacklevel=2,
        )
    else:
        assert False, \
            f"点击Clear Logs并确认后日志列表应为空，但仍有 {remaining} 条记录"

    # Step 4: 导出日志，校验导出文件中存在清除操作日志记录
    with page.expect_download(timeout=15000) as dl_info:
        page.get_by_role("button", name="Export Logs").click()
    download = dl_info.value
    assert download.suggested_filename != "", "导出日志应触发文件下载，文件名不为空"

    export_rows = _read_export_rows(download)

    # 在所有导出行中查找包含 "clear" 关键字（大小写不敏感）的消息
    clear_records = [
        (ts, lvl, msg) for ts, lvl, msg in export_rows
        if "clear" in msg.lower()
    ]
    assert len(clear_records) > 0, (
        f"导出文件共 {len(export_rows)} 条记录，但未找到任何包含清除日志操作的记录（关键字'clear'）。\n"
        + (f"现有记录消息：\n" + "\n".join(f"  {msg}" for _, _, msg in export_rows[:5])
           if export_rows else "导出文件为空。")
    )
