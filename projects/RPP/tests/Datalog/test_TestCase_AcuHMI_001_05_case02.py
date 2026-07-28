from pathlib import Path

from playwright.sync_api import Page

from projects.RPP.tests.Datalog import helpers_datalog as hd


# 用例编号：TestCase_AcuHMI_001_05_case02
# 用例标题：接入设备 log 栏，通过 "clear log" 可清除当前接入设备采集到的
#          所有 Datalog
# 预置条件：
#   1. AcuHMI 上电
#   2. 设备已接入并在线，且已采集数据超过 10mins
# 测试步骤：
#   2. 切换到 Datalog 页面，点击 Clear Logs 按钮并确认删除日志
#   3. 选中 clear 前有 datalog 数据的时间段，间隔 1mins，点击导出
#   4. 检查导出文件中是否有 datalog 数据
# 预期结果：
#   4. 导出文件中没有 datalog 数据
# 说明：破坏性用例（清空该设备全部历史 datalog），放本组最后执行；
#      清除后大间隔档位（6h/12h/1day…）需重新累计数据才能完整校验间隔。


def test_TestCase_AcuHMI_001_05_case02(app_page: Page, tmp_path: Path):
    page = app_page
    hd.goto_device_datalog(page)

    # 前置确认：清除前确实有数据
    start_date, end_date = hd.date_range_values(page)
    hd.select_interval(page, "1 minute")
    stamps_before, _ = hd.download_csv(page, tmp_path)
    assert stamps_before, "前置不满足：清除前设备应有 datalog 数据"

    # ── Step 2: Clear Logs 并确认 ──
    hd.clear_logs(page)
    assert page.locator(".el-message--error").count() == 0, \
        "清除日志后页面不应出现错误提示"

    # ── Step 3-4: 用 clear 前有数据的区间导出，文件应无数据 ──
    stamps_after, info = hd.download_csv(page, tmp_path)
    if stamps_after is None:
        # 产品直接拒绝下载（无数据提示）——同样满足"没有 datalog 数据"
        assert True
    else:
        assert len(stamps_after) == 0, (
            f"Clear Logs 后导出文件应无数据行，实际仍有 {len(stamps_after)} 行"
            f"（区间 {start_date}~{end_date}，文件 {info!r}）"
        )
