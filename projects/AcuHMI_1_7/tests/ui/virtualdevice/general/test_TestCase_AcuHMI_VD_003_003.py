# 用例编号: AcuHMI_VD_003_003
# 用例标题: 新增两个 Virtual Device → 删除甲 → 验证乙仍存在（含测后清理）
# 预置条件: 已登录系统；HMI 上至少有一台已添加的物理设备（本用例使用 pxm350）
# 测试步骤:
#   1. 新增 Virtual Device 甲（名 test_VirtualDevice_010_<ts>），
#      Calculated Meter Formula 通过 Select Device Parameter 弹窗选取 pxm350 的两个参数相加，
#      Parameter Name=test_Paramter，Post Label=test_postlabel，Unit=测试单位，点 Save
#   2. 断言甲已出现在 Virtual Devices 列表（遍历所有分页）
#   3. 新增 Virtual Device 乙（名 test_VirtualDevice_020_<ts>），同上字段，点 Save
#   4. 断言乙已出现在 Virtual Devices 列表（遍历所有分页）
#   5. 在列表中删除甲（Action 列最后一个按钮 → Yes，支持跨页查找）
#   6. 遍历所有分页，断言甲已消失，乙仍存在
# 预期结果:
#   - 甲保存成功，列表（所有分页）可见
#   - 乙保存成功，列表（所有分页）可见
#   - 甲删除后从所有分页消失；乙仍存在
# 测试后清理:
#   - finally 块兜底删除甲和乙，_delete_virtual_device 对不存在的行是 no-op

import time

from playwright.sync_api import Page, expect

from projects.AcuHMI_1_7.pages.login_page import LoginPage

# 用于 Select Device Parameter 弹窗的目标物理设备（HMI 上已添加）
_FORMULA_DEVICE = "pxm350"


def _nav_to_virtual_devices(page: Page) -> None:
    """确保当前停留在 Virtual Devices 列表页（不是 addVirtualMeter 子路径）。"""
    on_list = "/#/virtualMeter" in page.url and "addVirtualMeter" not in page.url
    if not on_list:
        if not any(s in page.url for s in [
            "/#/dashboard", "/#/physicalDevice", "/#/virtualMeter",
            "/#/webDevice", "/#/alarm", "/#/dataLog",
        ]):
            page.locator("header span").filter(has_text="Devices").first.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(500)
        page.locator(".left-nav-item").filter(has_text="Virtual Devices").click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _count_vd_across_pages(page: Page, name: str) -> int:
    """从第一页开始遍历所有分页，统计名称匹配的行数，返回总计数。

    列表默认每页 10 行，新建的 VD 会追加到末页，故需跨页查找。
    查完后回到第一页，保持页面状态一致。
    """
    _nav_to_virtual_devices(page)
    total = 0
    while True:
        row = page.locator("tbody").get_by_role("row").filter(has_text=name)
        total += row.count()
        next_btn = page.locator(".el-pagination .btn-next")
        if next_btn.count() == 0:
            break
        if next_btn.get_attribute("aria-disabled") == "true":
            break
        next_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)
    # 回到第一页（若当前不在第一页）
    page1_btn = page.locator(".el-pagination .el-pager li").first
    if page1_btn.count() > 0 and "is-active" not in (page1_btn.get_attribute("class") or ""):
        page1_btn.click()
        page.wait_for_timeout(500)
    return total


def _delete_virtual_device(page: Page, name: str) -> None:
    """在 Virtual Devices 列表（含分页）中删除指定名称的设备；若不存在则静默返回（no-op）。"""
    _nav_to_virtual_devices(page)
    while True:
        row = page.locator("tbody").get_by_role("row").filter(has_text=name)
        if row.count() > 0:
            row.locator(".el-button").last.click(force=True)
            page.wait_for_timeout(500)
            try:
                page.get_by_role("button", name="Yes").click(timeout=2000)
                page.wait_for_timeout(500)
            except Exception:
                try:
                    page.get_by_role("button", name="Confirm").click(timeout=2000)
                    page.wait_for_timeout(500)
                except Exception:
                    pass
            return
        next_btn = page.locator(".el-pagination .btn-next")
        if next_btn.count() == 0:
            break
        if next_btn.get_attribute("aria-disabled") == "true":
            break
        next_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(500)


def _select_device_parameter(page: Page, device_name: str, param_index: int) -> str:
    """点击 'Select Device Parameter' 按钮，在弹窗中选取指定设备的第 param_index 个参数，
    点 Confirm 将其插入公式输入框末尾。

    返回所选参数文本，用于断言或日志。
    弹窗是 El Plus v2 Select（aria-controls 为 None），通过全局
    .el-select-dropdown__item 定位可见选项。
    注意：两次调用之间需由调用方在公式框末尾追加 '+' 等运算符。
    """
    page.locator("button.el-button--Copy").first.click()
    page.wait_for_timeout(1500)

    dialog = page.locator('[aria-label="Select Device Parameter"]')
    expect(dialog).to_be_visible()

    # El Plus v2 Select：两个 .el-select__input，第一个是 Device，第二个是 Parameter
    select_inputs = dialog.locator(".el-select__input")
    device_input = select_inputs.first
    param_input = select_inputs.last

    # 打开 Device 下拉，选指定设备
    device_input.click()
    page.wait_for_timeout(800)
    page.locator(".el-select-dropdown__item").filter(has_text=device_name).first.click()
    page.wait_for_timeout(800)

    # 打开 Parameter 下拉，取可见选项列表
    param_input.click()
    page.wait_for_timeout(800)

    visible_items = [
        item for item in page.locator(".el-select-dropdown__item").all()
        if item.is_visible()
    ]
    assert len(visible_items) > param_index, (
        f"设备 '{device_name}' 的参数列表只有 {len(visible_items)} 项，"
        f"无法取第 {param_index} 个（0-based）"
    )
    selected_text = visible_items[param_index].inner_text().strip()
    visible_items[param_index].click()
    page.wait_for_timeout(800)

    # 点 Confirm 插入公式
    dialog.get_by_role("button", name="Confirm").click()
    page.wait_for_timeout(800)

    return selected_text


def _add_virtual_device(page: Page, vd_name: str) -> None:
    """新增一个 Virtual Device，Calculated Meter Formula 通过 Select Device Parameter
    弹窗选取 _FORMULA_DEVICE 的两个参数相加（param0+"+"param1），保存后等待跳回列表页。

    调用方负责断言保存结果。
    """
    _nav_to_virtual_devices(page)
    page.get_by_role("button", name="Add Virtual Device").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)

    page.get_by_label("Virtual Device Name", exact=True).fill(vd_name)

    # 表单默认已有一行参数，直接填写（无需点 Add Parameter）
    page.get_by_placeholder("---Enter Parameter Name---").first.fill("test_Paramter")
    page.get_by_placeholder("---Enter Post Label---").first.fill("test_postlabel")

    # 公式：通过 Select Device Parameter 弹窗选第 1 个参数
    _select_device_parameter(page, _FORMULA_DEVICE, 0)

    # 在公式框末尾手动追加 '+'，再通过弹窗选第 2 个参数
    formula_input = page.get_by_placeholder("---Enter Calculated Meter Formula---").first
    formula_input.click()
    formula_input.press("End")
    formula_input.type("+")
    page.wait_for_timeout(300)

    _select_device_parameter(page, _FORMULA_DEVICE, 1)

    page.get_by_placeholder("---Enter Unit---").first.fill("测试单位")

    page.get_by_role("button", name="Save").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1000)


def test_TestCase_AcuHMI_VD_003_003(login_page: LoginPage) -> None:
    login_page.open()
    login_page.login()
    page = login_page.page

    ts = str(int(time.time()))[-6:]
    vd_name_a = f"test_VirtualDevice_010_{ts}"
    vd_name_b = f"test_VirtualDevice_020_{ts}"

    try:
        # ── 步骤 1：新增 Virtual Device 甲 ───────────────────────────────────
        _add_virtual_device(page, vd_name_a)

        # 步骤 2：断言甲出现在列表（遍历所有分页）
        count_a = _count_vd_across_pages(page, vd_name_a)
        assert count_a > 0, (
            f"Virtual Device 甲 '{vd_name_a}' 保存后在所有分页均未找到"
        )

        # ── 步骤 3：新增 Virtual Device 乙 ───────────────────────────────────
        _add_virtual_device(page, vd_name_b)

        # 步骤 4：断言乙出现在列表（遍历所有分页）
        count_b = _count_vd_across_pages(page, vd_name_b)
        assert count_b > 0, (
            f"Virtual Device 乙 '{vd_name_b}' 保存后在所有分页均未找到"
        )

        # ── 步骤 5：删除甲 ────────────────────────────────────────────────────
        _delete_virtual_device(page, vd_name_a)

        # ── 步骤 6：核心双向断言（遍历所有分页）────────────────────────────────
        count_a_after = _count_vd_across_pages(page, vd_name_a)
        count_b_after = _count_vd_across_pages(page, vd_name_b)

        # 先断言乙 presence（防空转：先确认列表可查到乙，再断言甲已删除）
        assert count_b_after > 0, (
            f"Virtual Device 乙 '{vd_name_b}' 在删除甲之后从所有分页消失，"
            f"预期乙应仍存在"
        )
        assert count_a_after == 0, (
            f"Virtual Device 甲 '{vd_name_a}' 删除后仍出现在列表中，"
            f"预期已被删除"
        )

    finally:
        # 兜底清理：甲已在步骤 5 删除，但 _delete_virtual_device 对不存在行是 no-op，安全
        # 两个清理操作各自独立 try，避免甲清理异常导致乙未清理
        try:
            _delete_virtual_device(page, vd_name_a)
        except Exception:
            pass
        try:
            _delete_virtual_device(page, vd_name_b)
        except Exception:
            pass
