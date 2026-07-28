"""
SNMP 页面配置脚本
功能：
  1. 配置 v2c 参数（端口、Community、Trap、设备勾选）
  2. 配置 v3 参数（安全名称、认证、加密）
  3. 设备勾选状态控制（单选、全选、取消选择）
运行: python configure_snmp.py
"""
import logging
from playwright.sync_api import sync_playwright, Page

log = logging.getLogger(__name__)

import sys
from pathlib import Path

# 迁移到 RPP：连接配置改用 RPP 项目自身的 settings（projects/RPP/settings.py），
# 不再依赖 projects.AcuHMI_1_7。将 projects/RPP/ 加入 sys.path 后按裸模块名导入。
_RPP_ROOT = Path(__file__).resolve().parents[2]  # projects/RPP/
if str(_RPP_ROOT) not in sys.path:
    sys.path.insert(0, str(_RPP_ROOT))
from settings import BASE_URL, DEFAULT_USERNAME as USERNAME, DEFAULT_PASSWORD as PASSWORD

SNMP_CONFIG_V2C = {
    "version":       "v2c",
    "enable":        True,
    "port":          "161",
    "ro_community":  "123456789012",
    "trap_enable":   True,
    "trap_target_1": "192.168.2.9",
    "buffer_size":   "30",
    "hold_time":     "0",
}

SNMP_CONFIG_V3 = {
    "version":        "v3",
    "enable":         True,
    "port":           "161",
    "security_name":  "testuser",
    "auth_protocol":  "MD5",
    "auth_password":  "authpass123",
    "priv_protocol":  "DES",
    "priv_password":  "privpass123",
    "trap_enable":    False,
    "buffer_size":    "30",
    "hold_time":      "0",
}


def _goto_snmp_menu(page: Page) -> None:
    """RPP 网关菜单导航到 SNMP 配置页：Settings → Protocols → SNMP。

    RPP 的 SNMP 路由为 #/protocols/snmp，但存在路由守卫，登录后直接 goto 会被
    重定向回默认落地页，必须走菜单点击。菜单层级与 HMI 不同：
      - Settings 是顶部 <span>（非 button/a）
      - Protocols 是自定义 .left-nav-item（非 Element Plus .el-menu-item）
      - SNMP 是标准 .el-menu-item（横向 tab）
    """
    page.locator("text=Settings").first.click(timeout=5000)
    page.wait_for_timeout(800)
    page.locator(".left-nav-item").filter(has_text="Protocols").first.click(timeout=5000)
    page.wait_for_timeout(800)
    page.locator(".el-menu-item").filter(has_text="SNMP").first.click(timeout=5000)
    page.wait_for_timeout(2000)


def login_and_goto_snmp(page: Page) -> None:
    """登录并通过菜单导航到 SNMP 配置页面（RPP 网关）。"""
    log.info("[Browser] 导航到 %s", BASE_URL)
    page.goto(BASE_URL, timeout=15000)
    page.wait_for_selector('input[placeholder="Enter User Name"]', timeout=10000)
    page.fill('input[placeholder="Enter User Name"]', USERNAME)
    page.fill('input[placeholder="Enter Password"]', PASSWORD)
    page.keyboard.press("Enter")
    page.wait_for_timeout(3000)

    cancel_btn = page.query_selector('button:has-text("Cancel")')
    if cancel_btn:
        cancel_btn.click()
        page.wait_for_timeout(1000)
        log.info("[Browser] 关闭默认密码提示弹窗")

    _goto_snmp_menu(page)
    log.info("[Browser] 已进入 SNMP 配置页面")


def get_page_devices(page: Page) -> list:
    """从 SNMP 页面设备表格读取所有设备名称列表。"""
    rows = page.query_selector_all(".el-table__row")
    names = []
    for row in rows:
        cells = row.query_selector_all("td")
        if cells:
            name = cells[0].inner_text().strip()
            if name:
                names.append(name)
    return names


def set_device_selection(page: Page, selected_names, fallback_devices=None) -> None:
    """
    设置设备勾选状态：只勾选 selected_names 中的设备，取消其余所有设备。

    参数:
        selected_names:   需要勾选的设备名称列表。
                          [] 表示取消所有设备。
                          None 表示全选所有设备。
        fallback_devices: 当 selected_names 中没有任何设备匹配页面时使用的备选列表。
                          例如：fallback_devices=["AcuvimIIW"]

    注意：Element Plus checkbox 的真实选中状态通过 .el-checkbox__input 的
    CSS class "is-checked" 反映，不能依赖 input 元素的 is_checked() 方法。
    """
    rows = page.query_selector_all(".el-table__row")
    log.info("[Selection] 页面设备行数: %d  目标勾选: %s", len(rows), selected_names)

    any_checked = False
    for row in rows:
        # 用 textContent 而非 innerText：不受 CSS visibility/display 影响，
        # 即使下拉框覆盖或元素短暂不可见也能正确读取文字。
        row_text = row.evaluate("el => el.textContent") or ""
        checkbox_el = row.query_selector(".el-checkbox")
        if checkbox_el is None:
            continue

        # Element Plus: 通过 .el-checkbox__input 的 class 判断是否选中
        input_wrapper = row.query_selector(".el-checkbox__input")
        if input_wrapper is None:
            continue
        class_attr = input_wrapper.get_attribute("class") or ""
        is_checked = "is-checked" in class_attr

        should_check = (
            selected_names is None
            or any(name.lower() in row_text.lower() for name in selected_names)
        )

        if should_check:
            any_checked = True

        device_name = row_text.strip()[:40]
        if should_check != is_checked:
            action = "勾选" if should_check else "取消"
            log.info("[Selection] %s: %s", action, device_name)
            print(f"  [Checkbox] {action}: {device_name}")
            checkbox_el.click()
            page.wait_for_timeout(400)
        else:
            state = "已选中" if is_checked else "已取消"
            log.debug("[Selection] 无需变动 %s (%s)", device_name, state)

    # 若没有任何设备被匹配，且提供了 fallback，则用 fallback 重新选择
    if not any_checked and fallback_devices and selected_names is not None:
        log.warning("[Selection] 未找到设备 %s，使用 fallback: %s", selected_names, fallback_devices)
        print(f"  [Checkbox] 未找到 {selected_names}，切换到 fallback: {fallback_devices}")
        set_device_selection(page, fallback_devices)


def save_and_wait(page: Page, wait_ms: int = 5000) -> bool:
    """点击保存，等待成功弹窗并关闭，返回 True 表示无表单错误。"""
    page.click('button:has-text("Save")')
    page.wait_for_timeout(wait_ms)

    # 处理保存成功后的确认弹窗（ElMessageBox / ElDialog）
    for selector in [
        'button:has-text("OK")',
        'button:has-text("确定")',
        'button:has-text("Confirm")',
    ]:
        btn = page.query_selector(selector)
        if btn:
            log.info("[Save] 检测到保存成功弹窗，点击关闭")
            btn.click()
            page.wait_for_timeout(1000)
            break

    error_els = page.query_selector_all(".el-form-item.is-error")
    ok = len(error_els) == 0
    log.info("[Save] 保存结果: %s  (表单错误数=%d)", "成功" if ok else "失败", len(error_els))
    return ok


def configure_snmp_v2c(config: dict = None, selected_devices=None,
                        headless: bool = False) -> None:
    """
    完整配置 SNMP v2c 参数并保存。

    参数:
        config:           v2c 配置字典，None 使用 SNMP_CONFIG_V2C 默认值
        selected_devices: 勾选的设备名称列表；None=全选；[]=全部取消
    """
    if config is None:
        config = SNMP_CONFIG_V2C

    log.info("[configure_snmp_v2c] 开始配置  selected_devices=%s", selected_devices)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(ignore_https_errors=True)
        page = context.new_page()
        login_and_goto_snmp(page)

        _click_radio(page, "Enable" if config["enable"] else "Disable", 0)

        # 等待表单展开（SNMP 从 Disabled→Enabled 时字段异步出现）
        if config["enable"]:
            try:
                page.wait_for_selector('input[placeholder="Enter Port"]', timeout=5000)
            except Exception:
                page.wait_for_timeout(1000)

        if selected_devices is not None:
            set_device_selection(page, selected_devices)
        else:
            set_device_selection(page, None)

        port_f = page.locator('input[placeholder="Enter Port"]')
        port_f.click()
        port_f.fill(config["port"])

        comm_f = page.locator('input[placeholder="Enter RO Community"]')
        comm_f.click()
        comm_f.fill(config["ro_community"])

        _click_radio(page, "Enable" if config["trap_enable"] else "Disable", 1)
        if config.get("trap_enable") and config.get("trap_target_1"):
            t1 = page.locator('input[placeholder="Enter Trap Target 1"]')
            t1.click()
            t1.fill(config["trap_target_1"])

        buf_f = page.locator('input[placeholder="Enter Report Buffer Size"]')
        buf_f.click()
        buf_f.fill(config["buffer_size"])

        hold_f = page.locator('input[placeholder="Enter Report Hold Time"]')
        hold_f.click()
        hold_f.fill(config["hold_time"])

        ok = save_and_wait(page)
        print(f"[configure_snmp_v2c] {'OK' if ok else 'WARN'}: "
              f"port={config['port']} community={config['ro_community']} "
              f"selected={selected_devices}")
        browser.close()


def configure_snmp_v3(config: dict = None, selected_devices=None,
                       headless: bool = False) -> None:
    """
    配置 SNMP v3 参数（USM 认证加密）。

    v3 UI 字段（以下 selector 需与实际页面核对）：
      - SNMP Version 下拉 → SNMPv3
      - input[placeholder="Enter Security Name"]
      - Auth Protocol 下拉（MD5/SHA）
      - input[placeholder="Enter Auth Password"]
      - Priv Protocol 下拉（DES/AES）
      - input[placeholder="Enter Priv Password"]
    """
    if config is None:
        config = SNMP_CONFIG_V3

    log.info("[configure_snmp_v3] 开始 v3 配置  user=%s auth=%s priv=%s",
             config["security_name"], config["auth_protocol"], config["priv_protocol"])

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(ignore_https_errors=True)
        page = context.new_page()
        login_and_goto_snmp(page)

        _click_radio(page, "Enable" if config["enable"] else "Disable", 0)

        _select_version(page, "v3")
        page.wait_for_timeout(800)

        if selected_devices is not None:
            set_device_selection(page, selected_devices)
        else:
            set_device_selection(page, None)

        sn = page.locator('input[placeholder="Enter Security Name"]')
        if sn.count() > 0:
            sn.click()
            sn.fill(config["security_name"])
            log.info("[v3] Security Name: %s", config["security_name"])

        _select_dropdown_by_text(page, config["auth_protocol"])

        ap = page.locator('input[placeholder="Enter Auth Password"]')
        if ap.count() > 0:
            ap.click()
            ap.fill(config["auth_password"])

        _select_dropdown_by_text(page, config["priv_protocol"])

        pp = page.locator('input[placeholder="Enter Priv Password"]')
        if pp.count() > 0:
            pp.click()
            pp.fill(config["priv_password"])

        port_f = page.locator('input[placeholder="Enter Port"]')
        port_f.click()
        port_f.fill(config["port"])

        _click_radio(page, "Enable" if config["trap_enable"] else "Disable", 1)

        buf_f = page.locator('input[placeholder="Enter Report Buffer Size"]')
        buf_f.click()
        buf_f.fill(config["buffer_size"])

        hold_f = page.locator('input[placeholder="Enter Report Hold Time"]')
        hold_f.click()
        hold_f.fill(config["hold_time"])

        ok = save_and_wait(page)
        print(f"[configure_snmp_v3] {'OK' if ok else 'WARN'}: user={config['security_name']}")
        browser.close()


def get_v3_config_from_page(headless: bool = False) -> dict:
    """读取页面当前 v3 配置。"""
    result = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(ignore_https_errors=True)
        page = context.new_page()
        login_and_goto_snmp(page)

        for placeholder, key in [
            ("Enter Security Name", "security_name"),
            ("Enter Auth Password", "auth_password"),
            ("Enter Priv Password", "priv_password"),
        ]:
            el = page.locator(f'input[placeholder="{placeholder}"]')
            result[key] = el.input_value() if el.count() > 0 else ""

        browser.close()
    return result


# ─── 内部辅助 ─────────────────────────────────────────────────────────────────

def _click_radio(page: Page, label: str, nth: int) -> None:
    page.locator(".el-radio").filter(has_text=label).nth(nth).click()
    page.wait_for_timeout(300)


def _select_version(page: Page, version: str) -> None:
    """在 SNMP Version 下拉中选择版本。"""
    version_label = "SNMP v2c" if version == "v2c" else "SNMP v3"
    try:
        triggers = page.locator(".el-select")
        if triggers.count() > 0:
            triggers.first.click()
            page.wait_for_timeout(400)
            option = page.locator(f".el-select-dropdown__item:has-text('{version_label}')")
            if option.count() > 0:
                option.first.click()
                page.wait_for_timeout(400)
                log.info("[Version] 已切换到 %s", version_label)
    except Exception as e:
        log.warning("[Version] 版本切换异常: %s", e)


def _select_dropdown_by_text(page: Page, value: str) -> None:
    """在任意 el-select 下拉中选择包含指定文本的选项。"""
    try:
        option = page.locator(f".el-select-dropdown__item:has-text('{value}')")
        if option.count() > 0:
            option.first.click()
            page.wait_for_timeout(300)
    except Exception:
        pass


# ─── 页面级函数（接收已有 page，供测试 session 复用）────────────────────────────

def goto_snmp(page: Page) -> None:
    """从已登录状态重新导航到 SNMP 页面（不重新登录）。"""
    _goto_snmp_menu(page)
    log.info("[goto_snmp] 已导航到 SNMP 页面")


def apply_snmp_v2c(page: Page, config: dict, selected_devices=None, fallback_devices=None) -> bool:
    """
    在已打开的 SNMP page 上应用 v2c 配置并保存。

    config 键：enable, port, ro_community, trap_enable, trap_target_1,
               buffer_size, hold_time（均可选，缺省时不修改对应字段）。
    返回 True 表示保存成功（无表单校验错误）。

    注意：SNMP 关闭时页面仅显示 Enable/Trap 两个 radio，其余字段全部隐藏。
    enable=False 时直接保存，不尝试填写隐藏字段。
    """
    enable = config.get("enable", True)
    _click_radio(page, "Enable" if enable else "Disable", 0)

    if not enable:
        # SNMP 关闭时表单折叠，直接保存即可
        return save_and_wait(page)

    # 等待表单展开（SNMP 从 Disabled→Enabled 时字段异步出现）
    try:
        page.wait_for_selector('input[placeholder="Enter Port"]', timeout=5000)
    except Exception:
        page.wait_for_timeout(1000)

    _select_version(page, "v2c")
    page.wait_for_timeout(600)

    set_device_selection(page, selected_devices, fallback_devices=fallback_devices)

    port_f = page.locator('input[placeholder="Enter Port"]')
    port_f.click()
    port_f.fill(str(config.get("port", "161")))

    comm_f = page.locator('input[placeholder="Enter RO Community"]')
    comm_f.click()
    comm_f.fill(config.get("ro_community", ""))

    trap_en = config.get("trap_enable", False)
    _click_radio(page, "Enable" if trap_en else "Disable", 1)
    page.wait_for_timeout(400)

    if trap_en and config.get("trap_target_1"):
        t1 = page.locator('input[placeholder="Enter Trap Target 1"]')
        if t1.count() > 0:
            t1.click()
            t1.fill(config["trap_target_1"])

    if "buffer_size" in config:
        buf_f = page.locator('input[placeholder="Enter Report Buffer Size"]')
        if buf_f.count() > 0:
            buf_f.click()
            buf_f.fill(str(config["buffer_size"]))

    if "hold_time" in config:
        hold_f = page.locator('input[placeholder="Enter Report Hold Time"]')
        if hold_f.count() > 0:
            hold_f.click()
            hold_f.fill(str(config["hold_time"]))

    return save_and_wait(page)


def apply_snmp_v3(page: Page, config: dict, selected_devices=None, fallback_devices=None) -> bool:
    """
    在已打开的 SNMP page 上应用 v3 配置并保存。

    config 键：enable, port, username, password, auth_protocol,
               priv_protocol, priv_password（priv_protocol="NONE PRIV" 时不填 priv_password）。
    返回 True 表示保存成功。

    注意：SNMP 关闭时页面仅显示 Enable/Trap radio，其余字段隐藏。
    enable=False 时直接保存。
    """
    enable = config.get("enable", True)
    _click_radio(page, "Enable" if enable else "Disable", 0)

    if not enable:
        return save_and_wait(page)

    # 等待表单展开
    try:
        page.wait_for_selector('input[placeholder="Enter Port"]', timeout=5000)
    except Exception:
        page.wait_for_timeout(1000)

    _select_version(page, "v3")
    page.wait_for_timeout(800)

    set_device_selection(page, selected_devices, fallback_devices=fallback_devices)

    port_f = page.locator('input[placeholder="Enter Port"]')
    port_f.click()
    port_f.fill(str(config.get("port", "161")))

    # Security Name / Username
    for ph in ["Enter Security Name", "Enter Username", "Enter User Name"]:
        f = page.locator(f'input[placeholder="{ph}"]')
        if f.count() > 0:
            f.click()
            f.fill(config.get("username", config.get("security_name", "")))
            break

    # Auth Protocol dropdown（el-select 第 1 个，第 0 个是 Version）
    _open_dropdown_and_select(page, 1, config.get("auth_protocol", "MD5"))

    # Auth/User Password
    for ph in ["Enter User Password", "Enter Auth Password", "Enter Password"]:
        f = page.locator(f'input[placeholder="{ph}"]')
        if f.count() > 0:
            f.click()
            f.fill(config.get("password", config.get("auth_password", "")))
            break

    # Privacy Protocol dropdown（第 2 个）
    priv = config.get("priv_protocol", "NONE PRIV")
    _open_dropdown_and_select(page, 2, priv)

    # Privacy Password（NONE PRIV 无此字段）
    if priv.upper().replace("_", " ") not in ("NONE PRIV",):
        for ph in ["Enter Privacy Password", "Enter Priv Password"]:
            f = page.locator(f'input[placeholder="{ph}"]')
            if f.count() > 0:
                f.click()
                f.fill(config.get("priv_password", ""))
                break

    # Trap: 默认关闭
    _click_radio(page, "Disable", 1)

    return save_and_wait(page)


def download_mib_file(page: Page) -> str:
    """点击 Download MIB File 按钮，等待下载完成，返回文件名；失败返回空字符串。"""
    btn = page.locator(
        'button:has-text("Download MIB File"), a:has-text("Download MIB File"),'
        'button:has-text("MIB File Download"), a:has-text("MIB File Download")'
    )
    if btn.count() == 0:
        log.warning("[MIB] 未找到 MIB 下载按钮")
        return ""
    try:
        with page.expect_download(timeout=30000) as dl_info:
            btn.first.click()
        dl = dl_info.value
        filename = dl.suggested_filename
        log.info("[MIB] 下载完成: %s", filename)
        print(f"[MIB] 下载完成: {filename}")
        return filename
    except Exception as e:
        log.warning("[MIB] 下载失败: %s", e)
        print(f"[MIB] 下载失败: {e}")
        return ""


def _open_dropdown_and_select(page: Page, nth: int, value: str) -> None:
    """打开第 nth 个 el-select 下拉并选择包含 value 的选项。"""
    selects = page.locator(".el-select")
    cnt = selects.count()
    if cnt <= nth:
        log.warning("[Dropdown] el-select 数量不足: count=%d, nth=%d", cnt, nth)
        return
    try:
        selects.nth(nth).click()
        page.wait_for_timeout(400)
        option = page.locator(f".el-select-dropdown__item:has-text('{value}')")
        if option.count() > 0:
            option.first.click()
            page.wait_for_timeout(300)
        else:
            log.warning("[Dropdown] 未找到选项: %s（nth=%d）", value, nth)
    except Exception as e:
        log.warning("[Dropdown] 操作失败 nth=%d value=%s: %s", nth, value, e)


if __name__ == "__main__":
    configure_snmp_v2c(headless=False)
