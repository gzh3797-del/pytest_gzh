# -*- coding: utf-8 -*-
"""
datetime_helpers.py — Date & Time（系统时间 / NTP / 时区）用例共用辅助

供 systemsettings/general/ 下所有「会修改系统时间」的用例
（005_01_case01 / case02 / case03_xx / case04）共用：

  1. 时间一致性判定统一采用「本机本地时间对比」——测试机与设备同处一个时区，
     设备本地时钟应 ≈ 测试机本地时间（datetime.now()），不再做 UTC 换算。
     这从根本上规避了时区缩写歧义：例如 CST 既是美国中部 UTC-6，又是中国标准 UTC+8，
     旧的缩写映射会把中国 CST 误判为 -6 小时，凭空差出约 14 小时。

  2. 执行后状态恢复（snapshot / restore）：用例在修改任何设置之前先 snapshot，
     用例结束后 restore——还原时区、NTP 开关、NTP Server，并用已知有效 NTP
     重新同步，确保设备时钟回到正确的当前时间，不污染后续用例。
     （case01 会把时区改成 New_York、case04 改成随机时区且不还原，
       不恢复时区会让后续用例的「本机对比」整体偏移十几个小时。）
"""
import re
from datetime import datetime
from typing import Optional

from playwright.sync_api import Page

# teardown 重新同步时使用的已知有效 NTP 服务器，保证时钟能被拉回正确当前时间
_KNOWN_GOOD_NTP = ["time.google.com", "time.nist.gov", "time.apple.com"]
_DATETIME_HASH = "#/systemSettings/dateTime"


# ── 基础读写 ──────────────────────────────────────────────────────────────────

def nav_to_datetime(page: Page) -> None:
    """跳转到 System Settings → Date & Time 并等待页面稳定。"""
    base = page.url.split("#")[0]
    page.goto(base + _DATETIME_HASH)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)


def read_device_clock(page: Page) -> tuple[datetime, str]:
    """读取 Device Clock 输入框，返回 (datetime, 原始字符串)。"""
    clock_input = page.get_by_placeholder("--Select Device Clock--").first
    assert clock_input.count() > 0, "未找到 Device Clock 输入框"
    clock_str = clock_input.input_value().strip()
    assert clock_str, "Device Clock 输入框值为空"
    for fmt in ("%Y/%m/%d %I:%M %p", "%Y/%m/%d %I:%M:%S %p"):
        try:
            return datetime.strptime(clock_str, fmt), clock_str
        except ValueError:
            continue
    raise AssertionError(f"Device Clock 格式无法解析：'{clock_str}'")


def assert_clock_near_now(page: Page, ref_local: datetime, label: str,
                          tol_seconds: int = 120) -> None:
    """读取 Device Clock，与测试机本地时间 ref_local 比较，断言误差 ≤ tol_seconds。

    采用本机本地时间对比（设备与测试机同时区），无需 UTC 换算，规避时区缩写歧义。
    """
    device_dt, clock_str = read_device_clock(page)
    diff = abs((device_dt - ref_local).total_seconds())
    assert diff <= tol_seconds, (
        f"{label} Sync 后 Device Clock 与本机时间不一致：\n"
        f"  Device Clock = {clock_str}\n"
        f"  本机参考时间 = {ref_local.strftime('%Y/%m/%d %I:%M:%S %p')}\n"
        f"  差值 = {diff:.0f} 秒（允许 ≤ {tol_seconds} 秒）"
    )


# ── 对话框 / 保存 ─────────────────────────────────────────────────────────────

def _dismiss_dialog(page: Page, wait_ms: int = 800) -> None:
    """关闭可能弹出的 el-message-box 确认框（优先点 primary 按钮，兜底回车）。"""
    try:
        page.wait_for_selector(".el-overlay.is-message-box", timeout=wait_ms)
    except Exception:
        return
    page.wait_for_timeout(200)
    overlay = page.locator(".el-overlay.is-message-box")
    primary = overlay.locator(".el-message-box__btns .el-button--primary")
    if primary.count() > 0:
        try:
            primary.first.click(timeout=3000)
            page.wait_for_timeout(400)
            return
        except Exception:
            pass
    page.keyboard.press("Enter")
    page.wait_for_timeout(400)


def _click_save(page: Page) -> None:
    """点击 Save 并等待加载完成，处理前后可能出现的确认框。"""
    _dismiss_dialog(page, wait_ms=400)
    save_btn = page.get_by_role("button", name="Save").first
    if save_btn.count() == 0:
        return
    save_btn.click(timeout=15000)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(800)
    _dismiss_dialog(page, wait_ms=1000)


# ── NTP / 时区 读写 ───────────────────────────────────────────────────────────

def is_ntp_enabled(page: Page) -> bool:
    """返回当前 NTP Enable 单选是否处于选中（Enable）状态。"""
    ntp_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    if ntp_item.count() == 0:
        return True
    enable_radio = ntp_item.locator(".el-radio").filter(has_text="Enable")
    return "is-checked" in (enable_radio.get_attribute("class") or "")


def set_ntp_enable(page: Page, enable: bool) -> None:
    """把 NTP Enable 切到目标状态（已是目标态则不动），并关闭可能的确认框。"""
    label = "Enable" if enable else "Disable"
    ntp_item = page.locator(".el-form-item").filter(has_text="NTP Enable").first
    if ntp_item.count() == 0:
        return
    radio = ntp_item.locator(".el-radio").filter(has_text=label)
    if "is-checked" not in (radio.get_attribute("class") or ""):
        ntp_item.locator(".el-radio__label").filter(has_text=label).click()
        page.wait_for_timeout(600)
    _dismiss_dialog(page, wait_ms=600)


def read_ntp_servers(page: Page) -> list:
    """读取 NTP Server 1/2/3 的当前值（输入框不存在则记为空串）。"""
    servers = []
    for i in range(1, 4):
        inp = page.get_by_placeholder(f"NTP Server {i}").first
        servers.append(inp.input_value().strip() if inp.count() > 0 else "")
    return servers


def fill_ntp_servers(page: Page, servers: list) -> None:
    """依次写入 NTP Server 1/2/3（仅对可见且未禁用的输入框）。"""
    for i, srv in enumerate(servers, 1):
        inp = page.get_by_placeholder(f"NTP Server {i}").first
        if inp.count() > 0 and inp.is_visible() and not inp.is_disabled():
            inp.fill(srv)
            page.wait_for_timeout(100)


def read_timezone_id(page: Page) -> Optional[str]:
    """从 Time Zone 字段文字中提取 IANA 时区名（如 'Asia/Shanghai'），取不到返回 None。"""
    tz_fi = page.locator(".el-form-item").filter(has_text="Time Zone").first
    if tz_fi.count() == 0:
        return None
    match = re.search(r"[A-Za-z]+/[A-Za-z0-9_+\-/]+", tz_fi.inner_text())
    return match.group(0) if match else None


def select_timezone(page: Page, search_key: str) -> bool:
    """打开 Time Zone 下拉，搜索并选中含 search_key 的项；成功返回 True。"""
    tz_fi = page.locator(".el-form-item").filter(has_text="Time Zone").first
    if tz_fi.count() == 0:
        return False
    tz_select = tz_fi.locator(".el-select").first
    if tz_select.count() == 0:
        return False
    for _ in range(3):
        try:
            tz_select.click(timeout=5000)
            break
        except Exception:
            _dismiss_dialog(page, wait_ms=600)
    page.wait_for_timeout(400)
    search_inp = page.locator(".el-select-dropdown input").first
    if search_inp.count() > 0 and search_inp.is_visible():
        search_inp.fill(search_key)
        page.wait_for_timeout(400)
    for opt in page.locator(".el-select-dropdown__item").all():
        try:
            if opt.is_visible() and search_key in opt.inner_text():
                opt.click()
                page.wait_for_timeout(200)
                return True
        except Exception:
            pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)
    return False


# ── 快照 / 恢复 ───────────────────────────────────────────────────────────────

def snapshot_datetime_settings(page: Page) -> dict:
    """记录执行前的 Date & Time 关键状态（时区 / NTP 开关 / NTP Server）。

    必须在用例修改任何设置之前调用（登录并进入 Date & Time 页之后）。
    """
    return {
        "tz_id": read_timezone_id(page),
        "ntp_enabled": is_ntp_enabled(page),
        "servers": read_ntp_servers(page),
    }


def restore_datetime_settings(page: Page, snap: dict) -> None:
    """把 Date & Time 状态恢复为执行前：还原时区 + NTP 配置，并用已知有效 NTP
    重新同步，确保设备时钟回到正确的当前时间，不污染后续用例。
    """
    nav_to_datetime(page)

    # 1) 还原时区（case01/case04 等可能改过时区，不还原会让后续用例「本机对比」整体偏移）
    if snap.get("tz_id"):
        select_timezone(page, snap["tz_id"])

    # 2) 用已知有效 NTP 把时钟拉回正确当前时间：启用 NTP + 写有效 server → Save → Sync
    set_ntp_enable(page, True)
    fill_ntp_servers(page, _KNOWN_GOOD_NTP)
    _click_save(page)
    sync_btn = page.get_by_role("button", name="Sync")
    if sync_btn.count() > 0:
        sync_btn.first.click()
        try:
            page.wait_for_selector("text=Synced", timeout=90000)
        except Exception:
            pass
        page.wait_for_timeout(500)

    # 3) 还原执行前的 NTP Server / 开关（时钟已在第 2 步同步正确，此步仅复原配置）
    nav_to_datetime(page)
    fill_ntp_servers(page, snap.get("servers") or ["", "", ""])
    set_ntp_enable(page, snap.get("ntp_enabled", True))
    _click_save(page)
