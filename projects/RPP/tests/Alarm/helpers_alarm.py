"""Alarm Config 用例组共享操作封装。

页面结构依据：knowledge/gateway/AcuHMI17/requirements/context/
Devices_Alarm_ActiveAlarm_context.md / Devices_Alarm_AlarmLogs_context.md /
SystemSettings_AlarmNotification_context.md + 2026-07-17 联机实测。

要点：
- 全局 Alarm 页与设备详情页的二级导航均为 el-menu--horizontal 的 menuitem；
- 设备 Alarm Config 的 Add Alarm 是内联表单页（URL 带 ?type=add），非弹窗；
- 行内主色按钮=编辑、危险色按钮=删除（图标按钮，无文本）。
"""
from __future__ import annotations

from playwright.sync_api import Locator, Page

from projects.RPP.tests.Alarm import config_alarm as cfg

# Devices 模块内的已知路由片段（判断当前是否在 Devices 侧）
_DEVICE_ROUTES = ("/#/dashboard", "/#/physicalDevices", "/#/virtualMeter",
                  "/#/webDevice", "/#/alarm", "/#/dataLog")

# 触发用参数：数值远离真实量测值，保证下一轮询必然越限（UNDERFLOW）。
# 多规则并发的用例按顺序取不同参数，避免同参数规则互相干扰。
TRIGGER_PARAMS = (
    ("System Frequency", "70", "90"),
    ("Phase A Line-to-Neutral Voltage", "500", "900"),
    ("Phase B Line-to-Neutral Voltage", "500", "900"),
)


def _wait(page: Page, extra_ms: int = 500) -> None:
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(extra_ms)


# ── 登录 ────────────────────────────────────────────────────────────────────

def login(page: Page) -> None:
    page.goto(cfg.BASE_URL + "/#/login")
    _wait(page)
    page.get_by_role("textbox", name="Enter User Name").fill(cfg.USERNAME)
    page.get_by_role("textbox", name="Enter Password").fill(cfg.PASSWORD)
    page.get_by_role("button", name="Sign In").click()
    _wait(page, 1000)
    # 关闭"修改默认密码"提醒弹窗（如有）
    try:
        page.get_by_role("button", name="Cancel").click(timeout=3000)
    except Exception:
        pass


def ensure_logged_in(page: Page) -> None:
    if "/#/login" in page.url or not page.url.startswith(cfg.BASE_URL):
        login(page)


# ── 导航 ────────────────────────────────────────────────────────────────────

def ensure_devices_module(page: Page) -> None:
    """确保当前位于 Devices 模块（左侧导航为 Dashboard~Data Log）。"""
    if not any(r in page.url for r in _DEVICE_ROUTES):
        page.locator("header span").filter(has_text="Devices").first.click()
        _wait(page)


def goto_global_alarm(page: Page, submenu: str | None = None) -> None:
    """进入 Devices → Alarm，submenu 可选 'Unacknowledged Alarms' / 'Alarm Logs'。"""
    ensure_devices_module(page)
    page.locator(".left-nav-item").filter(has_text="Alarm").click()
    _wait(page)
    if submenu:
        page.get_by_role("menuitem", name=submenu).first.click()
        _wait(page)


def goto_alarm_notification(page: Page) -> None:
    """进入 System Settings → Alarm Notification（直连 /#/systemSettings/alarm
    存在渲染竞态，稳定做法是先落 dateTime 再点顶部 menuitem）。"""
    page.goto(cfg.BASE_URL + "/#/systemSettings/dateTime")
    _wait(page)
    page.get_by_role("menuitem", name="Alarm Notification").click()
    _wait(page, 800)


def goto_device_alarm(page: Page, submenu: str,
                      device_name: str = cfg.TRIGGER_DEVICE) -> None:
    """进入 Physical Devices → 指定设备详情 → Alarm → Alarm Config / Alarm Logs。"""
    ensure_devices_module(page)
    page.locator(".left-nav-item").filter(has_text="Physical Devices").click()
    _wait(page)
    row = page.locator("tr.el-table__row").filter(has_text=device_name).first
    row.wait_for(timeout=10_000)
    row.locator("td").first.click()
    _wait(page, 800)
    item = page.get_by_role("menuitem", name=submenu)
    alarm_sub = page.locator("li.el-sub-menu").filter(has_text="Alarm")
    if alarm_sub.count() > 0 and (item.count() == 0 or not item.first.is_visible()):
        alarm_sub.first.click()
        page.wait_for_timeout(500)
        item = page.get_by_role("menuitem", name=submenu)
    item.first.click()
    _wait(page, 800)


# ── 表格读取 ─────────────────────────────────────────────────────────────────

def table_headers(page: Page) -> list[str]:
    """当前页面数据表的列头（限定 el-table 表头区，排除日期控件的星期表头）。"""
    ths = page.locator(".el-table__header-wrapper th")
    return [t for t in
            (ths.nth(i).inner_text().strip() for i in range(ths.count())) if t]


def data_rows(page: Page, text: str | None = None) -> Locator:
    rows = page.locator("tr.el-table__row")
    return rows.filter(has_text=text) if text else rows


def ack_status_of_row(row: Locator) -> str:
    """行内 Ack Status 单元格文本（Acknowledged / Unacknowledge），无则空串。"""
    for cell in row.locator("td").all():
        txt = cell.inner_text().strip()
        if "acknowledge" in txt.lower():
            return txt
    return ""


def count_rows_by_ack(page: Page, label: str, acknowledged: bool) -> int:
    """Alarm Logs 页中 label 对应、且确认状态匹配的行数。

    告警日志跨轮次累积，判断"本轮新告警/新确认"必须按状态计数，不能只看 label 存在。
    """
    rows = data_rows(page, label)
    n = 0
    for i in range(rows.count()):
        txt = ack_status_of_row(rows.nth(i)).lower()
        if not txt:
            continue
        is_unack = txt.startswith("unacknowl")
        if acknowledged and not is_unack:
            n += 1
        elif not acknowledged and is_unack:
            n += 1
    return n


def cell_text(page: Page, row: Locator, column: str) -> str:
    """按列名取行内单元格文本（依当前表头定位列序）。"""
    headers = table_headers(page)
    assert column in headers, f"当前表格无 {column!r} 列，实际列: {headers}"
    return row.locator("td").nth(headers.index(column)).inner_text().strip()


def wait_for_log_rows(page: Page, label: str, min_count: int = 1) -> bool:
    """在全局 Alarm Logs 页等待出现 label 相关记录（不区分确认状态）。"""
    goto_global_alarm(page, "Alarm Logs")
    for _ in range(cfg.POLL_ROUNDS):
        if data_rows(page, label).count() >= min_count:
            return True
        page.wait_for_timeout(cfg.POLL_STEP_MS)
        page.reload()
        _wait(page, 800)
    return data_rows(page, label).count() >= min_count


def unack_total(page: Page) -> int:
    """全局 Unacknowledged Alarms 页的未确认告警总数（优先分页 Total，回退行数）。"""
    total = page.locator(".el-pagination__total")
    if total.count() > 0:
        digits = "".join(ch for ch in total.first.inner_text() if ch.isdigit())
        if digits:
            return int(digits)
    return data_rows(page).count()


def rule_alarm_on(page: Page, label: str) -> bool:
    """设备 Alarm Config 页中该规则的告警状态（Status 列为图标无文本：
    ON=el-icon warning，正常=el-icon success，未轮询='-'）。"""
    row = data_rows(page, label).first
    headers = table_headers(page)
    assert "Status" in headers, f"Alarm Config 无 Status 列，实际列: {headers}"
    html = row.locator("td").nth(headers.index("Status")).evaluate(
        "el => el.innerHTML") or ""
    return "warning" in html


def unack_tab_present(page: Page) -> bool:
    """全局 Alarm 页是否存在 Unacknowledged Alarms 二级 tab
    （Alarm Acknowledgement Enable=Disable 时该 tab 整体隐藏）。"""
    goto_global_alarm(page)
    return page.get_by_role(
        "menuitem", name="Unacknowledged Alarms").count() > 0


# ── Alarm Acknowledgement Enable 开关 ────────────────────────────────────────

def _ack_radio(page: Page, enable: bool) -> Locator:
    fi = page.locator(".el-form-item").filter(
        has_text="Alarm Acknowledgement Enable").first
    return fi.locator(".el-radio").filter(
        has_text="Enable" if enable else "Disable").first


def get_ack_enable(page: Page) -> bool:
    goto_alarm_notification(page)
    radio = _ack_radio(page, True)
    return "is-checked" in (radio.get_attribute("class") or "")


def set_ack_enable(page: Page, enable: bool) -> None:
    """设置 Alarm Acknowledgement Enable 并保存（已是目标状态则不动）。"""
    goto_alarm_notification(page)
    radio = _ack_radio(page, enable)
    if "is-checked" in (radio.get_attribute("class") or ""):
        return
    radio.click()
    page.wait_for_timeout(300)
    # Alarm Email Enable=Enable 且收件人为空时 Save 会被必填校验拦下，
    # 先把 Email 关掉（测试环境不依赖告警邮件）。
    email_fi = page.locator(".el-form-item").filter(
        has_text="Alarm Email Enable").first
    email_enable = email_fi.locator(".el-radio").filter(has_text="Enable").first
    recipient = page.locator("[placeholder='Enter Recipient 1']")
    if ("is-checked" in (email_enable.get_attribute("class") or "")
            and recipient.count() > 0 and not recipient.first.input_value()):
        email_fi.locator(".el-radio").filter(has_text="Disable").first.click()
        page.wait_for_timeout(300)
    page.get_by_role("button", name="Save").click()
    _wait(page, 800)
    assert page.locator(".el-message--error").count() == 0, \
        "保存 Alarm Acknowledgement Enable 配置失败（出现错误提示）"


# ── 告警规则 CRUD（设备 Alarm Config 页） ────────────────────────────────────

def add_alarm_rule(page: Page, label: str, param: str = "System Frequency",
                   vmin: str = "70", vmax: str = "90") -> None:
    """在当前设备 Alarm Config 页新增一条告警规则（内联表单页）。"""
    page.get_by_role("button", name="Add Alarm").click()
    _wait(page, 600)
    page.locator("[placeholder='Enter Label']").fill(label)
    param_fi = page.locator(".el-form-item").filter(has_text="Parameter").first
    param_fi.locator(".el-select").click()
    page.wait_for_timeout(400)
    page.locator("li.el-select-dropdown__item").filter(
        has_text=param).first.click()
    page.wait_for_timeout(300)
    page.locator("[placeholder='Enter Min Value']").fill(vmin)
    page.locator("[placeholder='Enter Max Value']").fill(vmax)
    page.get_by_role("button", name="Save").click()
    _wait(page, 800)
    assert data_rows(page, label).count() > 0 or \
        page.locator(".el-message--success").count() > 0, \
        f"新增告警规则 {label!r} 后列表中未出现该条目"


def edit_alarm_rule_range(page: Page, label: str, vmin: str, vmax: str) -> None:
    """把已有规则阈值改为 [vmin, vmax]（用于制造/消除告警条件）。"""
    row = data_rows(page, label).first
    row.wait_for(timeout=10_000)
    row.locator("button.el-button--primary").first.click()
    _wait(page, 600)
    page.locator("[placeholder='Enter Min Value']").fill(vmin)
    page.locator("[placeholder='Enter Max Value']").fill(vmax)
    page.get_by_role("button", name="Save").click()
    _wait(page, 800)


def _confirm_dialog(page: Page) -> None:
    """点击二次确认弹窗的确认按钮；无弹窗时静默返回。

    只点可见按钮：页面上常驻隐藏的日期选择面板（el-picker-panel）里也有
    OK 按钮，按文本过滤会误匹配到它导致 click 一直等待。
    """
    for txt in ("OK", "Confirm", "Yes", "确认"):
        btns = page.locator(
            ".el-message-box button, .el-dialog button, .el-popconfirm button"
        ).filter(has_text=txt)
        for i in range(btns.count()):
            btn = btns.nth(i)
            if btn.is_visible():
                btn.click()
                _wait(page)
                return
    # popconfirm 主按钮兜底
    fallback = page.locator(".el-popconfirm__action .el-button--primary")
    if fallback.count() > 0 and fallback.first.is_visible():
        fallback.first.click()
        _wait(page)


def delete_alarm_rule(page: Page, label: str) -> None:
    rows = data_rows(page, label)
    while rows.count() > 0:
        rows.first.locator("button.el-button--danger").first.click()
        page.wait_for_timeout(400)
        _confirm_dialog(page)
        page.wait_for_timeout(600)


def cleanup_test_rules(page: Page, device_name: str = cfg.TRIGGER_DEVICE) -> None:
    """删除本模块创建的全部规则（label 以 RULE_PREFIX 开头），幂等。"""
    goto_device_alarm(page, "Alarm Config", device_name)
    rows = data_rows(page, cfg.RULE_PREFIX)
    while rows.count() > 0:
        rows.first.locator("button.el-button--danger").first.click()
        page.wait_for_timeout(400)
        _confirm_dialog(page)
        page.wait_for_timeout(600)


# ── 触发 / 确认 / 等待 ───────────────────────────────────────────────────────

def trigger_alarm(page: Page, label: str, param_idx: int = 0,
                  device_name: str = cfg.TRIGGER_DEVICE) -> None:
    """创建必触发规则：先删同名旧规则保证 OFF→ON 状态翻转，产生新告警。"""
    param, vmin, vmax = TRIGGER_PARAMS[param_idx]
    goto_device_alarm(page, "Alarm Config", device_name)
    delete_alarm_rule(page, label)
    add_alarm_rule(page, label, param, vmin, vmax)


def wait_for_unack(page: Page, label: str) -> bool:
    """在全局 Unacknowledged Alarms 页等待 label 告警出现（约一个轮询周期内）。"""
    goto_global_alarm(page, "Unacknowledged Alarms")
    for _ in range(cfg.POLL_ROUNDS):
        if data_rows(page, label).count() > 0:
            return True
        page.wait_for_timeout(cfg.POLL_STEP_MS)
        page.reload()
        _wait(page, 800)
    return data_rows(page, label).count() > 0


def wait_for_log_status(page: Page, label: str, acknowledged: bool,
                        min_count: int = 1) -> bool:
    """在全局 Alarm Logs 页等待 label 出现指定确认状态的记录。"""
    goto_global_alarm(page, "Alarm Logs")
    for _ in range(cfg.POLL_ROUNDS):
        if count_rows_by_ack(page, label, acknowledged) >= min_count:
            return True
        page.wait_for_timeout(cfg.POLL_STEP_MS)
        page.reload()
        _wait(page, 800)
    return count_rows_by_ack(page, label, acknowledged) >= min_count


def ack_alarm(page: Page, label: str) -> None:
    """在全局 Unacknowledged Alarms 页确认 label 对应的第一条告警。"""
    goto_global_alarm(page, "Unacknowledged Alarms")
    row = data_rows(page, label).first
    row.wait_for(timeout=10_000)
    row.get_by_role("button", name="Acknowledge").click()
    page.wait_for_timeout(400)
    _confirm_dialog(page)
    _wait(page, 800)


# ── Alarm Logs 过滤区（Interval / Serial Number / Monitor ID / Search / Reset）──

def set_interval_filter(page: Page, which: str) -> None:
    """通过日期面板选择 Interval 区间（datetimerange）。

    which='today'  → [今天 00:00:00, 明天 00:00:00]，覆盖当天全部记录；
    which='future' → [下月 1 日, 下月 2 日]，未来区间必然检索为空（负向用）。

    实测该控件直接向输入框键入文本不会同步组件内部值（面板关闭即回滚、
    Search 按空条件执行），必须通过面板点选日期 + footer OK 确认。
    """
    page.locator("input.el-range-input").nth(0).click()
    page.wait_for_timeout(500)
    panel = page.locator(".el-picker-panel:visible").first
    left = panel.locator(".el-picker-panel__content.is-left")
    right = panel.locator(".el-picker-panel__content.is-right")
    if which == "today":
        today_cell = left.locator("td.available.today").first
        day = int(today_cell.inner_text().strip())
        today_cell.click()
        page.wait_for_timeout(300)
        # 明天：同月下一天；月末则取右面板（下月）1 日
        tomorrow = left.locator("td.available").get_by_text(
            str(day + 1), exact=True)
        if tomorrow.count() > 0:
            tomorrow.first.click()
        else:
            right.locator("td.available").get_by_text(
                "1", exact=True).first.click()
    elif which == "future":
        right.locator("td.available").get_by_text("1", exact=True).first.click()
        page.wait_for_timeout(300)
        right.locator("td.available").get_by_text("2", exact=True).first.click()
    else:
        raise ValueError(f"未知区间类型: {which!r}")
    page.wait_for_timeout(300)
    ok_btn = panel.locator(".el-picker-panel__footer button").last
    ok_btn.click()
    page.wait_for_timeout(400)
    assert page.locator("input.el-range-input").nth(0).input_value(), \
        "Interval 面板确认后 Start Date 输入框应有值"


def select_serial_filter(page: Page, serial: str) -> None:
    """选择 Serial Number 下拉（过滤区唯一可见 el-select）。"""
    page.locator(".el-select:visible").first.click()
    page.wait_for_timeout(400)
    page.locator("li.el-select-dropdown__item").filter(
        has_text=serial).first.click()
    page.wait_for_timeout(300)


def fill_monitor_id_filter(page: Page, monitor_id: str) -> None:
    page.locator("[placeholder='Enter Monitor ID']").fill(monitor_id)


def click_search(page: Page) -> None:
    page.get_by_role("button", name="Search").click()
    _wait(page, 800)


def click_reset(page: Page) -> None:
    page.get_by_role("button", name="Reset").click()
    _wait(page, 800)


def column_values(page: Page, column: str) -> list[str]:
    """当前页所有数据行指定列的文本。"""
    headers = table_headers(page)
    assert column in headers, f"当前表格无 {column!r} 列，实际列: {headers}"
    idx = headers.index(column)
    rows = data_rows(page)
    return [rows.nth(i).locator("td").nth(idx).inner_text().strip()
            for i in range(rows.count())]


def ensure_alarm_log_data(page: Page) -> dict:
    """确保 Alarm Logs 中存在本模块的目标记录（label=at_alog1）且总行数≥2。

    记录不存在时现场触发一次（约一个轮询周期），随后确认告警并删除规则——
    日志记录会保留，后续检索用例直接复用，避免每条用例重复等待。
    返回 {'label','serial','monitor_id'}，serial/monitor_id 从日志行实时读取
    （规则重建后 Monitor ID 会变化，不能写死）。
    """
    label = cfg.RULE_PREFIX + "alog1"
    goto_global_alarm(page, "Alarm Logs")
    if data_rows(page, label).count() == 0:
        trigger_alarm(page, label)
        assert wait_for_log_rows(page, label), \
            f"触发规则 {label!r} 后，Alarm Logs 未在轮询周期内出现记录"
        try:
            ack_all_alarms(page)
        except Exception:
            pass
        cleanup_test_rules(page)
        goto_global_alarm(page, "Alarm Logs")
    assert data_rows(page).count() >= 2, \
        "前置条件不满足：Alarm Logs 中应至少有 2 条告警记录"
    row = data_rows(page, label).first
    return {
        "label": label,
        "serial": cell_text(page, row, "Serial Number"),
        "monitor_id": cell_text(page, row, "Monitor ID"),
    }


def ack_all_alarms(page: Page) -> None:
    """全局 Unacknowledged Alarms 页一键确认全部（列表为空时不操作）。"""
    goto_global_alarm(page, "Unacknowledged Alarms")
    if data_rows(page).count() == 0:
        return
    page.get_by_role("button", name="Ack All Alarms").click()
    page.wait_for_timeout(400)
    _confirm_dialog(page)
    _wait(page, 800)
